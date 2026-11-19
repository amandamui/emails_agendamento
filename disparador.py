"""Fase 2: cobra TODOS os fornecedores, cruzando o consolidado com a base de contatos.

Uso:
    python disparador.py                      # SIMULAÇÃO: mostra quem receberia, não envia nada
    python disparador.py --enviar             # envia de verdade pelo Outlook
    python disparador.py 2026-11-24 [--enviar]  # usa outra data de referência
"""
import re
import sys
import unicodedata
from datetime import datetime

import pandas as pd
from rapidfuzz import fuzz, process

from config import PASTA_PROJETO
from teste_email import enviar_email, ler_excel, localizar_consolidado, notas_a_cobrar

# =========================================================
# ⚙️ CONFIGURAÇÕES
# =========================================================
ARQUIVO_CONTATOS = PASTA_PROJETO / "base fornecedores" / "baseagdteste.xlsx"
ARQUIVO_PENDENTES = PASTA_PROJETO / "atrasadas" / "Fornecedores_Pendentes.xlsx"
ARQUIVO_LOG = PASTA_PROJETO / "atrasadas" / "Log_Envios.csv"

# Nota mínima (0-100) para considerar que a Razão Social da nota é o fornecedor da base.
# Comparação feita SEM os termos genéricos abaixo; abaixo da nota, a nota vai para os pendentes.
SIMILARIDADE_MINIMA = 90
TERMOS_GENERICOS = {
    "LTDA", "EIRELI", "SA", "S", "A", "ME", "EPP", "IND", "INDS", "INDUSTRIA", "COM", "COMERCIO",
    "DO", "DA", "DE", "DOS", "DAS", "E", "BRASIL", "BRAZIL", "AUTOMOTIVE", "AUTOMOTIVA", "AUTOMOTIVAS",
    "AUTOM", "PROD", "PRODUTOS", "PECAS", "AUTOPECAS",
}

# =========================================================
# 🔗 CRUZAMENTO NOTAS × CONTATOS
# =========================================================


def nome_chave(razao_social) -> str:
    """'R C Borrachas Automotivas Ltda.' -> 'RC BORRACHAS' (sem acento, pontuação e termos genéricos)."""
    texto = unicodedata.normalize("NFKD", str(razao_social)).encode("ascii", "ignore").decode().upper()
    texto = re.sub(r"\b([A-Z]) (?=[A-Z]\b)", r"\1", texto)  # junta siglas soltas: "R C" -> "RC"
    return " ".join(t for t in re.findall(r"[A-Z0-9]+", texto) if t not in TERMOS_GENERICOS)


def carregar_contatos() -> pd.DataFrame | None:
    if not ARQUIVO_CONTATOS.exists():
        print(f"❌ Base de contatos não encontrada: {ARQUIVO_CONTATOS}")
        return None
    print("📂 Lendo base de fornecedores...")
    contatos = ler_excel(ARQUIVO_CONTATOS).dropna(subset=["RAZÃO SOCIAL", "EMAIL DE CONTATO"])
    contatos["Chave"] = contatos["RAZÃO SOCIAL"].map(nome_chave)
    # Junta todos os e-mails do fornecedor (vários na mesma célula e/ou em várias linhas), sem repetir.
    # Outlook separa destinatários com ';'.
    def juntar_emails(celulas: pd.Series) -> str:
        emails = (e.strip().lower() for c in celulas.astype(str) for e in re.split(r"[,;\s]+", c))
        return ";".join(dict.fromkeys(e for e in emails if "@" in e))

    return contatos.groupby("Chave", as_index=False).agg(
        {"RAZÃO SOCIAL": "first", "EMAIL DE CONTATO": juntar_emails})


def cruzar(notas: pd.DataFrame, contatos: pd.DataFrame) -> pd.DataFrame:
    """Adiciona 'Fornecedor Base', 'EMAIL DE CONTATO' e 'Similaridade' (match feito uma vez por fornecedor)."""
    base = contatos.set_index("Chave")
    chaves_base = list(base.index)
    chave_por_razao, nota_por_razao = {}, {}
    for razao in notas["Razão Social"].dropna().unique():
        achado = process.extractOne(nome_chave(razao), chaves_base, scorer=fuzz.token_sort_ratio,
                                    score_cutoff=SIMILARIDADE_MINIMA)
        if achado:
            chave_por_razao[razao], nota_por_razao[razao] = achado[0], round(achado[1], 1)

    notas = notas.copy()
    chave = notas["Razão Social"].map(chave_por_razao)
    notas["Fornecedor Base"] = chave.map(base["RAZÃO SOCIAL"])
    notas["EMAIL DE CONTATO"] = chave.map(base["EMAIL DE CONTATO"])
    notas["Similaridade"] = notas["Razão Social"].map(nota_por_razao)
    return notas


# =========================================================
# 📧 ENVIO / RELATÓRIOS
# =========================================================


def disparar(df_enviar: pd.DataFrame, enviar: bool) -> list[dict]:
    logs = []
    for (fornecedor, email), grupo in df_enviar.groupby(["Fornecedor Base", "EMAIL DE CONTATO"]):
        print(f"   ✉️ {fornecedor} ({email}) - {len(grupo)} notas")
        if not enviar:
            continue
        ok = enviar_email(grupo.drop(columns=["Fornecedor Base", "EMAIL DE CONTATO", "Similaridade"]),
                          str(fornecedor), destinatario=str(email))
        logs.append({
            "Data/Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Fornecedor": fornecedor,
            "E-mail": email,
            "Qtd Notas": len(grupo),
            "Status": "Sucesso" if ok else "Erro",
        })
    return logs


def salvar_relatorios(df_pendentes: pd.DataFrame, logs: list[dict]):
    if df_pendentes.empty:
        print("\n✅ Todas as notas a cobrar têm e-mail cadastrado.")
    else:
        try:
            df_pendentes.drop(columns=["Fornecedor Base", "EMAIL DE CONTATO", "Similaridade"]).to_excel(
                ARQUIVO_PENDENTES, index=False)
            print(f"\n⚠️ {len(df_pendentes)} notas de {df_pendentes['Razão Social'].nunique()} fornecedores "
                  f"sem e-mail cadastrado. Lista em: {ARQUIVO_PENDENTES.name}")
        except PermissionError:
            print(f"\n❌ Feche o {ARQUIVO_PENDENTES.name} no Excel para salvar os pendentes.")

    if logs:
        pd.DataFrame(logs).to_csv(ARQUIVO_LOG, index=False, sep=";", encoding="utf-8-sig",
                                  mode="a", header=not ARQUIVO_LOG.exists())  # acumula o histórico
        print(f"📊 Log de envios atualizado: {ARQUIVO_LOG.name}")


# =========================================================
# 🏁 EXECUÇÃO
# =========================================================


def run(data_referencia: pd.Timestamp | None = None, enviar: bool = False):
    arquivo = localizar_consolidado()
    if arquivo is None:
        print("❌ Nenhum consolidado encontrado. Rode antes o extrator.py ou o atrasadas.py.")
        return
    contatos = carregar_contatos()
    if contatos is None:
        return

    print(f"📂 Lendo notas: {arquivo.name}")
    notas = notas_a_cobrar(ler_excel(arquivo), data_referencia)
    if notas.empty:
        print("✅ Nenhuma nota atende aos critérios de cobrança.")
        return

    notas = cruzar(notas, contatos)
    tem_email = notas["EMAIL DE CONTATO"].notna()
    df_enviar, df_pendentes = notas[tem_email], notas[~tem_email]

    print(f"\n{'📧 Enviando' if enviar else '🧪 SIMULAÇÃO (nada será enviado)'}: "
          f"{df_enviar['Fornecedor Base'].nunique()} fornecedores, {len(df_enviar)} notas")
    logs = disparar(df_enviar, enviar)
    salvar_relatorios(df_pendentes, logs)
    if not enviar:
        print("\nℹ️ Para enviar de verdade: python disparador.py --enviar")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--enviar"]
    run(pd.Timestamp(args[0]) if args else None, enviar="--enviar" in sys.argv)
