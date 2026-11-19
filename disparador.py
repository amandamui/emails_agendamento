"""Fase 2: cobra TODOS os fornecedores, cruzando o consolidado com a base de contatos.

Uso:
    python disparador.py                      # SIMULAÇÃO: mostra quem receberia, não envia nada
    python disparador.py --teste              # envia TODOS os e-mails para o EMAIL_TESTE do .env
    python disparador.py --enviar             # envia de verdade para os fornecedores
    python disparador.py 2026-11-24 [--teste | --enviar]  # usa outra data de referência
"""
import re
import sys
import unicodedata
from datetime import datetime

import pandas as pd
from rapidfuzz import fuzz, process

from config import PASTA_CONTATOS, PASTA_ENVIOS
from teste_email import EMAIL_DESTINATARIO, enviar_email, ler_excel, localizar_consolidado, notas_a_cobrar

# =========================================================
# ⚙️ CONFIGURAÇÕES
# =========================================================
ARQUIVO_CONTATOS = PASTA_CONTATOS / "baseagdteste.xlsx"
ARQUIVO_PENDENTES = PASTA_ENVIOS / "Fornecedores_Pendentes.xlsx"
ARQUIVO_LOG = PASTA_ENVIOS / "Log_Envios.csv"

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


def disparar(df_enviar: pd.DataFrame, modo: str) -> list[dict]:
    """modo: 'simulacao' (só lista), 'teste' (tudo para o EMAIL_TESTE) ou 'enviar' (fornecedores)."""
    logs = []
    for (fornecedor, email), grupo in df_enviar.groupby(["Fornecedor Base", "EMAIL DE CONTATO"]):
        print(f"   ✉️ {fornecedor} ({email}) - {len(grupo)} notas")
        if modo == "simulacao":
            continue
        teste = modo == "teste"
        destino = EMAIL_DESTINATARIO if teste else str(email)
        ok = enviar_email(grupo.drop(columns=["Fornecedor Base", "EMAIL DE CONTATO", "Similaridade"]),
                          str(fornecedor), destinatario=destino,
                          prefixo_assunto=f"[TESTE → {email}] " if teste else "")
        logs.append({
            "Data/Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Fornecedor": fornecedor,
            "E-mail": destino,
            "Qtd Notas": len(grupo),
            "Status": ("Sucesso" if ok else "Erro") + (" (teste)" if teste else ""),
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
        df_log = pd.DataFrame(logs)
        try:
            df_log.to_csv(ARQUIVO_LOG, index=False, sep=";", encoding="utf-8-sig",
                          mode="a", header=not ARQUIVO_LOG.exists())  # acumula o histórico
            print(f"📊 Log de envios atualizado: {ARQUIVO_LOG}")
        except PermissionError:
            # Log aberto no Excel: os e-mails já saíram, então o registro não pode se perder.
            reserva = ARQUIVO_LOG.with_name(f"Log_Envios_{datetime.now():%Y%m%d_%H%M%S}.csv")
            df_log.to_csv(reserva, index=False, sep=";", encoding="utf-8-sig")
            print(f"⚠️ {ARQUIVO_LOG.name} estava aberto no Excel; envios registrados em: {reserva}")
            print(f"   Copie essas linhas para o {ARQUIVO_LOG.name} quando fechá-lo.")


# =========================================================
# 🏁 EXECUÇÃO
# =========================================================


def preparar_consolidado(extrair: bool = False, consolidar: bool = False):
    """Garante um consolidado: extrai do portal e/ou consolida o que faltar (imports tardios: sem ciclo)."""
    if extrair:
        import extrator
        extrator.baixar_planilhas()
    if extrair or consolidar or localizar_consolidado() is None:
        if not (extrair or consolidar):
            print("⚠️ Nenhum consolidado encontrado; consolidando as planilhas antes do disparo…\n")
        import atrasadas
        atrasadas.consolidar(extrair_se_vazio=not extrair)  # 1_downloads vazia → extrai do portal (uma vez só)
        print()
    return localizar_consolidado()


def run(data_referencia: pd.Timestamp | None = None, modo: str = "simulacao",
        extrair: bool = False, consolidar: bool = False):
    arquivo = preparar_consolidado(extrair, consolidar)
    if arquivo is None:
        print("❌ Não foi possível gerar o consolidado; nada a disparar.")
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

    titulo = {"simulacao": "🧪 SIMULAÇÃO (nada será enviado)",
              "teste": f"🧪 TESTE (todos os e-mails vão para {EMAIL_DESTINATARIO})",
              "enviar": "📧 ENVIANDO PARA OS FORNECEDORES"}[modo]
    print(f"\n{titulo}: {df_enviar['Fornecedor Base'].nunique()} fornecedores, {len(df_enviar)} notas")
    logs = disparar(df_enviar, modo)
    salvar_relatorios(df_pendentes, logs)
    if modo == "simulacao":
        print("\nℹ️ Teste no seu e-mail: python disparador.py --teste | Envio real: python disparador.py --enviar")


def main(extrair: bool = False, consolidar: bool = False):
    """Ponto de entrada comum de extrator.py, atrasadas.py e disparador.py: [AAAA-MM-DD] [--teste | --enviar]."""
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if flags - {"--teste", "--enviar"} or len(flags) > 1:
        raise SystemExit("Uso: python <script>.py [AAAA-MM-DD] [--teste | --enviar]")
    datas = [a for a in sys.argv[1:] if not a.startswith("--")]
    run(pd.Timestamp(datas[0]) if datas else None,
        modo="teste" if "--teste" in flags else "enviar" if "--enviar" in flags else "simulacao",
        extrair=extrair, consolidar=consolidar)


if __name__ == "__main__":
    main()
