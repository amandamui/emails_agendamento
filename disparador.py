"""Fase 2: cobra TODOS os fornecedores, cruzando o consolidado com a base de contatos.

Uso:
    python disparador.py                      # SIMULAÇÃO: mostra quem receberia, não envia nada
    python disparador.py --teste              # envia TODOS os e-mails para o EMAIL_TESTE do .env
    python disparador.py --enviar             # envia de verdade para os fornecedores
    python disparador.py 2026-11-24 [--teste | --enviar]  # usa outra data de referência
"""
import re
import sys
from datetime import datetime

import pandas as pd

from config import PASTA_CONTATOS, PASTA_ENVIOS, PASTA_HISTORICO
from teste_email import EMAIL_DESTINATARIO, enviar_email, ler_excel, localizar_consolidado, notas_a_cobrar

# =========================================================
# ⚙️ CONFIGURAÇÕES
# =========================================================
ARQUIVO_CONTATOS = PASTA_CONTATOS / "baseagdteste.xlsx"
ARQUIVO_PENDENTES = PASTA_ENVIOS / "Fornecedores_Pendentes.xlsx"
ARQUIVO_LOG = PASTA_ENVIOS / "Log_Envios.csv"

COLUNA_CNPJ_NOTA = "Emitente"
COLUNA_CNPJ_BASE = "CNPJ FORNECEDOR"
COLUNAS_CRUZAMENTO = ["Fornecedor Base", "EMAIL DE CONTATO", "Cruzamento"]
EMAIL_VALIDO = re.compile(r"[a-z0-9._%+'-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}")

# =========================================================
# 🔗 CRUZAMENTO NOTAS × CONTATOS (pelo CNPJ)
# =========================================================


def cnpj_digitos(valor) -> str | None:
    """'02.162.259/0003-26', 2162259000326 ou '2162259000326' -> '02162259000326' (14 dígitos)."""
    if pd.isna(valor):
        return None
    if isinstance(valor, float):
        valor = int(valor)
    digitos = re.sub(r"\D", "", str(valor))
    return digitos.zfill(14) if 0 < len(digitos) <= 14 else None


def carregar_contatos() -> pd.DataFrame | None:
    if not ARQUIVO_CONTATOS.exists():
        print(f"❌ Base de contatos não encontrada: {ARQUIVO_CONTATOS}")
        return None
    print("📂 Lendo base de fornecedores...")
    contatos = ler_excel(ARQUIVO_CONTATOS).dropna(subset=["RAZÃO SOCIAL", "EMAIL DE CONTATO"])
    contatos["CNPJ"] = contatos[COLUNA_CNPJ_BASE].map(cnpj_digitos)
    sem_cnpj = contatos["CNPJ"].isna()
    if sem_cnpj.any():
        print(f"⚠️ {sem_cnpj.sum()} linha(s) da base sem CNPJ válido (ignoradas): "
              + ", ".join(contatos.loc[sem_cnpj, "RAZÃO SOCIAL"].astype(str)))
    contatos = contatos[~sem_cnpj].copy()
    contatos["Raiz"] = contatos["CNPJ"].str[:8]
    contatos["Linha Excel"] = contatos.index + 2  # cabeçalho na linha 1
    contatos["Nome"] = contatos["RAZÃO SOCIAL"].astype(str).str.strip().str.upper()
    return contatos


def _agrupar(contatos: pd.DataFrame, chave: str) -> pd.DataFrame:
    """Um registro por chave (CNPJ ou raiz) com TODOS os e-mails juntos (células e linhas), sem repetir.
    Outlook separa destinatários com ';'. Chave com Razões Sociais diferentes na base é AMBÍGUA
    (ex.: CNPJ de um fornecedor copiado na linha de outro) e fica de fora: a nota vai para os pendentes."""
    def pedacos(celulas: pd.Series) -> list[str]:
        """Separa por ',', ';', '/', espaço ou quebra de linha; descarta o que não tem '@' (nomes, '-')."""
        emails = (e.strip().lower() for c in celulas.astype(str) for e in re.split(r"[,;/\s]+", c))
        return list(dict.fromkeys(e for e in emails if "@" in e))

    grupos = contatos.groupby(chave).agg({"RAZÃO SOCIAL": "first", "Nome": "nunique", "Linha Excel": list})
    emails = contatos.groupby(chave)["EMAIL DE CONTATO"].agg(pedacos)

    def juntar_validos(lista: object) -> str | None:
        if not isinstance(lista, list):
            raise TypeError("A lista de e-mails agrupados deve ser uma lista.")
        return ";".join(e for e in lista if EMAIL_VALIDO.fullmatch(e)) or None

    def listar_invalidos(lista: object) -> list[str]:
        if not isinstance(lista, list):
            raise TypeError("A lista de e-mails agrupados deve ser uma lista.")
        return [e for e in lista if not EMAIL_VALIDO.fullmatch(e)]

    grupos["EMAIL DE CONTATO"] = emails.map(juntar_validos)
    # E-mails grudados ('a@x.com.brb@y.com.br') não têm separação segura: não adivinhar, segurar a nota.
    grupos["E-mails Inválidos"] = emails.map(listar_invalidos)
    grupos["Ambíguo"] = grupos["Nome"] > 1
    return grupos


def cruzar(notas: pd.DataFrame, contatos: pd.DataFrame) -> pd.DataFrame:
    """Adiciona 'Fornecedor Base', 'EMAIL DE CONTATO' e 'Cruzamento' pelo CNPJ do Emitente:
    1º CNPJ exato; 2º raiz (8 primeiros dígitos = mesma empresa, outra filial).
    Sem CNPJ na base, ou CNPJ/raiz ambíguo (Razões Sociais diferentes) → pendente."""
    por_cnpj, por_raiz = _agrupar(contatos, "CNPJ"), _agrupar(contatos, "Raiz")
    notas = notas.copy()
    cnpj = notas[COLUNA_CNPJ_NOTA].map(cnpj_digitos)
    raiz = cnpj.str[:8]
    na_base = cnpj.isin(por_cnpj.index)
    exato = na_base & ~cnpj.map(por_cnpj["Ambíguo"]).fillna(False).astype(bool)
    pela_raiz = ~na_base & raiz.isin(por_raiz.index) & ~raiz.map(por_raiz["Ambíguo"]).fillna(False).astype(bool)

    # Avisa só dos conflitos que de fato seguraram notas
    conflitos = pd.concat([por_cnpj.loc[por_cnpj.index.isin(cnpj[na_base & ~exato])],
                           por_raiz.loc[por_raiz.index.isin(raiz[~na_base & ~pela_raiz]) & por_raiz["Ambíguo"]]])
    for chave, g in conflitos.iterrows():
        print(f"⚠️ CNPJ {chave} aparece na base com Razões Sociais diferentes (linhas {g['Linha Excel']}); "
              "notas ficaram pendentes até corrigir a base.")

    # E-mail inválido (grudado, sem domínio...) ou nenhum e-mail válido → pendente, nunca envio parcial/errado
    def do_grupo(coluna: str) -> pd.Series:
        return cnpj.map(por_cnpj[coluna]).where(exato, raiz.map(por_raiz[coluna]))

    # E-mail inválido (grudado, sem domínio...) ou nenhum válido → pendente, nunca envio parcial ou errado
    invalidos = do_grupo("E-mails Inválidos").map(lambda x: ", ".join(x) if isinstance(x, list) else "")
    problema = (exato | pela_raiz) & ((invalidos != "") | do_grupo("EMAIL DE CONTATO").isna())
    if problema.any():
        avisos = pd.DataFrame({"Fornecedor": notas["Razão Social"], "Inválidos": invalidos,
                               "Linhas": do_grupo("Linha Excel").map(str)})[problema]
        for (nome, ruins, linhas), g in avisos.groupby(["Fornecedor", "Inválidos", "Linhas"]):
            motivo = f"e-mail inválido na base ({ruins})" if ruins else "nenhum e-mail válido na base"
            print(f"⚠️ {nome}: {motivo}, linhas {linhas}; {len(g)} nota(s) pendentes até corrigir a base.")
    exato, pela_raiz = exato & ~problema, pela_raiz & ~problema

    for coluna, destino in (("RAZÃO SOCIAL", "Fornecedor Base"), ("EMAIL DE CONTATO", "EMAIL DE CONTATO")):
        notas[destino] = do_grupo(coluna).where(exato | pela_raiz)  # o resto fica vazio → pendente
    notas["Cruzamento"] = None
    notas.loc[exato, "Cruzamento"] = "CNPJ"
    notas.loc[pela_raiz, "Cruzamento"] = "Raiz do CNPJ"
    return notas


# =========================================================
# 📧 ENVIO / RELATÓRIOS
# =========================================================


def disparar(df_enviar: pd.DataFrame, modo: str) -> list[dict]:
    """modo: 'simulacao' (só lista), 'teste' (tudo para o EMAIL_TESTE) ou 'enviar' (fornecedores)."""
    logs = []
    grupos = sorted(((g["Razão Social"].mode().iloc[0], email, g)
                     for (_, email), g in df_enviar.groupby(["Fornecedor Base", "EMAIL DE CONTATO"])),
                    key=lambda item: item[0].upper())
    for fornecedor, email, grupo in grupos:
        print(f"   ✉️   {fornecedor} ({email}) - {len(grupo)} notas")
        if modo == "simulacao":
            continue
        teste = modo == "teste"
        destino = EMAIL_DESTINATARIO if teste else str(email)
        ok = enviar_email(grupo.drop(columns=COLUNAS_CRUZAMENTO),
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
            df_pendentes.drop(columns=COLUNAS_CRUZAMENTO).to_excel(ARQUIVO_PENDENTES, index=False)
            print(f"\n⚠️  {len(df_pendentes)} notas de {df_pendentes[COLUNA_CNPJ_NOTA].nunique()} CNPJs "
                  f"sem cadastro (ou com conflito/e-mail inválido) na base (CNPJ na coluna '{COLUNA_CNPJ_NOTA}'). Lista em: {ARQUIVO_PENDENTES.name}")
        except PermissionError:
            print(f"\n❌ Feche o {ARQUIVO_PENDENTES.name} no Excel para salvar os pendentes.")

    registrar_log(logs)


def registrar_log(logs: list[dict]):
    """Acumula os envios no Log_Envios.csv (também usado pela devolutiva)."""
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


def salvar_historico(consolidado: pd.DataFrame):
    """Foto do consolidado em 5_historico/Atrasadas_AAAA-MM-DD.xlsx, base da devolutiva semanal.
    A data é a do download das planilhas (não a de hoje): com o cache, a foto não finge ser de um dia
    em que o portal não foi lido. Mesma data = sobrescreve (uma foto por dia)."""
    import atrasadas
    planilhas = atrasadas.planilhas_baixadas()
    data = (datetime.fromtimestamp(max(p.stat().st_mtime for p in planilhas)) if planilhas
            else datetime.now())
    destino = PASTA_HISTORICO / f"Atrasadas_{data:%Y-%m-%d}.xlsx"
    try:
        consolidado.to_excel(destino, index=False)
        print(f"🗂️ Foto do dia salva no histórico: {destino.name}")
    except PermissionError:
        print(f"⚠️ Feche o {destino.name} no Excel: a foto do histórico não foi atualizada.")


# =========================================================
# 🏁 EXECUÇÃO
# =========================================================


def preparar_consolidado(atualizar: bool = False, consolidar: bool = False):
    """Cache: se 1_downloads já tem planilhas, reaproveita (sem limite de data); só abre o portal com a pasta
    vazia ou com --atualizar. Se o consolidado é mais novo que as planilhas (já foram lidas), reaproveita também;
    senão reconsolida. Imports tardios: sem ciclo."""
    import atrasadas
    planilhas = [] if atualizar else atrasadas.planilhas_baixadas()
    if not planilhas:
        print("🔄 Baixando de novo do portal (--atualizar)…\n" if atualizar
              else "⚠️ Nenhuma planilha em 1_downloads; baixando do portal…\n")
        import extrator
        extrator.baixar_planilhas()
        print()
        planilhas = atrasadas.planilhas_baixadas()
    else:
        baixadas_em = datetime.fromtimestamp(max(p.stat().st_mtime for p in planilhas))
        print(f"♻️ Usando as {len(planilhas)} planilhas já baixadas em {baixadas_em:%d/%m/%Y %H:%M} "
              "(para baixar de novo: --atualizar)")
        if atrasadas.extracao_interrompida():
            print("⚠️ A última extração não terminou: podem faltar filiais. Se for o caso, rode com --atualizar.")

    arquivo = localizar_consolidado()
    if (consolidar or arquivo is None or not planilhas
            or arquivo.stat().st_mtime < max(p.stat().st_mtime for p in planilhas)):
        arquivo = atrasadas.consolidar()  # None se não houver planilhas: nunca cai num consolidado antigo
        print()
    else:
        print(f"♻️ Consolidado em dia: {arquivo.name}\n")
    return arquivo


def run(data_referencia: pd.Timestamp | None = None, modo: str = "simulacao",
        atualizar: bool = False, consolidar: bool = False):
    arquivo = preparar_consolidado(atualizar, consolidar)
    if arquivo is None:
        print("❌ Não foi possível gerar o consolidado; nada a disparar.")
        return
    contatos = carregar_contatos()
    if contatos is None:
        return

    print(f"📂 Lendo notas: {arquivo.name}")
    consolidado = ler_excel(arquivo)
    if COLUNA_CNPJ_NOTA not in consolidado.columns:
        print(f"⚠️ Consolidado sem a coluna '{COLUNA_CNPJ_NOTA}' (CNPJ); reconsolidando…\n")
        arquivo = preparar_consolidado(consolidar=True)
        consolidado = ler_excel(arquivo) if arquivo else pd.DataFrame()
        if COLUNA_CNPJ_NOTA not in consolidado.columns:
            print(f"❌ As planilhas do portal não trouxeram a coluna '{COLUNA_CNPJ_NOTA}'; nada a disparar.")
            return
    salvar_historico(consolidado)
    notas = notas_a_cobrar(consolidado, data_referencia)
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


def mostrar_uso():
    # Dois espaços depois de ⚠️/ℹ️: o terminal do Windows desenha esses emojis por cima do espaço seguinte.
    print("\nℹ️  Como usar:\n"
          "   python disparador.py               → simulação (não envia nada)\n"
          "   python disparador.py --teste       → envia tudo para o seu e-mail (EMAIL_TESTE)\n"
          "   python disparador.py --enviar      → envia para os fornecedores\n"
          "   python disparador.py --atualizar   → baixa de novo do portal (apaga sozinho as planilhas antigas;\n"
          "                                        combina com --teste/--enviar)\n"
          "\n"
          "⚠️  ATENÇÃO: NÃO APAGUE estas pastas:\n"
          "   atrasadas/5_historico   → fotos semanais usadas pela devolutiva\n"
          "   atrasadas/3_envios      → log de envios")


def main(consolidar: bool = False):
    """Ponto de entrada comum de extrator.py, atrasadas.py e disparador.py:
    [AAAA-MM-DD] [--teste | --enviar] [--atualizar]  (--atualizar ignora as planilhas já baixadas e baixa de novo)."""
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if flags - {"--teste", "--enviar", "--atualizar"} or {"--teste", "--enviar"} <= flags:
        raise SystemExit("Uso: python <script>.py [AAAA-MM-DD] [--teste | --enviar] [--atualizar]")
    datas = [a for a in sys.argv[1:] if not a.startswith("--")]
    run(pd.Timestamp(datas[0]) if datas else None,
        modo="teste" if "--teste" in flags else "enviar" if "--enviar" in flags else "simulacao",
        atualizar="--atualizar" in flags, consolidar=consolidar)
    mostrar_uso()


if __name__ == "__main__":
    main()
