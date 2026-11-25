import pandas as pd
import win32com.client as win32
import win32api
import os
import pathlib
import re
import tempfile

from config import FORNECEDOR_TESTE, PASTA_CONSOLIDADO, obrigatorio, opcional
from regras import (DIAS_CORRIDOS_CRITICO, DIAS_UTEIS_MINIMO, deve_cobrar, dias_corridos_desde,
                    dias_uteis_atraso, e_critico)

# =========================================================
# ⚙️ CONFIGURAÇÕES
# =========================================================

# Quem está enviando e quem está recebendo o teste
EMAIL_REMETENTE = obrigatorio("EMAIL_REMETENTE")
EMAIL_DESTINATARIO = obrigatorio("EMAIL_TESTE")

# 🚨 Regras em regras.py: dias exibidos e alerta crítico = dias CORRIDOS desde a Emissão;
#    mínimo para cobrar = 4º dia útil após a Agenda (sem fins de semana e feriados).

# =========================================================
# 📧 FUNÇÕES DE E-MAIL
# =========================================================

# =========================================================
# 🎨 PALETA (cores sólidas: o Outlook desktop não suporta rgba/transparência)
# =========================================================
VERMELHO = "#c9302c"         # destaque de atraso crítico
VERMELHO_CLARO = "#fdecea"   # fundo "vermelho translúcido" (linha crítica / resumo)
VERMELHO_BORDA = "#f1b8b5"
CINZA_BORDA = "#d0d4d9"
CINZA_CABECALHO = "#f2f3f5"
TEXTO = "#333333"
TEXTO_SUAVE = "#6b7280"
AZUL = "#1f6fb2"



def _fmt_data(valor) -> str:
    data = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    return "" if pd.isna(data) else data.strftime("%d/%m/%Y")


def _fmt_int(valor) -> str:
    return "" if pd.isna(valor) else str(int(valor))


def gerar_tabela_html(df: pd.DataFrame) -> str:
    """Tabela em grade (estilo do sketch), com estilos inline para funcionar no Outlook."""
    # Mais atrasadas primeiro
    if "Dias" in df.columns:
        df = df.sort_values("Dias", ascending=False)

    colunas = [  # (título, alinhamento)
        ("Atraso", "center"), ("Filial", "left"), ("Razão Social", "left"), ("NF", "center"),
        ("Itens / Volumes", "center"), ("Status", "center"), ("Emissão", "center"),
        ("Agenda", "center"), ("Qt.", "center"), ("Cancelamento", "center"),
    ]
    borda = f"border:1px solid {CINZA_BORDA};"
    th_style = (f"{borda} background-color:{CINZA_CABECALHO}; color:{TEXTO}; padding:8px 6px; "
                "font-size:11px; font-weight:bold; text-transform:uppercase; white-space:nowrap;")

    cabecalho = "".join(f'<th style="{th_style} text-align:{al};">{titulo}</th>' for titulo, al in colunas)

    linhas = []
    for _, r in df.iterrows():
        dias = int(r["Dias"]) if pd.notna(r.get("Dias")) else None
        critico = bool(r.get("Crítico", False))
        fundo = f"background-color:{VERMELHO_CLARO};" if critico else ""

        texto_dias = "" if dias is None else f"{dias} dias"
        dias_html = f'<b style="color:{VERMELHO};">{texto_dias}</b>' if critico else texto_dias
        cancelamento = _fmt_data(r.get("Cancelamento"))
        alerta = f' &nbsp;<span style="color:{VERMELHO}; font-size:15px;">&#9888;</span>' if critico else ""

        valores = [
            dias_html,
            r.get("Filial", ""),
            r.get("Razão Social", ""),
            _fmt_int(r.get("NF")),
            f'{_fmt_int(r.get("Itens"))} / {_fmt_int(r.get("Volumes"))}',
            r.get("Status", ""),
            _fmt_data(r.get("Emissão")),
            _fmt_data(r.get("Agenda")),
            _fmt_int(r.get("Qt. Cancelamento")) or "-",
            (cancelamento or "-") + alerta,
        ]
        tds = "".join(
            f'<td style="{borda} {fundo} padding:8px 6px; font-size:12px; color:{TEXTO}; text-align:{al};'
            f'{"" if titulo == "Razão Social" else " white-space:nowrap;"}">{v}</td>'
            for v, (titulo, al) in zip(valores, colunas)
        )
        linhas.append(f"<tr>{tds}</tr>")

    return (
        f'<table cellpadding="0" cellspacing="0" width="100%" '
        f'style="border-collapse:collapse; font-family:Segoe UI, Arial, sans-serif; margin:10px 0 20px 0;">'
        f"<thead><tr>{cabecalho}</tr></thead><tbody>{''.join(linhas)}</tbody></table>"
    )

PASTA_ASSINATURAS = pathlib.Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"


def carregar_assinatura() -> str:
    """Lê a assinatura do Outlook de quem está rodando o script.

    Usa a assinatura definida em ASSINATURA_OUTLOOK no .env (nome do arquivo, sem .htm)
    ou, se não definida, a mais recente da pasta de assinaturas do Outlook.
    """
    arquivos = sorted(PASTA_ASSINATURAS.glob("*.htm"), key=lambda p: p.stat().st_mtime, reverse=True)
    nome = opcional("ASSINATURA_OUTLOOK")
    if nome:
        arquivos = [a for a in arquivos if a.stem.lower() == nome.lower()]
    if not arquivos:
        print("⚠️ Assinatura do Outlook não encontrada; o e-mail sairá sem assinatura.")
        return ""

    bruto = arquivos[0].read_bytes()
    if bruto[:2] in (b"\xff\xfe", b"\xfe\xff"):
        html = bruto.decode("utf-16")
    else:
        try:
            html = bruto.decode("utf-8")
        except UnicodeDecodeError:
            html = bruto.decode("cp1252")

    corpo = re.search(r"<body[^>]*>(.*)</body>", html, re.S | re.I)
    corpo = corpo.group(1) if corpo else html
    # Remove os parágrafos vazios que o Outlook coloca antes/depois da assinatura
    corpo = re.sub(r"<p[^>]*>(?:\s|&nbsp;)*(?:</p>|(?=<p|<table|$))", "", corpo, flags=re.I)
    # Remove larguras fixas (ex.: 887px) para não estourar o layout do e-mail
    corpo = re.sub(r"\bwidth:\s*\d+px;?", "", corpo, flags=re.I)          # style="WIDTH: 887px"
    corpo = re.sub(r"\swidth=\"?\d+\"?(?=[\s>])", "", corpo, flags=re.I)  # width=887 (mantém %)
    return corpo.strip()


def _espaco(altura: int = 16) -> str:
    """Espaçador vertical (o Outlook ignora margin em tabelas)."""
    return (f'<table cellpadding="0" cellspacing="0" width="100%"><tr>'
            f'<td height="{altura}" style="height:{altura}px; line-height:{altura}px; font-size:1px;">&nbsp;</td>'
            f"</tr></table>")


def _bloco(conteudo: str, fundo: str, borda_esquerda: str = "", borda: str = "") -> str:
    """Caixa de destaque feita com <table> (o Outlook ignora padding/raio em <div>)."""
    estilo_borda = f"border:1px solid {borda};" if borda else ""
    estilo_borda += f"border-left:4px solid {borda_esquerda};" if borda_esquerda else ""
    return (_espaco() +
            f'<table cellpadding="0" cellspacing="0" width="100%"><tr>'
            f'<td style="background-color:{fundo}; {estilo_borda} padding:12px 16px; font-size:13px; color:{TEXTO};">'
            f"{conteudo}</td></tr></table>")


def montar_corpo_email(df_fornecedor: pd.DataFrame, fornecedor: str) -> str:
    qtd_notas = len(df_fornecedor)
    qtd_filiais = df_fornecedor["Filial"].nunique()
    qtd_criticas = int(df_fornecedor["Crítico"].sum()) if "Crítico" in df_fornecedor.columns else 0
    s_notas = "nota" if qtd_notas == 1 else "notas"
    s_filiais = "filial" if qtd_filiais == 1 else "filiais"
    criticas = (f'<br><span style="color:{VERMELHO};">&#9888; <b>{qtd_criticas}</b> com mais de {DIAS_CORRIDOS_CRITICO} dias corridos de atraso.</span>'
                if qtd_criticas else "")

    resumo = _bloco(
        f'<b style="font-size:14px;">Resumo da Ocorrência</b><br>'
        f"Identificamos <b>{qtd_notas} {s_notas}</b> em situação de No-Show, "
        f"distribuídas em <b>{qtd_filiais} {s_filiais}</b>.{criticas}",
        fundo=VERMELHO_CLARO, borda=VERMELHO_BORDA, borda_esquerda=VERMELHO,
    )
    tratativas = _bloco(
        "<b>Para regularizar a situação, solicitamos as seguintes tratativas:</b><br>"
        "&bull; <b>Se a nota foi entregue:</b> encaminhar o comprovante de entrega (canhoto assinado) em resposta a este e-mail.<br>"
        "&bull; <b>Se a nota não foi entregue:</b> realizar o cancelamento do agendamento atual no portal e criar um novo registro de agendamento.",
        fundo="#eef5fb", borda_esquerda=AZUL,
    )
    aviso = _bloco(
        "<b>Atenção:</b> é de suma importância que o fornecedor dê um retorno com as devidas tratativas.",
        fundo="#fff8e1", borda_esquerda="#e0a800",
    )
    assinatura = carregar_assinatura()

    return f"""
<html>
<body style="margin:0; padding:20px; background-color:#f4f5f7; font-family:Segoe UI, Arial, sans-serif; font-size:14px; color:{TEXTO}; line-height:1.5;">
<table cellpadding="0" cellspacing="0" width="100%" style="max-width:900px; margin:0 auto; background-color:#ffffff; border:1px solid #e3e5e8;">
  <!-- Cabeçalho -->
  <tr><td style="background-color:{VERMELHO}; padding:18px 24px; text-align:center; color:#ffffff;">
    <div style="font-size:20px; font-weight:bold;">Ação Necessária: Notas Atrasadas</div>
    <div style="font-size:14px;">Fornecedor: <b>{fornecedor.upper()}</b></div>
  </td></tr>

  <!-- Corpo -->
  <tr><td style="padding:24px;">
    <p style="margin:0 0 12px 0;">Prezados,</p>
    <p style="margin:0 0 4px 0;">Identificamos que as notas fiscais relacionadas abaixo encontram-se com o status de <b>No-Show</b>.</p>
    <p style="margin:0;">Isso significa que a entrega foi agendada no portal, porém a transportadora não compareceu ou não realizou a entrega na data acordada.</p>

    {resumo}
    {_espaco()}
    {gerar_tabela_html(df_fornecedor)}
    {tratativas}
    {aviso}
    {_espaco(20)}

    <p style="margin:0; text-align:center;">Para verificar os detalhes ou atualizar a agenda, acesse o portal:</p>
    {_espaco(10)}
    <table cellpadding="0" cellspacing="0" align="center" style="margin:0 auto;"><tr>
      <td align="center" style="background-color:{AZUL}; padding:7px 16px;">
        <a href="https://app.brautoparts.com.br/srm/agendamentos" style="color:#ffffff; text-decoration:none; font-weight:bold; font-size:12px;">Acessar Portal de Agendamentos</a>
      </td></tr></table>
    {_espaco(24)}

    <p style="margin:0;">Atenciosamente,</p>
    {assinatura}
  </td></tr>

  <!-- Rodapé -->
  <tr><td style="background-color:#f8f9fa; border-top:1px solid #e3e5e8; padding:12px; text-align:center; font-size:11px; color:{TEXTO_SUAVE};">
    Este é um e-mail automático gerado pelo sistema de Agendamentos da Brautoparts.
  </td></tr>
</table>
</body>
</html>"""


def localizar_consolidado() -> pathlib.Path | None:
    candidatos = [p for p in PASTA_CONSOLIDADO.glob("Atrasadas_Consolidado*.xlsx") if not p.name.startswith("~$")]
    return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None

def ler_excel(arquivo: pathlib.Path) -> pd.DataFrame:
    try:
        return pd.read_excel(arquivo)
    except PermissionError:
        print(f"⚠️ {arquivo.name} está aberto no Excel; lendo uma cópia temporária...")
        copia = pathlib.Path(tempfile.gettempdir()) / f"copia_{arquivo.name}"
        win32api.CopyFile(str(arquivo), str(copia), 0)
        try:
            return pd.read_excel(copia)
        finally:
            copia.unlink(missing_ok=True)

def filtrar_fornecedor(consolidado: pd.DataFrame, fornecedor: str,
                       data_referencia: pd.Timestamp | None = None) -> pd.DataFrame:
    """Notas do fornecedor (Razão Social contém o nome) a partir do 4º dia útil após a Agenda.

    O atraso ('Dias') é RECALCULADO aqui, na data de referência (padrão: hoje), para não
    depender do valor gravado no Excel no dia da consolidação.
    """
    for coluna in ("Razão Social", "Agenda", "Emissão"):
        if coluna not in consolidado.columns:
            print(f"❌ Coluna '{coluna}' não encontrada no consolidado.")
            return consolidado.iloc[0:0]

    print(f"🔍 Filtrando fornecedor: {fornecedor}")
    df = consolidado[
        consolidado["Razão Social"].astype(str).str.contains(fornecedor, case=False, na=False, regex=False)
    ]
    return notas_a_cobrar(df, data_referencia)


def notas_a_cobrar(df: pd.DataFrame, data_referencia: pd.Timestamp | None = None) -> pd.DataFrame:
    """Recalcula 'Dias'/'Crítico' na data de referência e mantém só as notas a partir do 4º dia útil."""
    df = df.copy()
    df["Dias"] = dias_corridos_desde(df["Emissão"], data_referencia)  # exibido no e-mail (HOJE() - Emissão)
    df["Crítico"] = e_critico(df["Dias"])
    cobrar = deve_cobrar(dias_uteis_atraso(df["Agenda"], data_referencia))  # única regra em dias úteis
    return df[cobrar]


def enviar_email(df_fornecedor: pd.DataFrame, fornecedor: str, destinatario: str = EMAIL_DESTINATARIO,
                 prefixo_assunto: str = "") -> bool:
    """Monta e envia pelo Outlook o e-mail de cobrança de um fornecedor. Retorna True se enviou."""
    if df_fornecedor.empty:
        print(f"⚠️ Nenhuma nota de {fornecedor} com {DIAS_UTEIS_MINIMO}+ dias úteis de atraso; e-mail não enviado.")
        return False

    corpo_email = montar_corpo_email(df_fornecedor, fornecedor)
    assunto = f"{prefixo_assunto}Ação Necessária: Notas Atrasadas – {fornecedor.upper()}"

    print(f"📧 Disparando e-mail para {fornecedor}...")
    return enviar_outlook(destinatario, assunto, corpo_email)


def enviar_outlook(destinatario: str, assunto: str, corpo_html: str) -> bool:
    """Envio pelo Outlook (fonte única: cobrança e devolutiva). Vários destinatários separados por ';'."""
    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = EMAIL_REMETENTE
        mail.To = destinatario
        mail.Subject = assunto
        mail.HTMLBody = corpo_html
        # mail.Display()  # Descomente para revisar antes de enviar (e comente o Send)
        mail.Send()
        print(f"✅ E-mail enviado com sucesso para {destinatario}!")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")
        print("   Verifique se o Outlook está aberto e configurado nesta máquina.")
        return False


def enviar_email_teste(data_referencia: pd.Timestamp | None = None):
    """Envia o e-mail a partir do consolidado já salvo em disco (sem abrir o portal)."""
    print("📂 Lendo a planilha...")
    arquivo = localizar_consolidado()
    if arquivo is None:
        print("❌ Nenhum consolidado encontrado.")
        return
    consolidado = ler_excel(arquivo)
    enviar_email(filtrar_fornecedor(consolidado, FORNECEDOR_TESTE, data_referencia), FORNECEDOR_TESTE)


if __name__ == "__main__":
    import sys
    # Opcional: simular outra data de envio, ex.: python teste_email.py 2026-11-24
    enviar_email_teste(pd.Timestamp(sys.argv[1]) if len(sys.argv) > 1 else None)