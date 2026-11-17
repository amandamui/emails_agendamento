import pathlib

import pandas as pd
import win32com.client as win32

from config import obrigatorio

# =========================================================
# ⚙️ CONFIGURAÇÕES
# =========================================================

# Pasta onde ficam as planilhas baixadas (mesma usada pelo teste_email.py)
DOWNLOADS_DIR = pathlib.Path(__file__).parent / "atrasadas"
ARQUIVO_FINAL = DOWNLOADS_DIR / "Atrasadas_Consolidado.xlsx"

FORNECEDOR_TESTE = "Gates"  # Nome do fornecedor para filtrar
EMAIL_TESTE = obrigatorio("EMAIL_TESTE")  # Destinatário do teste (definido no .env)

# =========================================================
# 📧 FUNÇÕES DE E-MAIL (FASE 1)
# =========================================================


def gerar_tabela_html(df: pd.DataFrame) -> str:
    """Gera o HTML da tabela formatado com cabeçalho em negrito."""
    # Garante que a coluna 'Dias' seja a primeira se existir
    cols_desejadas = ['Dias', 'Filial', 'Razão Social', 'NF', 'Itens', 'Volumes', 'Status', 'Emissão', 'Agenda', 'Qt. Cancelamento']

    # Filtra apenas as colunas que existem no DataFrame (para evitar erro caso falte alguma)
    cols_existentes = [c for c in cols_desejadas if c in df.columns]
    df_exibicao = df[cols_existentes].copy()

    # Mostra as datas (Emissão, Agenda) como dd/mm/aaaa (sem horário)
    for col in df_exibicao.columns:
        if pd.api.types.is_datetime64_any_dtype(df_exibicao[col]):
            df_exibicao[col] = df_exibicao[col].dt.strftime('%d/%m/%Y')
    if 'Qt. Cancelamento' in df_exibicao.columns:
        df_exibicao['Qt. Cancelamento'] = df_exibicao['Qt. Cancelamento'].map(lambda v: '' if pd.isna(v) else int(v))

    # Converte para HTML com estilização básica
    html = df_exibicao.to_html(index=False, border=0, classes='tabela-atrasos', na_rep='')

    # CSS inline para deixar parecido com a imagem
    css = """
    <style>
        .tabela-atrasos {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
            font-size: 12px;
        }
        .tabela-atrasos th {
            background-color: #f2f2f2;
            border: 1px solid #ddd;
            padding: 8px;
            text-align: center;
            font-weight: bold;
        }
        .tabela-atrasos td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: center;
        }
    </style>
    """
    return css + html


def enviar_email_teste(df_fornecedor: pd.DataFrame, nome_fornecedor: str, email_destino: str):
    """Monta e envia o e-mail conforme a imagem para o destinatário de teste."""

    print(f"📧 Preparando e-mail para {nome_fornecedor}...")

    # 1. Gera a tabela HTML
    tabela_html = gerar_tabela_html(df_fornecedor)

    # 2. Monta o corpo do e-mail (HTML)
    corpo_email = f"""
    <html>
    <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
        <p>Prezados, bom dia!</p>

        <p>Pedimos, por gentileza, que verifiquem por qual motivo as notas relacionadas abaixo foram agendadas e não entregues.</p>

        <br>
        {tabela_html}
        <br>

        <p>Estamos tratando de filiais com problemas de capacidade.</p>

        <p>Ficamos no aguardo de um retorno <span style="color: red; font-weight: bold;">URGENTE</span>.</p>

        <br>
        <p>Atenciosamente,</p>
        <p><b>Luan Araujo</b><br>
        Produtos - Matriz<br>
        Fone: +55 11 3155-7423<br>
        luan.araujo@brautoparts.com.br</p>
    </body>
    </html>
    """

    # 3. Assunto
    assunto = f"NOTAS ATRASADAS – {nome_fornecedor.upper()}"

    # 4. Envio via Outlook (win32com)
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = email_destino
        mail.Subject = assunto
        mail.HTMLBody = corpo_email

        # mail.Display()  # Descomente para ver o e-mail antes de enviar (e comente o Send)
        mail.Send()
        print(f"✅ E-mail enviado com sucesso para {email_destino}!")

    except Exception as e:
        print(f"❌ Erro ao enviar e-mail via Outlook: {e}")
        print("   Verifique se o Outlook está aberto e configurado nesta máquina.")


# =========================================================
# 📥 LEITURA DAS PLANILHAS BAIXADAS
# =========================================================


def carregar_tabelas() -> list[pd.DataFrame]:
    """Lê todas as planilhas baixadas na pasta DOWNLOADS_DIR (exceto o consolidado)."""
    if not DOWNLOADS_DIR.exists():
        print(f"⚠️ Pasta não encontrada: {DOWNLOADS_DIR}")
        return []

    tabelas = []
    for arquivo in sorted(DOWNLOADS_DIR.glob("*.xls*")):
        if arquivo.name.startswith("Atrasadas_Consolidado") or arquivo.name.startswith("~$"):
            continue  # ignora o próprio consolidado e arquivos temporários do Excel
        try:
            tabelas.append(pd.read_excel(arquivo))
            print(f"   📄 Lido: {arquivo.name}")
        except Exception as e:
            print(f"   ❌ Erro ao ler {arquivo.name}: {e}")
    return tabelas


# =========================================================
# 🏁 EXECUÇÃO
# =========================================================


def run():
    # TODO: se o login/download automático (navegador) for integrado aqui,
    # ele deve salvar as planilhas em DOWNLOADS_DIR antes desta etapa.
    print(f"📂 Lendo planilhas em: {DOWNLOADS_DIR}")
    tabelas = carregar_tabelas()

    if not tabelas:
        print("⚠️ Nenhuma tabela foi carregada; não há consolidado para salvar.")
        return

    print("\n📊 Consolidando todas as planilhas em uma única tabela…")
    consolidado = pd.concat(tabelas, ignore_index=True)

    # ----------------------------------------------------
    # TRATAMENTO DA PLANILHA (Conforme modelo)
    # ----------------------------------------------------
    print("🧹 Tratando a planilha base...")

    # 1. Excluir colunas que não vão no e-mail
    colunas_excluir = [
        "Empresa", "CNPJ Filial", "Emitente", "Valor (R$)", "Protocolo",
        "Vencimento", "Agendou", "Período", "Entrega", "Cancelamento",
        # O portal exporta sem acento ("Usuario"); mantém as duas grafias
        "Usuário Agendamento", "Usuario Agendamento",
        "Usuário Cancelamento", "Usuario Cancelamento",
    ]
    # Exclui apenas se a coluna existir (para não dar erro)
    consolidado = consolidado.drop(columns=[c for c in colunas_excluir if c in consolidado.columns])

    # 2. Calcular coluna 'Dias' (HOJE() - Data de Emissão)
    if 'Emissão' in consolidado.columns:
        # Converte para data caso venha como string
        consolidado['Emissão'] = pd.to_datetime(consolidado['Emissão'], dayfirst=True, errors='coerce')
        consolidado['Dias'] = (pd.Timestamp.today().normalize() - consolidado['Emissão']).dt.days
        # Reorganiza para 'Dias' ser a primeira coluna
        cols = ['Dias'] + [c for c in consolidado.columns if c != 'Dias']
        consolidado = consolidado[cols]
    else:
        print("⚠️ Coluna 'Emissão' não encontrada; coluna 'Dias' não foi calculada.")

    # Salva o arquivo tratado (para conferência)
    try:
        consolidado.to_excel(ARQUIVO_FINAL, index=False)
        print(f"✅ Arquivo único criado em: {ARQUIVO_FINAL}")
    except PermissionError:
        # Acontece quando o consolidado anterior está aberto no Excel
        alternativo = ARQUIVO_FINAL.with_name(f"Atrasadas_Consolidado_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx")
        consolidado.to_excel(alternativo, index=False)
        print(f"⚠️ {ARQUIVO_FINAL.name} está aberto no Excel. Salvo como: {alternativo}")

    # ----------------------------------------------------
    # FASE 1: TESTE COM 1 FORNECEDOR
    # ----------------------------------------------------
    if 'Razão Social' not in consolidado.columns:
        print("❌ Coluna 'Razão Social' não encontrada; não é possível filtrar o fornecedor.")
        return

    print(f"\n🔍 Filtrando dados do fornecedor: {FORNECEDOR_TESTE}")

    # Filtra o DataFrame pela Razão Social (contains para pegar "Gates do Brasil...")
    df_fornecedor = consolidado[
        consolidado['Razão Social'].astype(str).str.contains(FORNECEDOR_TESTE, case=False, na=False, regex=False)
    ]

    if not df_fornecedor.empty:
        enviar_email_teste(df_fornecedor, FORNECEDOR_TESTE, EMAIL_TESTE)
    else:
        print(f"⚠️ Nenhum registro encontrado para o fornecedor {FORNECEDOR_TESTE}.")


if __name__ == "__main__":
    run()
