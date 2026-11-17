import pandas as pd
import win32com.client as win32
import win32api
import pathlib
import tempfile

from config import obrigatorio

# =========================================================
# ⚙️ CONFIGURAÇÕES
# =========================================================
DOWNLOADS_DIR = pathlib.Path(__file__).parent / "atrasadas"
ARQUIVO_EXCEL = DOWNLOADS_DIR / "Atrasadas_Consolidado.xlsx"

# Quem está enviando e quem está recebendo o teste
EMAIL_REMETENTE = obrigatorio("EMAIL_REMETENTE")
EMAIL_DESTINATARIO = obrigatorio("EMAIL_TESTE")
FORNECEDOR_TESTE = "Gates" 

# =========================================================
# 📧 FUNÇÕES DE E-MAIL
# =========================================================

def gerar_tabela_html(df: pd.DataFrame) -> str:
    cols_desejadas = ['Dias', 'Filial', 'Razão Social', 'NF', 'Itens', 'Volumes', 'Status', 'Emissão', 'Agenda', 'Qt. Cancelamento']
    cols_existentes = [c for c in cols_desejadas if c in df.columns]
    df_exibicao = df[cols_existentes].copy()
    
    # Formata datas para dd/mm/aaaa
    if 'Emissão' in df_exibicao.columns:
        df_exibicao['Emissão'] = pd.to_datetime(df_exibicao['Emissão']).dt.strftime('%d/%m/%Y')
    if 'Agenda' in df_exibicao.columns:
        df_exibicao['Agenda'] = pd.to_datetime(df_exibicao['Agenda']).dt.strftime('%d/%m/%Y')
    
    # Destaca em vermelho negrito os atrasos maiores que 30 dias
    def formatar_dias(valor):
        try:
            if int(valor) > 30:
                return f'<span style="color: #d9534f; font-weight: bold;">{valor} dias</span>'
            return f"{valor} dias"
        except:
            return valor

    if 'Dias' in df_exibicao.columns:
        df_exibicao['Dias'] = df_exibicao['Dias'].apply(formatar_dias)

    html = df_exibicao.to_html(index=False, border=0, classes='tabela-moderna', escape=False, na_rep='-')
    
    css = """
    <style>
        .tabela-moderna { border-collapse: collapse; width: 100%; font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; margin-top: 15px; margin-bottom: 15px; }
        .tabela-moderna th { background-color: #2c3e50; color: #ffffff; border-bottom: 2px solid #34495e; padding: 10px; text-align: center; font-weight: 600; }
        .tabela-moderna td { border-bottom: 1px solid #e0e0e0; padding: 10px; text-align: center; color: #333333; }
        .tabela-moderna tr:nth-child(even) { background-color: #f9f9f9; }
        .tabela-moderna tr:hover { background-color: #f1f1f1; }
    </style>
    """
    return css + html

def localizar_consolidado() -> pathlib.Path | None:
    """Retorna o consolidado mais recente (o extrator salva com data/hora no nome quando o original está aberto)."""
    candidatos = [p for p in DOWNLOADS_DIR.glob("Atrasadas_Consolidado*.xlsx") if not p.name.startswith("~$")]
    return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None


def ler_excel(arquivo: pathlib.Path) -> pd.DataFrame:
    """Lê a planilha mesmo se ela estiver aberta no Excel (lê uma cópia temporária)."""
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


def enviar_email_teste():
    print("Lendo a planilha...")
    arquivo = localizar_consolidado()
    if arquivo is None:
        print(f"❌ Nenhum consolidado encontrado em {DOWNLOADS_DIR}. Rode o extrator.py primeiro.")
        return
    print(f"   Arquivo: {arquivo.name}")
    consolidado = ler_excel(arquivo)

    print(f"Filtrando fornecedor: {FORNECEDOR_TESTE}")
    df_fornecedor = consolidado[consolidado['Razão Social'].astype(str).str.contains(FORNECEDOR_TESTE, case=False, na=False)]
    
    if df_fornecedor.empty:
        print(f"⚠️ Nenhum registro encontrado.")
        return

    tabela_html = gerar_tabela_html(df_fornecedor)
    qtd_notas = len(df_fornecedor)
    qtd_filiais = df_fornecedor['Filial'].nunique()
    
    corpo_email = f"""
    <html>
    <body style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #f4f5f7; padding: 20px;">
        <div style="max-width: 800px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
            
            <div style="background-color: #d9534f; padding: 20px; text-align: center; color: white;">
                <h2 style="margin: 0; font-size: 20px;">Ação Necessária: Notas Atrasadas</h2>
                <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">Fornecedor: <b>{FORNECEDOR_TESTE.upper()}</b></p>
            </div>

            <div style="padding: 30px;">
                <p>Prezados, bom dia!</p>
                
                <p>Pedimos, por gentileza, que verifiquem por qual motivo as notas relacionadas abaixo foram agendadas e não entregues.</p>
                
                <div style="background-color: #f8f9fa; border-left: 4px solid #d9534f; padding: 15px; margin: 20px 0;">
                    <p style="margin: 0; font-size: 15px;"><b>Resumo do Atraso:</b></p>
                    <p style="margin: 5px 0 0 0;">Identificamos <b>{qtd_notas} notas</b> com agendamento não cumprido, distribuídas em <b>{qtd_filiais} filiais</b>.</p>
                </div>

                {tabela_html}
                
                <div style="background-color: #fcf8e3; border: 1px solid #faebcc; color: #8a6d3b; padding: 15px; border-radius: 4px; margin: 20px 0;">
                    <p style="margin: 0;"><b>Atenção:</b> Estamos tratando de filiais com problemas de capacidade. Ficamos no aguardo de um retorno <b>URGENTE</b> sobre a previsão de entrega.</p>
                </div>
                
                <p>Para verificar os detalhes ou atualizar a agenda, acesse o portal:</p>
                
                <div style="text-align: center; margin: 30px 0;">
                    <a href="https://app.brautoparts.com.br/srm/agendamentos" style="background-color: #2980b9; color: white; padding: 12px 25px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 15px; display: inline-block;">Acessar Portal de Agendamentos</a>
                </div>

                <p style="font-size: 13px; color: #7f8c8d;">Se o botão não funcionar, copie e cole o link no seu navegador: <br> https://app.brautoparts.com.br/srm/agendamentos</p>

                <br>
                <p>Atenciosamente,</p>
                <p style="margin-top: 10px;">
                    <b>Luan Araujo</b><br>
                    <span style="color: #7f8c8d; font-size: 13px;">Produtos - Matriz</span><br>
                    <span style="color: #7f8c8d; font-size: 13px;">Fone: +55 11 3155-7423</span><br>
                    <a href="mailto:luan.araujo@brautoparts.com.br" style="color: #2980b9; text-decoration: none;">luan.araujo@brautoparts.com.br</a>
                </p>
            </div>
            
            <div style="background-color: #f8f9fa; padding: 15px; text-align: center; font-size: 12px; color: #95a5a6; border-top: 1px solid #e0e0e0;">
                Este é um e-mail automático gerado pelo sistema de Agendamentos da Brautoparts.
            </div>
        </div>
    </body>
    </html>
    """
    
    assunto = f"TESTE: Ação Necessária: Notas Atrasadas – {FORNECEDOR_TESTE.upper()}"
    
    print(f"Disparando e-mail de {EMAIL_REMETENTE} para {EMAIL_DESTINATARIO} pelo Outlook...")
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
    
        mail.SentOnBehalfOfName = EMAIL_REMETENTE 
        
        mail.To = EMAIL_DESTINATARIO
        mail.Subject = assunto
        mail.HTMLBody = corpo_email
        
        mail.Send()
        print(f"✅ E-mail enviado com sucesso de {EMAIL_REMETENTE} para {EMAIL_DESTINATARIO}!")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

if __name__ == "__main__":
    enviar_email_teste()