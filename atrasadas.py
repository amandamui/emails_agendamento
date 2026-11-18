import pathlib

import pandas as pd

from config import FORNECEDOR_TESTE
from regras import dias_corridos_atraso
from teste_email import enviar_email, filtrar_fornecedor

# =========================================================
# ⚙️ CONFIGURAÇÕES
# =========================================================

# Pasta onde ficam as planilhas baixadas (mesma usada pelo teste_email.py)
DOWNLOADS_DIR = pathlib.Path(__file__).parent / "atrasadas"
ARQUIVO_FINAL = DOWNLOADS_DIR / "Atrasadas_Consolidado.xlsx"

# O layout e o envio do e-mail ficam em teste_email.py (fonte única).


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
        "Vencimento", "Agendou", "Período", "Entrega",  # "Cancelamento" (data) é mantida: vai no e-mail
        # O portal exporta sem acento ("Usuario"); mantém as duas grafias
        "Usuário Agendamento", "Usuario Agendamento",
        "Usuário Cancelamento", "Usuario Cancelamento",
    ]
    # Exclui apenas se a coluna existir (para não dar erro)
    consolidado = consolidado.drop(columns=[c for c in colunas_excluir if c in consolidado.columns])

    # 2. Datas como data (o portal pode exportar como texto)
    for col in ('Emissão', 'Agenda'):
        if col in consolidado.columns:
            consolidado[col] = pd.to_datetime(consolidado[col], dayfirst=True, errors='coerce')

    # 3. Coluna 'Dias' = dias CORRIDOS de atraso desde a Agenda (regra em regras.py).
    #    É uma foto do dia da consolidação, para conferência; no envio o valor é recalculado.
    if 'Agenda' in consolidado.columns:
        consolidado['Dias'] = dias_corridos_atraso(consolidado['Agenda'])
        cols = ['Dias'] + [c for c in consolidado.columns if c != 'Dias']
        consolidado = consolidado[cols]
    else:
        print("⚠️ Coluna 'Agenda' não encontrada; coluna 'Dias' (atraso) não foi calculada.")

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
    # FASE 1: TESTE COM 1 FORNECEDOR (cobra a partir do 4º dia útil após a Agenda)
    # ----------------------------------------------------
    print()
    enviar_email(filtrar_fornecedor(consolidado, FORNECEDOR_TESTE), FORNECEDOR_TESTE)


if __name__ == "__main__":
    run()
