import pathlib

import pandas as pd

from config import PASTA_CONSOLIDADO, PASTA_DOWNLOADS
from regras import dias_corridos_desde

# =========================================================
# ⚙️ CONFIGURAÇÕES
# =========================================================

# Caminhos definidos em config.py (atrasadas/1_downloads e atrasadas/2_consolidado)
ARQUIVO_FINAL = PASTA_CONSOLIDADO / "Atrasadas_Consolidado.xlsx"

# O layout e o envio do e-mail ficam em teste_email.py (fonte única).


# =========================================================
# 📥 LEITURA DAS PLANILHAS BAIXADAS
# =========================================================


def carregar_tabelas() -> list[pd.DataFrame]:
    """Lê todas as planilhas baixadas na pasta PASTA_DOWNLOADS (exceto o consolidado)."""
    if not PASTA_DOWNLOADS.exists():
        print(f"⚠️ Pasta não encontrada: {PASTA_DOWNLOADS}")
        return []

    tabelas = []
    for arquivo in sorted(PASTA_DOWNLOADS.glob("*.xls*")):
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


def consolidar(extrair_se_vazio: bool = True) -> pathlib.Path | None:
    """Consolida as planilhas de 1_downloads; se não houver nenhuma, extrai antes do portal.

    Retorna o caminho do consolidado salvo (ou None se não houver dados).
    """
    print(f"📂 Lendo planilhas em: {PASTA_DOWNLOADS}")
    tabelas = carregar_tabelas()

    if not tabelas and not extrair_se_vazio:
        print("❌ Nenhuma planilha baixada; não há consolidado para salvar.")
        return None
    if not tabelas:
        print("⚠️ Nenhuma planilha baixada; iniciando a extração no portal…\n")
        import extrator  # import tardio: só exige BRA_EMAIL/BRA_SENHA quando precisa extrair
        extrator.baixar_planilhas()
        tabelas = carregar_tabelas()
        if not tabelas:
            print("❌ A extração não trouxe nenhuma planilha; não há consolidado para salvar.")
            return None

    print("\n📊 Consolidando todas as planilhas em uma única tabela…")
    consolidado = pd.concat(tabelas, ignore_index=True)

    # ----------------------------------------------------
    # TRATAMENTO DA PLANILHA (Conforme modelo)
    # ----------------------------------------------------
    print("🧹 Tratando a planilha base...")

    # 1. Excluir colunas que não vão no e-mail.
    #    "Emitente" (CNPJ do fornecedor) é mantida: o disparador cruza com a base de contatos por ela.
    colunas_excluir = [
        "Empresa", "CNPJ Filial", "Valor (R$)", "Protocolo",
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

    # 3. Coluna 'Dias' = dias CORRIDOS desde a Emissão (HOJE() - Emissão).
    #    É uma foto do dia da consolidação, para conferência; no envio o valor é recalculado.
    if 'Emissão' in consolidado.columns:
        consolidado['Dias'] = dias_corridos_desde(consolidado['Emissão'])
        cols = ['Dias'] + [c for c in consolidado.columns if c != 'Dias']
        consolidado = consolidado[cols]
    else:
        print("⚠️ Coluna 'Emissão' não encontrada; coluna 'Dias' não foi calculada.")

    # Salva o arquivo tratado (para conferência)
    try:
        consolidado.to_excel(ARQUIVO_FINAL, index=False)
        print(f"✅ Arquivo único criado em: {ARQUIVO_FINAL}")
        return ARQUIVO_FINAL
    except PermissionError:
        # Acontece quando o consolidado anterior está aberto no Excel
        alternativo = ARQUIVO_FINAL.with_name(f"Atrasadas_Consolidado_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx")
        consolidado.to_excel(alternativo, index=False)
        print(f"⚠️ {ARQUIVO_FINAL.name} está aberto no Excel. Salvo como: {alternativo}")
        return alternativo


if __name__ == "__main__":
    # Reconsolida (extraindo antes, se 1_downloads estiver vazia) e segue para o disparo.
    # Aceita os mesmos parâmetros do disparador: [AAAA-MM-DD] [--teste | --enviar]
    import disparador
    disparador.main(consolidar=True)
