import pathlib

import pandas as pd

from config import MARCADOR_EXTRACAO, PASTA_CONSOLIDADO, PASTA_DOWNLOADS
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


def planilhas_baixadas() -> list[pathlib.Path]:
    """Planilhas de 1_downloads (o cache da extração), sem os temporários do Excel (~$)."""
    return sorted(p for p in PASTA_DOWNLOADS.glob("Agendamento*.xls*") if not p.name.startswith("~$"))


def extracao_interrompida() -> bool:
    """Há planilhas, mas o extrator não chegou ao fim (marcador ausente): podem faltar filiais."""
    return bool(planilhas_baixadas()) and not MARCADOR_EXTRACAO.exists()


def carregar_tabelas() -> list[pd.DataFrame]:
    """Lê todas as planilhas baixadas na pasta PASTA_DOWNLOADS."""
    tabelas = []
    for arquivo in planilhas_baixadas():
        try:
            tabelas.append(pd.read_excel(arquivo))
            print(f"   📄 Lido: {arquivo.name}")
        except Exception as e:
            print(f"   ❌ Erro ao ler {arquivo.name}: {e}")
    return tabelas


# =========================================================
# 🏁 EXECUÇÃO
# =========================================================


def consolidar() -> pathlib.Path | None:
    """Consolida as planilhas de 1_downloads (quem decide se extrai antes é o disparador).

    Retorna o caminho do consolidado salvo (ou None se não houver dados).
    """
    print(f"📂 Lendo planilhas em: {PASTA_DOWNLOADS}")
    tabelas = carregar_tabelas()
    if not tabelas:
        print("❌ Nenhuma planilha baixada; não há consolidado para salvar.")
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
    # Aceita os mesmos parâmetros do disparador: [AAAA-MM-DD] [--teste | --enviar] [--atualizar]
    import disparador
    disparador.main(consolidar=True)
