"""Configurações sensíveis lidas do arquivo .env (que NUNCA vai para o Git).

Copie o .env.example para .env e preencha os valores.
"""
import os
import pathlib

from dotenv import load_dotenv

PASTA_PROJETO = pathlib.Path(__file__).parent
load_dotenv(PASTA_PROJETO / ".env")

# Pastas de saída (fonte única dos caminhos), uma por etapa do fluxo. Nunca versionar (LGPD).
PASTA_DADOS = PASTA_PROJETO / "atrasadas"
PASTA_DOWNLOADS = PASTA_DADOS / "1_downloads"      # planilhas brutas de cada filial
PASTA_CONSOLIDADO = PASTA_DADOS / "2_consolidado"  # Atrasadas_Consolidado.xlsx
PASTA_ENVIOS = PASTA_DADOS / "3_envios"            # Log_Envios.csv e Fornecedores_Pendentes.xlsx
PASTA_ERROS = PASTA_DADOS / "4_erros"              # prints das filiais com problema
PASTA_HISTORICO = PASTA_DADOS / "5_historico"      # Atrasadas_AAAA-MM-DD.xlsx: foto semanal para a devolutiva
PASTA_CONTATOS = PASTA_PROJETO / "base fornecedores"
# Gravado pelo extrator ao terminar todas as filiais; sem ele, a pasta 1_downloads pode estar incompleta.
MARCADOR_EXTRACAO = PASTA_DOWNLOADS / "_extracao_concluida.txt"
for _pasta in (PASTA_DOWNLOADS, PASTA_CONSOLIDADO, PASTA_ENVIOS, PASTA_ERROS, PASTA_HISTORICO):
    _pasta.mkdir(parents=True, exist_ok=True)


def obrigatorio(nome: str) -> str:
    valor = os.getenv(nome, "").strip()
    if not valor:
        raise SystemExit(
            f"❌ Variável '{nome}' não configurada. "
            f"Copie o .env.example para .env e preencha os valores."
        )
    return valor


def opcional(nome: str, padrao: str = "") -> str:
    return os.getenv(nome, padrao).strip()


# Desativa a verificação de certificado HTTPS só se for explicitamente necessário
# (ex.: proxy corporativo que intercepta TLS). Padrão: verificação LIGADA.
IGNORAR_CERTIFICADO = opcional("IGNORAR_CERTIFICADO", "0") == "1"

# Fornecedor da fase de teste (filtro "contém" na Razão Social)
FORNECEDOR_TESTE = opcional("FORNECEDOR_TESTE", "Gates")
