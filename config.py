"""Configurações sensíveis lidas do arquivo .env (que NUNCA vai para o Git).

Copie o .env.example para .env e preencha os valores.
"""
import os
import pathlib

from dotenv import load_dotenv

PASTA_PROJETO = pathlib.Path(__file__).parent
load_dotenv(PASTA_PROJETO / ".env")


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
