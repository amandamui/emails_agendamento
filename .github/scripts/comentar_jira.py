"""Roda no GitHub Actions a cada push: o Claude resume as mudanças e comenta no card do Jira.

O card é identificado pela chave (ex.: A122-11) no nome da branch ou na mensagem do commit.
Sem chave, não faz nada.

Variáveis de ambiente (GitHub → Settings → Secrets and variables → Actions):
  ANTHROPIC_API_KEY, JIRA_BASE_URL (ex.: https://suaempresa.atlassian.net), JIRA_EMAIL, JIRA_API_TOKEN
Fornecidas pelo próprio GitHub: GITHUB_EVENT_PATH, GITHUB_REF_NAME, GITHUB_REPOSITORY, GITHUB_SERVER_URL
"""
import json
import os
import re
import subprocess
import sys

import anthropic
import requests

# Projeto A122 no Jira; o número depois do hífen identifica o ticket (A122-11, A122-14, ...).
# Pode ser trocado pela variável JIRA_PROJETO sem mexer no código.
PROJETO_JIRA = os.environ.get("JIRA_PROJETO", "A122")
CHAVE_JIRA = re.compile(rf"\b{re.escape(PROJETO_JIRA)}-\d+\b", re.IGNORECASE)
LIMITE_DIFF = 150_000  # caracteres; acima disso o Claude recebe só a lista de arquivos alterados

# Arquivos que nunca devem ir para o Claude nem para o Jira (dados de negócio/segredos)
IGNORAR = (":(exclude)*.xlsx", ":(exclude)*.xls", ":(exclude)*.png", ":(exclude).env", ":(exclude)atrasadas/*")

PROMPT_SISTEMA = """Você escreve comentários de acompanhamento em cards do Jira, em português do Brasil,
para a equipe de negócio (não técnica) de uma automação que cobra fornecedores por notas fiscais
agendadas e não entregues.

Receberá os commits de um push (mensagens e diff). Escreva:
1. Uma linha de resumo do que mudou.
2. Até 5 tópicos com o impacto prático (regras de negócio, e-mail enviado, planilha), sem jargão.
3. Se houver risco ou algo que precise de validação do negócio, um tópico "Atenção:".

Seja objetivo, no máximo ~150 palavras. Use a sintaxe de texto do Jira: tópicos com "* ",
negrito com *texto*. Não invente nada que não esteja no diff."""


def git(*args: str) -> str:
    return subprocess.run(["git", *args], capture_output=True, text=True, check=True).stdout


def commits_do_push() -> list[dict]:
    with open(os.environ["GITHUB_EVENT_PATH"], encoding="utf-8") as f:
        evento = json.load(f)
    return [c for c in evento.get("commits", []) if c.get("distinct", True)]


def montar_entrada(commits: list[dict]) -> str:
    partes = []
    for c in commits:
        diff = git("show", "--format=", "--patch", c["id"], "--", ".", *IGNORAR)
        partes.append(f"### Commit {c['id'][:7]} — {c['message']}\n{diff}")
    entrada = "\n\n".join(partes)
    if len(entrada) > LIMITE_DIFF:
        # Não corta o diff no meio: manda só o resumo por arquivo e avisa o Claude
        resumo = "\n\n".join(
            f"### Commit {c['id'][:7]} — {c['message']}\n"
            + git("show", "--format=", "--stat", c["id"], "--", ".", *IGNORAR)
            for c in commits
        )
        entrada = "(Diff grande demais; abaixo só a lista de arquivos alterados.)\n\n" + resumo
    return entrada


def resumir_com_claude(entrada: str) -> str:
    client = anthropic.Anthropic()
    resposta = client.beta.messages.create(
        model="claude-opus-5-5",
        max_tokens=16000,
        output_config={"effort": "medium"},
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        system=PROMPT_SISTEMA,
        messages=[{"role": "user", "content": entrada}],
    )
    if resposta.stop_reason == "refusal":
        raise RuntimeError(f"Claude recusou o resumo (request-id {resposta._request_id}).")
    return "".join(b.text for b in resposta.content if b.type == "text").strip()


def comentar(chave: str, texto: str) -> None:
    url = f"{os.environ['JIRA_BASE_URL'].rstrip('/')}/rest/api/2/issue/{chave}/comment"
    r = requests.post(
        url,
        json={"body": texto},
        auth=(os.environ["JIRA_EMAIL"], os.environ["JIRA_API_TOKEN"]),
        timeout=30,
    )
    if r.status_code == 404:
        print(f"⚠️ Card {chave} não encontrado no Jira (ou sem permissão); ignorado.")
        return
    r.raise_for_status()
    print(f"✅ Comentário publicado em {chave}.")


def main() -> int:
    faltando = [v for v in ("ANTHROPIC_API_KEY", "JIRA_BASE_URL", "JIRA_EMAIL", "JIRA_API_TOKEN") if not os.environ.get(v)]
    if faltando:
        print(f"⚠️ Secrets não configurados no GitHub: {', '.join(faltando)}. Comentário no Jira ignorado.")
        return 0

    commits = commits_do_push()
    if not commits:
        print("Nenhum commit novo neste push.")
        return 0

    branch = os.environ.get("GITHUB_REF_NAME", "")
    texto_busca = branch + " " + " ".join(c["message"] for c in commits)
    chaves = sorted({k.upper() for k in CHAVE_JIRA.findall(texto_busca)})
    if not chaves:
        print(f"Nenhum ticket {PROJETO_JIRA}-<número> na branch ou nos commits; nada a comentar.")
        return 0

    resumo = resumir_com_claude(montar_entrada(commits))
    repo_url = f"{os.environ['GITHUB_SERVER_URL']}/{os.environ['GITHUB_REPOSITORY']}"
    lista = "\n".join(f"* [{c['id'][:7]}|{repo_url}/commit/{c['id']}] {c['message'].splitlines()[0]}" for c in commits)
    texto = (
        f"🤖 *Resumo automático (Claude) — push na branch {branch}*\n\n{resumo}\n\n"
        f"----\n*Commits:*\n{lista}"
    )
    for chave in chaves:
        comentar(chave, texto)
    return 0


if __name__ == "__main__":
    sys.exit(main())
