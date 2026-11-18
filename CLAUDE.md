# Automação de Agendamentos Atrasados

Visão geral, instalação e uso: ver README.md.

## Git + Jira (projeto A122)

- `A122` é a chave do projeto; o número após o hífen identifica cada ticket (`A122-11`, `A122-14`, ...).
- Toda branch e todo commit citam o ticket a que se referem. Um commit pode citar mais de um
  (`A122-11 A122-14: ...`) e o resumo é comentado em cada um.
  - Branch: `feature/A122-11-descricao-curta` (ou `fix/...`). Nunca commitar direto na `main`.
  - Commit: `A122-11: descrição no imperativo, em português`.
- Se não houver card claro para a mudança, perguntar qual usar antes de criar a branch.
- O push dispara `.github/workflows/jira-claude.yml`, que comenta o resumo no card.
  Os cards mudam de coluna pelo Jira Automation (branch → Em andamento, PR → Em revisão, merge → Concluído).

## Regras de negócio

- Ficam só em `regras.py`. Atraso é em dias corridos desde a Agenda; a única regra em dias úteis
  é o mínimo para cobrar (4º dia útil). Envio futuro: semanal, às terças.

## Segurança

- Nunca versionar `.env`, planilhas (`*.xlsx`), prints (`*.png`) nem a pasta `atrasadas/` (LGPD).
- O e-mail sai direto pelo Outlook (`mail.Send()`): não rodar `teste_email.py`/`atrasadas.py` sem pedido explícito.
