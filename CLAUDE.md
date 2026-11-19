# Automação de Agendamentos Atrasados

Visão geral, instalação e uso: ver README.md. Fluxo: `extrator.py` (portal, Playwright) → `atrasadas.py`
(consolida) → `teste_email.py` (filtra e envia pelo Outlook). Windows + Outlook + Python 3.11+, venv em `.venv`.

## Comandos

```powershell
pip install -r requirements.txt; playwright install chromium   # setup (versões fixadas)
.venv\Scripts\python -m py_compile extrator.py atrasadas.py teste_email.py regras.py config.py disparador.py   # checagem de sintaxe
.venv\Scripts\python disparador.py            # fase 2 (todos os fornecedores): SIMULAÇÃO, seguro de rodar
```

- Não há suíte de testes. Para validar regras, chamar `regras.py` direto com `data_referencia` fixa
  (ex.: `dias_uteis_atraso(pd.Series(['16/11/2026']), pd.Timestamp('2026-11-20'))` → 3, pois 20/11 é feriado).
- **Não rodar** `extrator.py`, `atrasadas.py`, `teste_email.py` nem `disparador.py --enviar` sem pedido explícito: enviam e-mail real
  (`mail.Send()`) e/ou logam no portal com a senha do `.env`.
- Importar `teste_email` (ou `atrasadas`, que o importa) exige `EMAIL_REMETENTE`/`EMAIL_TESTE` no `.env`
  (`SystemExit` no import). Para testes isolados, importar só `regras`.

## Regras de negócio (não alterar sem pedido)

- Ficam só em `regras.py`; constantes `DIAS_UTEIS_MINIMO=4`, `DIAS_CORRIDOS_CRITICO=60`, `DIA_ENVIO_SEMANAL=1` (terça).
- Dias exibidos e alerta crítico (>60) = dias **corridos** desde a Emissão (HOJE()-Emissão).
- A única regra em dias **úteis** é o mínimo para cobrar: 4º dia útil após a Agenda (agenda exclusive,
  hoje inclusive; sem fins de semana, feriados nacionais + Carnaval/Corpus Christi + `FERIADOS_EXTRAS`).
- `filtrar_fornecedor` recalcula `Dias` na data de referência; não confiar no valor gravado no Excel.
- Envio futuro: semanal, às terças.

## Padrões de código

- Código, nomes, mensagens e commits em português; mensagens de console com emoji (`✅ ⚠️ ❌ 📂`).
- `teste_email.py` é a **fonte única** do layout/envio do e-mail e da aplicação das regras (`notas_a_cobrar`);
  `atrasadas.py` e `disparador.py` só reutilizam (não duplicar envio, leitura de Excel nem regras).
- Funções de data aceitam `data_referencia: pd.Timestamp | None` (padrão hoje) para serem reproduzíveis.
- Datas das planilhas são dd/mm/aaaa: sempre `pd.to_datetime(..., dayfirst=True, errors="coerce")`.
- pandas 3.x: usar `.to_numpy(dtype="datetime64[D]")`, não `.values.astype(...)`.
- Segredos e configs só via `config.obrigatorio()`/`config.opcional()`, lendo o `.env`.

## Armadilhas

- Consolidado aberto no Excel → `PermissionError`; o código salva cópia com data/hora e lê o mais recente
  (`Atrasadas_Consolidado*.xlsx`), ignorando arquivos `~$`.
- Cruzamento nota × base de contatos (`disparador.py`): comparar só com `nome_chave()` (sem LTDA, BRASIL,
  AUTOMOTIVE…) e `token_sort_ratio` ≥ 90. `token_set_ratio` no nome cheio mandava HELLA/BORGWARNER para o
  e-mail da ZF. Na dúvida, a nota vai para `Fornecedores_Pendentes.xlsx`, nunca para outro fornecedor.
- Base de contatos (hoje `baseagdteste.xlsx`, só teste): vários e-mails por fornecedor, na mesma célula ou em
  linhas repetidas, são juntados em um único e-mail com todos no "Para" (não usar `drop_duplicates`).
- "Faltam notas do fornecedor no e-mail" costuma ser a regra do 4º dia útil (Agenda recente), não bug:
  a coluna `Dias` do Excel conta desde a Emissão. Conferir com `dias_uteis_atraso(df["Agenda"])`.
- `IGNORAR_CERTIFICADO` deve ficar `0` (desligar expõe a senha do portal a interceptação).

## Git + Jira (projeto A122)

- `A122` é a chave do projeto; o número após o hífen identifica cada ticket (`A122-11`, `A122-14`, ...).
- Toda branch e todo commit citam o ticket a que se referem. Um commit pode citar mais de um
  (`A122-11 A122-14: ...`) e o resumo é comentado em cada um.
  - Branch: `feature/A122-11-descricao-curta` (ou `fix/...`). Nunca commitar direto na `main`.
  - Commit: `A122-11: descrição no imperativo, em português`.
- Se não houver card claro para a mudança, perguntar qual usar antes de criar a branch.
- O push dispara `.github/workflows/jira-claude.yml`, que comenta o resumo no card.
  Os cards mudam de coluna pelo Jira Automation (branch → Em andamento, PR → Em revisão, merge → Concluído).

## Segurança

- Nunca versionar `.env`, planilhas (`*.xlsx`), prints (`*.png`), `atrasadas/` nem `base fornecedores/` (LGPD).
- Conferir `git status` / `git diff --cached` antes de cada commit.

## Manutenção deste arquivo

- Ao descobrir um comando, convenção, restrição ou armadilha nova, registrar aqui de forma enxuta.
  Detalhes de uso e explicações longas vão para o README.md.
