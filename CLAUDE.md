# Automação de Agendamentos Atrasados

Visão geral, instalação e uso: ver README.md. Fluxo: `extrator.py` (portal, Playwright) → `atrasadas.py`
(consolida) → `disparador.py` (filtra, cruza contatos e envia pelo Outlook, usando `teste_email.py`).
Os três terminam em `disparador.main()` e completam a etapa que faltar (sem consolidado → consolida;
sem downloads → extrai do portal). Imports entre eles são tardios (dentro da função) para evitar ciclo. Windows + Outlook + Python 3.11+, venv em `.venv`.

## Comandos

```powershell
pip install -r requirements.txt; playwright install chromium   # setup (versões fixadas)
.venv\Scripts\python -m py_compile extrator.py atrasadas.py teste_email.py regras.py config.py disparador.py   # checagem de sintaxe
.venv\Scripts\python disparador.py            # SIMULAÇÃO (não envia); mas sem consolidado, extrai do portal
.venv\Scripts\python disparador.py --teste    # envia todos os e-mails para o EMAIL_TESTE (assunto "[TESTE → contato]")
```

- Não há suíte de testes. Para validar regras, chamar `regras.py` direto com `data_referencia` fixa
  (ex.: `dias_uteis_atraso(pd.Series(['16/11/2026']), pd.Timestamp('2026-11-20'))` → 3, pois 20/11 é feriado).
- **Não rodar** `extrator.py`, `atrasadas.py`, `teste_email.py` nem `disparador.py --teste/--enviar` sem pedido explícito: enviam e-mail real
  (`mail.Send()`) e/ou logam no portal com a senha do `.env`.
- Para testar o fluxo sem portal/Outlook: `unittest.mock.patch` em `extrator.baixar_planilhas`,
  `disparador.enviar_email` e nas pastas (`atrasadas.PASTA_DOWNLOADS`, `teste_email.PASTA_CONSOLIDADO`).
- Importar `teste_email` (ou `disparador`, que o importa) exige `EMAIL_REMETENTE`/`EMAIL_TESTE` no `.env`
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

- Pastas de saída só em `config.py` (`PASTA_DOWNLOADS`, `PASTA_CONSOLIDADO`, `PASTA_ENVIOS`, `PASTA_ERROS`), dentro de
  `atrasadas/1_downloads`, `2_consolidado`, `3_envios` e `4_erros`. Não escrever caminhos à mão nos scripts nem salvar na raiz.
- Log aberto no Excel → salva `Log_Envios_<data>.csv` à parte; nunca deixar o envio terminar sem registrar.
- Consolidado aberto no Excel → `PermissionError`; o código salva cópia com data/hora e lê o mais recente
  (`Atrasadas_Consolidado*.xlsx`), ignorando arquivos `~$`.
- Cruzamento nota × base de contatos (`disparador.py`): **só pelo CNPJ**, nunca pelo nome. Coluna `Emitente`
  do portal (mantida pelo `atrasadas.py`) × `CNPJ FORNECEDOR` da base, comparando só dígitos com `zfill(14)`
  (a base mistura formatos e perde o zero à esquerda). 1º CNPJ exato, 2º raiz (8 dígitos = outra filial).
  Nome falhava (portal corta em ~40 caracteres, base abrevia/renomeia) e fuzzy já mandou HELLA/BORGWARNER
  para a ZF. CNPJ/raiz com Razões Sociais diferentes na base = ambíguo → pendente com aviso das linhas.
  Na dúvida, a nota vai para `Fornecedores_Pendentes.xlsx`, nunca para outro fornecedor.
  A base só fornece o e-mail: o nome no assunto/corpo é a Razão Social do portal (a da base vem colada/anotada).
- Consolidado antigo sem `Emitente` → o disparador reconsolida de `1_downloads` sozinho.
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
