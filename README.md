# Automação de Agendamentos Atrasados

Robô que baixa do portal **app.brautoparts.com.br** (SRM → Agendamentos → aba *Atrasadas*) as notas agendadas e não entregues de todas as filiais, consolida tudo em uma planilha e envia por e-mail (Outlook) a cobrança para o fornecedor.

## Como funciona

```
extrator.py ──► 1. Login no portal e download da planilha de cada filial
                   (salva em atrasadas/1_downloads/Agendamento_<Filial>.xlsx)
     │
     ▼
atrasadas.py ─► 2. Consolida e trata a planilha
                   - remove colunas internas (CNPJ, valor, protocolo, usuários…)
                   - calcula "Dias" = dias corridos desde a Emissão (HOJE() − Emissão)
                   - salva atrasadas/2_consolidado/Atrasadas_Consolidado.xlsx
     │
     ▼
disparador.py ► 3. Filtra as notas a partir do 4º dia útil após a Agenda, cruza com a base
                   de contatos pelo CNPJ (Emitente) e dispara um e-mail por fornecedor
                   (layout do teste_email.py)
```

Qualquer um dos três scripts termina no disparo e completa sozinho a etapa que faltar.

**Cache:** se `atrasadas/1_downloads` já tem planilhas, elas são reaproveitadas (o portal não é aberto), não
importa de quando sejam; a data do download aparece na tela. Se o consolidado for mais novo que as planilhas (já
foram lidas), ele também é reaproveitado; senão, é refeito. Para baixar de novo do portal, use `--atualizar`.
O extrator grava `_extracao_concluida.txt` ao terminar todas as filiais; sem ele, o disparador avisa que a
última extração foi interrompida e podem faltar filiais.

| Arquivo | Função |
|---|---|
| `extrator.py` | Extrai do portal (se `1_downloads` estiver vazia), consolida e dispara |
| `atrasadas.py` | Reconsolida as planilhas de `1_downloads` (extrai antes, se estiver vazia) e dispara |
| `teste_email.py` | **Fonte única do e-mail**: layout, regra das 72h e envio. Rodando direto, envia a partir do consolidado já gerado, sem abrir o portal |
| `disparador.py` | **Fase 2**: cruza o consolidado com `base fornecedores/` e cobra todos os fornecedores com e-mail. Sem parâmetro só simula; `--teste` manda tudo para o `EMAIL_TESTE`; `--enviar` manda para os fornecedores. Notas sem contato vão para `atrasadas/3_envios/Fornecedores_Pendentes.xlsx` e os envios para `atrasadas/3_envios/Log_Envios.csv` |
| `regras.py` | **Regras de negócio**: dias exibidos em dias corridos desde a Emissão; só o mínimo para cobrar usa dias úteis após a Agenda (4º dia útil, sem fins de semana e feriados nacionais), alerta crítico (mais de 60 dias corridos desde a Emissão) |
| `config.py` | Lê as configurações e senhas do arquivo `.env` |
| `.env.example` | Modelo do `.env` (sem valores reais) |

## Requisitos

- Windows com **Microsoft Outlook** instalado e logado (o envio é feito pelo Outlook da máquina)
- **Google Chrome** instalado
- Python 3.11+

## Instalação

```powershell
# 1. Criar e ativar o ambiente virtual
python -m venv .venv
.venv\Scripts\Activate.ps1

# 2. Instalar as dependências
pip install -r requirements.txt
playwright install chromium

# 3. Configurar as credenciais
copy .env.example .env
notepad .env
```

Preencha o `.env`:

| Variável | Descrição |
|---|---|
| `BRA_EMAIL` / `BRA_SENHA` | Login do portal |
| `EMAIL_REMETENTE` | Conta do Outlook que envia o e-mail |
| `EMAIL_TESTE` | Para quem vai o e-mail de teste |
| `FORNECEDOR_TESTE` | Fornecedor da fase de teste (padrão: `Gates`) |
| `IGNORAR_CERTIFICADO` | Deixe `0`. Use `1` só se o proxy da empresa bloquear o HTTPS |

## Uso

```powershell
# Os três aceitam: [AAAA-MM-DD] [--teste | --enviar] [--atualizar]
#   sem parâmetro → SIMULAÇÃO (só lista quem receberia, não envia nada)
#   --teste       → envia todos os e-mails para o EMAIL_TESTE do .env
#   --enviar      → envia para os fornecedores
#   --atualizar   → ignora as planilhas já baixadas e baixa tudo de novo do portal

python extrator.py --teste       # usa as planilhas já baixadas (ou baixa, se não houver) e testa no seu e-mail
python extrator.py --atualizar   # baixa tudo de novo do portal
python atrasadas.py              # reconsolida o que já foi baixado e simula
python disparador.py             # usa o consolidado existente (gera o que faltar) e simula
python disparador.py 2026-11-24  # simula o envio de outra data (ex.: a próxima terça)

# Teste antigo (fase 1): um fornecedor só, a partir do consolidado existente
python teste_email.py
```

O fornecedor da fase de teste é definido em `FORNECEDOR_TESTE` no `.env` (padrão: `Gates`).

> ⚠️ O e-mail é enviado direto (`mail.Send()`). Para revisar antes de enviar, troque por `mail.Display()`.

### Saídas

- `atrasadas/1_downloads/Agendamento_<Filial>.xlsx` — planilha bruta de cada filial (apagadas a cada nova extração); `_extracao_concluida.txt` marca que a última extração terminou
- `atrasadas/2_consolidado/Atrasadas_Consolidado.xlsx` — planilha final tratada
- `atrasadas/3_envios/Log_Envios.csv` — histórico de envios do `disparador.py` (acumulado)
- `atrasadas/3_envios/Fornecedores_Pendentes.xlsx` — notas a cobrar cujo CNPJ (coluna `Emitente`) não está
  na base, ou está com Razões Sociais diferentes (conflito avisado no console com as linhas da base).
  O cruzamento é só pelo CNPJ: CNPJ exato e, se não houver, a raiz (8 primeiros dígitos, outra filial).
  Também ficam pendentes os fornecedores com e-mail inválido na base (ex.: dois e-mails grudados, sem
  separador). Na base, separe os e-mails por `;`, `,`, `/`, espaço ou quebra de linha; todos vão no "Para".
- `atrasadas/4_erros/<Filial>.png` — print da tela das filiais em que o botão de exportar não apareceu
- `atrasadas/5_historico/Atrasadas_AAAA-MM-DD.xlsx` — foto do consolidado gravada pelo `disparador.py` (uma por dia), base da devolutiva
- `atrasadas/3_envios/Devolutiva_AAAA-MM-DD.html` — prévia do balanço semanal (`python devolutiva.py`, sem enviar)

**Devolutiva semanal (terça, depois do disparador):** `python devolutiva.py` (simulação, gera a prévia), `--teste` (envia para o `EMAIL_TESTE`) ou `--enviar` (envia para `EMAIL_DEVOLUTIVA` do `.env`). Compara as notas cobradas na foto da semana passada com a de hoje: resolvidas, que continuam pendentes e novas.

## Problemas comuns

| Sintoma | Causa / solução |
|---|---|
| `PermissionError ... Atrasadas_Consolidado.xlsx` | A planilha está aberta no Excel. Feche-a. O robô salva uma cópia com data/hora no nome e o `teste_email.py` lê o arquivo mais recente. |
| `Variável 'BRA_SENHA' não configurada` | Falta o arquivo `.env` ou a variável dentro dele. |
| `⚪ Sem registros` em uma filial | A filial não tem notas atrasadas, ou a página demorou. Confira o print em `atrasadas/4_erros/`. |
| Erro ao enviar e-mail | O Outlook precisa estar aberto e configurado na máquina. |
| `Import "playwright" could not be resolved` no VS Code | Selecione o interpretador do `.venv` (`Ctrl+Shift+P` → *Python: Select Interpreter*). |

## Integração GitHub + Jira + Claude

Toda branch e commit citam o card do Jira (projeto **A122**):

```powershell
git checkout -b feature/A122-11-regra-dias-uteis
git commit -m "A122-11: Cobra a partir do 4º dia útil após a Agenda"
git push -u origin feature/A122-11-regra-dias-uteis
```

| O que acontece | Quem faz |
|---|---|
| Commits, branches e PRs aparecem no card (painel *Desenvolvimento*) | App **GitHub for Jira** |
| Card muda de coluna (branch → Em andamento; PR → Em revisão; merge → Concluído) | **Jira Automation** |
| Claude resume o push em português e comenta no card | GitHub Action `.github/workflows/jira-claude.yml` |

Secrets necessários no GitHub (*Settings → Secrets and variables → Actions*):
`ANTHROPIC_API_KEY`, `JIRA_BASE_URL` (ex.: `https://suaempresa.atlassian.net`), `JIRA_EMAIL`, `JIRA_API_TOKEN`.

Planilhas, prints, `.env` e a pasta `atrasadas/` nunca são enviados ao Claude nem ao Jira.

## Segurança

Este projeto lida com **senha do portal** e com **dados de negócio** (CNPJs, notas fiscais, e-mails de transportadoras). Regras:

1. **Nunca escreva senhas no código.** Elas ficam só no `.env`, que está no `.gitignore`.
2. **Nunca versione as planilhas nem os prints** (`atrasadas/`, `*.xlsx`, `*.png`): contêm dados de terceiros (LGPD).
3. **Mantenha o repositório privado** no GitHub.
4. Antes de cada commit, confira o que vai subir:
   ```powershell
   git status
   git diff --cached
   ```
   Se aparecer `.env`, `.xlsx` ou a pasta `atrasadas/`, **não faça o commit**.
5. Ative no GitHub: *Settings → Code security → Secret scanning* e *Push protection*.
6. A verificação de certificado HTTPS fica **ligada** por padrão. Desligá-la (`IGNORAR_CERTIFICADO=1`) permite que alguém na rede intercepte a senha do portal.
7. **Se uma senha for parar no GitHub**, apagar o arquivo não resolve, porque ela continua no histórico. Troque a senha no portal imediatamente.

## Próximos passos

- [ ] Agendar a execução automática semanal, **toda terça-feira** (Agendador de Tarefas do Windows). `regras.DIA_ENVIO_SEMANAL` já reserva essa configuração
- [ ] Cadastrar feriados estaduais/municipais das filiais em `regras.FERIADOS_EXTRAS`
- [ ] Usar uma conta de serviço no portal em vez de um usuário pessoal
