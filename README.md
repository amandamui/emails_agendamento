# Automação de Agendamentos Atrasados

Robô que baixa do portal **app.brautoparts.com.br** (SRM → Agendamentos → aba *Atrasadas*) as notas agendadas e não entregues de todas as filiais, consolida tudo em uma planilha e envia por e-mail (Outlook) a cobrança para o fornecedor.

## Como funciona

```
extrator.py ──► 1. Login no portal e download da planilha de cada filial
                   (salva em atrasadas/Agendamento_<Filial>.xlsx)
     │
     ▼
atrasadas.py ─► 2. Consolida e trata a planilha
                   - remove colunas internas (CNPJ, valor, protocolo, usuários…)
                   - calcula "Dias" = dias corridos desde a Emissão (HOJE() − Emissão)
                   - salva atrasadas/Atrasadas_Consolidado.xlsx
                3. Filtra o fornecedor (só notas a partir do 4º dia útil após a Agenda)
                   e envia o e-mail pelo Outlook, usando o layout do teste_email.py
```

| Arquivo | Função |
|---|---|
| `extrator.py` | Fluxo completo: login, download por filial e, no final, chama o `atrasadas.py` |
| `atrasadas.py` | Consolida as planilhas baixadas, trata as colunas e chama o envio do e-mail (fase 1: um fornecedor) |
| `teste_email.py` | **Fonte única do e-mail**: layout, regra das 72h e envio. Rodando direto, envia a partir do consolidado já gerado, sem abrir o portal |
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
# Fluxo completo: baixa, consolida e envia o e-mail
python extrator.py

# Só consolidar/enviar a partir das planilhas que já estão em atrasadas/
python atrasadas.py

# Só testar o layout do e-mail a partir do consolidado existente
python teste_email.py

# Simular o envio de outra data (ex.: a próxima terça-feira)
python teste_email.py 2026-11-24
```

O fornecedor da fase de teste é definido em `FORNECEDOR_TESTE` no `.env` (padrão: `Gates`).

> ⚠️ O e-mail é enviado direto (`mail.Send()`). Para revisar antes de enviar, troque por `mail.Display()`.

### Saídas

- `atrasadas/Agendamento_<Filial>.xlsx` — planilha bruta de cada filial (apagadas a cada nova execução)
- `atrasadas/Atrasadas_Consolidado.xlsx` — planilha final tratada
- `atrasadas/erros/<Filial>.png` — print da tela das filiais em que o botão de exportar não apareceu

## Problemas comuns

| Sintoma | Causa / solução |
|---|---|
| `PermissionError ... Atrasadas_Consolidado.xlsx` | A planilha está aberta no Excel. Feche-a. O robô salva uma cópia com data/hora no nome e o `teste_email.py` lê o arquivo mais recente. |
| `Variável 'BRA_SENHA' não configurada` | Falta o arquivo `.env` ou a variável dentro dele. |
| `⚪ Sem registros` em uma filial | A filial não tem notas atrasadas, ou a página demorou. Confira o print em `atrasadas/erros/`. |
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

- [ ] Enviar para todos os fornecedores, com e-mail de contato por fornecedor
- [ ] Agendar a execução automática semanal, **toda terça-feira** (Agendador de Tarefas do Windows). `regras.DIA_ENVIO_SEMANAL` já reserva essa configuração
- [ ] Cadastrar feriados estaduais/municipais das filiais em `regras.FERIADOS_EXTRAS`
- [ ] Usar uma conta de serviço no portal em vez de um usuário pessoal
