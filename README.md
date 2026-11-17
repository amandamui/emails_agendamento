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
                   - calcula "Dias" = hoje − data de emissão
                   - salva atrasadas/Atrasadas_Consolidado.xlsx
                3. Filtra o fornecedor e envia o e-mail pelo Outlook
```

| Arquivo | Função |
|---|---|
| `extrator.py` | Fluxo completo: login, download por filial e, no final, chama o `atrasadas.py` |
| `atrasadas.py` | Consolida as planilhas baixadas, trata as colunas e envia o e-mail (fase 1: um fornecedor) |
| `teste_email.py` | Envia um e-mail de teste com layout novo a partir do consolidado já gerado, sem abrir o portal |
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
| `EMAIL_REMETENTE` | Conta do Outlook que envia (teste_email.py) |
| `EMAIL_TESTE` | Para quem vai o e-mail de teste |
| `IGNORAR_CERTIFICADO` | Deixe `0`. Use `1` só se o proxy da empresa bloquear o HTTPS |

## Uso

```powershell
# Fluxo completo: baixa, consolida e envia o e-mail
python extrator.py

# Só consolidar/enviar a partir das planilhas que já estão em atrasadas/
python atrasadas.py

# Só testar o layout do e-mail a partir do consolidado existente
python teste_email.py
```

O fornecedor da fase de teste é definido em `FORNECEDOR_TESTE` (`atrasadas.py` e `teste_email.py`).

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
- [ ] Agendar a execução automática (Agendador de Tarefas do Windows)
- [ ] Usar uma conta de serviço no portal em vez de um usuário pessoal
