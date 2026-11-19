from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
import os, time

from config import IGNORAR_CERTIFICADO, PASTA_DOWNLOADS, PASTA_ERROS, obrigatorio

# CONFIGURAÇÕES (credenciais vêm do arquivo .env — nunca escreva senhas no código)
EMAIL = obrigatorio("BRA_EMAIL")
SENHA = obrigatorio("BRA_SENHA")
URL_BASE = "https://app.brautoparts.com.br"

FILIAIS = [
    "Araguaina", "Bauru", "Belem", "Belem P", "Belo Horizonte", "Blumenau", "Brasilia",
    "Campinas", "Campo Grande", "Campo Grande D", "Cascavel", "Chapeco", "Contagem",
    "Cuiaba C", "Cuiaba D", "Cuiaba P", "Curitiba", "Florianopolis", "Fortaleza D",
    "Fortaleza P", "Goiania", "Joao Pessoa", "Joinville", "Juiz de Fora", "Londrina",
    "Maceio", "Manaus", "Marilia", "Maringa", "Natal", "Passo Fundo", "Pelotas", "Porto Alegre",
    "Porto Velho", "Pres. Prudente", "Recife", "Recife P", "Ribeirao Preto",
    "Rio de Janeiro", "S J Rio Preto D", "S J Rio Preto P", "Salvador C", "Salvador D", "Salvador P",
    "Santa Maria", "Santos", "Sao Jose dos Campos", "Sao Luis", "Sao Paulo",
    "Sao Paulo Leste", "Sao Paulo Sul", "Teresina", "Uberlandia", "Varginha",
    "Vitoria", "Vitoria Conquista"
]

def preenchimento_automatico(page, selector: str, value: str, tries: int = 3):
    loc = page.locator(selector)
    for i in range(1, tries + 1):
        loc.click(timeout=8000)
        loc.fill("", timeout=8000)
        loc.type(value, delay=25)
        if i < tries:
            time.sleep(0.2)

def acessar_portal_e_logar(page):
    print("🔗 Acessando portal…")
    try:
        page.goto(URL_BASE, wait_until="networkidle", timeout=70000)
    except PWTimeout:
        page.goto(URL_BASE, wait_until="domcontentloaded", timeout=70000)

    preenchimento_automatico(page, "#email", EMAIL)
    preenchimento_automatico(page, "#senha", SENHA)
    page.locator("#kt_sign_in_submit").click()

    try:
        page.wait_for_selector('a:has-text("Filiais"), a:has-text("Logística"), nav', timeout=70000, state="attached")
    except PWTimeout:
        raise RuntimeError("❌ Falha no login. Verifique as credenciais.")
    print("✅ Login realizado!")

def baixar_planilhas() -> int:
    """Etapa 1: login no portal e download da planilha de cada filial em 1_downloads. Retorna quantas baixou."""
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=False, slow_mo=150)
        context = browser.new_context(accept_downloads=True, ignore_https_errors=IGNORAR_CERTIFICADO)
        page = context.new_page()

        acessar_portal_e_logar(page)

        page.wait_for_selector('a[href="/srm"]', timeout=70000)
        page.locator('a[href="/srm"]').click()
        
        card_ag = page.locator('a.card[href="/srm/agendamentos"]')
        card_ag.wait_for(state="visible", timeout=70000)
        card_ag.click()
        
        aba_atrasadas = page.locator('a.nav-link:has-text("Atrasadas")').first
        aba_atrasadas.wait_for(state="visible", timeout=40000)
        aba_atrasadas.click()

        # Remove as planilhas da execução anterior para não duplicar dados no consolidado
        for antigo in PASTA_DOWNLOADS.glob("Agendamento*.xlsx"):
            antigo.unlink()

        botao_exportar = page.locator('button[data-bs-original-title="Exportar para o Excel"]').first
        msg_vazio = page.locator("text=Nenhum agendamento").first

        baixadas, sem_dados, falhas = [], [], []
        for filial in FILIAIS:
            print(f"📂 Baixando: {filial}")
            try:
                active_tab = page.locator('div.tab-pane.active.show').first
                active_tab.locator('div.mud-select:has(label:has-text("Filial")) div.mud-input').first.click()
                page.wait_for_selector('div.mud-list-item', timeout=40000)
                page.locator(f'div.mud-list-item:has-text("{filial}")').first.click(timeout=10000)
            except Exception as e:
                print(f"   ❌ Filial não encontrada na lista: {e.__class__.__name__}")
                falhas.append(filial)
                page.keyboard.press("Escape")
                continue

            page.get_by_role("button", name="Pesquisar").click()
            page.wait_for_timeout(2000)

            # Espera aparecer o botão de exportar OU a mensagem de "nenhum agendamento"
            try:
                botao_exportar.or_(msg_vazio).first.wait_for(state="visible", timeout=20000)
            except PWTimeout:
                pass

            if msg_vazio.is_visible() or not botao_exportar.is_visible():
                print("   ⚪ Sem registros (botão de exportar não apareceu)")
                sem_dados.append(filial)
                page.screenshot(path=str(PASTA_ERROS / f"{filial}.png"))
                continue

            try:
                with page.expect_download(timeout=40000) as download_info:
                    botao_exportar.click(timeout=15000)
                destino = PASTA_DOWNLOADS / f"Agendamento_{filial}.xlsx"
                download_info.value.save_as(str(destino))
                baixadas.append(filial)
            except Exception as e:
                print(f"   ❌ Erro ao baixar {filial}: {e.__class__.__name__}")
                falhas.append(filial)

        browser.close()

    print(f"\n📋 Resumo: {len(baixadas)} baixadas | {len(sem_dados)} sem registros | {len(falhas)} com erro")
    if sem_dados:
        print(f"   Sem registros: {', '.join(sem_dados)}  (prints em {PASTA_ERROS})")
    if falhas:
        print(f"   Com erro: {', '.join(falhas)}")
    return len(baixadas)


if __name__ == "__main__":
    # Fluxo completo: extrai do portal, consolida e segue para o disparo.
    # Aceita os mesmos parâmetros do disparador: [AAAA-MM-DD] [--teste | --enviar]
    import disparador
    disparador.main(extrair=True)