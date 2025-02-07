import asyncio
from playwright.async_api import async_playwright
from datetime import datetime
import os

async def scrape_site():
    # Defina o diretório de downloads
    download_dir = os.path.expanduser('~') + '/Downloads/KM'  # Caminho para a pasta de Downloads

    # Configuração do navegador
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            headless=False
        )  # Caminho para o seu Chrome
        context = await browser.new_context(
            accept_downloads=True  # Habilita o comportamento de download
        )
        page = await context.new_page()

        # Configura a captura do download
        download_path = None
        page.on('download', lambda download: download.save_as(os.path.join(download_dir, download.suggested_filename)))
        
        print("Acessando o site...")
        # Acesse o site
        await page.goto("https://go.denox.com.br/#trans/reports/tripReport/TripReport")
        await page.wait_for_load_state("load")  # Aguarda o carregamento inicial da página

        print("Preenchendo o email...")
        # Preencher o campo de email
        email_selector = "xpath=/html/body/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div[1]/div[1]/div/input"
        await page.wait_for_selector(email_selector, timeout=10000)
        await page.locator(email_selector).fill("kaique.pimentel@etp-transparana.com.br")

        print("Preenchendo a senha...")
        # Preencher o campo de senha
        password_selector = "xpath=/html/body/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div[1]/div[2]/div/input"
        await page.wait_for_selector(password_selector, timeout=10000)
        await page.locator(password_selector).fill("Kaique131029")

        print("Fazendo login...")
        # Clicar no botão de login
        login_button_selector = "xpath=/html/body/div/div[2]/div[1]/div[1]/div/div/div[4]/div/div[1]/div[3]/button"
        await page.wait_for_selector(login_button_selector, timeout=10000)
        await page.locator(login_button_selector).click()
        await page.wait_for_load_state("networkidle")

        print("Abrindo filtros...")
        # Clique para abrir o campo dos filtros
        filter_button_selector = "xpath=/html/body/div/div[2]/div[1]/div[2]/div/div/div[1]/div[1]/div[1]/div/div[2]/div/div/div"
        await page.wait_for_selector(filter_button_selector, timeout=10000)
        await page.click(filter_button_selector)

        print("Preenchendo período inicial e final...")
        # Preencher o filtro do período inicial
        start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        start_date_selector = "xpath=/html/body/div/div[2]/div[1]/div[2]/div/div/div[1]/div[1]/div[1]/div/div[2]/div/div[2]/div/div/div/div/div[2]/div[2]/div[1]/div[1]/div/div/input[1]"
        await page.wait_for_selector(start_date_selector, timeout=10000)
        await page.locator(start_date_selector).fill(start_date.strftime("%d/%m/%Y %H:%M"))

        # Preencher o filtro do período final
        end_date = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0)
        end_date_selector = "xpath=/html/body/div/div[2]/div[1]/div[2]/div/div/div[1]/div[1]/div[1]/div/div[2]/div/div[2]/div/div/div/div/div[2]/div[2]/div[1]/div[1]/div/div/input[2]"
        await page.wait_for_selector(end_date_selector, timeout=10000)
        await page.locator(end_date_selector).fill(end_date.strftime("%d/%m/%Y %H:%M"))

        # Clicar no botão "Filtrar"
        print("Aplicando filtros...")
        filter_button_selector = "xpath=/html/body/div/div[2]/div[1]/div[2]/div/div/div[1]/div[1]/div[1]/div/div[2]/div/div[2]/div/div/div/div/div[2]/div[1]/div[2]/a"
        await page.wait_for_selector(filter_button_selector, timeout=10000)
        await page.click(filter_button_selector)

        await asyncio.sleep(2)  # Aguarda a aplicação dos filtros

        print("Iniciando download do CSV...")
        # Clicar no botão de download
        download_button_selector = "xpath=/html/body/div/div[2]/div[1]/div[2]/div/div/div[1]/div[1]/div[1]/div/div[1]/div/div[2]/div[1]/a/i"
        await page.wait_for_selector(download_button_selector, timeout=10000)
        await page.click(download_button_selector)

        # Localizar o botão "Exportar CSV" e clicar usando o método get_by_text
        download_csv_selector = page.get_by_text("Exportar CSV")
        await download_csv_selector.click()

        # Aguardar o carregamento da página após o clique no botão "Exportar CSV"
        await page.wait_for_load_state("load", timeout=15000)

        # Localizar o botão "Sim" de forma mais específica
        confirm_download_selector = page.locator("button.btn.btn-primary.btn-focus:has-text('Sim')")
        await confirm_download_selector.click()

        # Aguardar 1 minuto antes de clicar no botão para salvar o download
        await asyncio.sleep(30)

        # Clicar no botão para salvar o download utilizando o XPath corretamente
        save_download_selector = page.locator("i.text-info.icon-download-alt >> nth=0")
        await save_download_selector.click()

        # Aguardar o carregamento da página depois do clique no botão para salvar o download
        await asyncio.sleep(60)

        print("Download concluído.")
        
        # Fechar o navegador
        await browser.close()

# Executa a função principal
asyncio.run(scrape_site())
