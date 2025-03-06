from playwright.sync_api import sync_playwright
import pandas as pd
import time

# Função para ler a planilha Excel
def ler_excel(nome_planilha):
    # Ler planilha excel
    df = pd.read_excel(nome_planilha, sheet_name="Sheet1", engine="openpyxl")
    return df

def preencher_formulario(page, df):
    # Mapeamento dos componentes da página
    input_address = page.locator("xpath=//label[contains(text(),'Address')]/following-sibling::input")
    input_company_name = page.locator("xpath=//label[contains(text(),'Company Name')]/following-sibling::input")
    input_first_name = page.locator("xpath=//label[contains(text(),'First Name')]/following-sibling::input")
    input_last_name = page.locator("xpath=//label[contains(text(),'Last Name')]/following-sibling::input")
    input_email = page.locator("xpath=//label[contains(text(),'Email')]/following-sibling::input")
    input_phone_number = page.locator("xpath=//label[contains(text(),'Phone Number')]/following-sibling::input")
    input_role_in_company = page.locator("xpath=//label[contains(text(),'Role in Company')]/following-sibling::input")
    click_submit = page.locator("xpath=//input[@type='submit']")

    # Preenchimento dos componentes da página
    for _, row in df.iterrows():
        time.sleep(0.5)
        input_first_name.fill(row["First Name"])
        time.sleep(0.5)
        input_last_name.fill(row["Last Name"])
        time.sleep(0.5)
        input_company_name.fill(row["Company Name"])
        time.sleep(0.5)
        input_role_in_company.fill(row["Role in Company"])
        time.sleep(0.5)
        input_address.fill(str(row["Address"]))
        time.sleep(0.5)
        input_email.fill(row["Email"])
        time.sleep(0.5)
        input_phone_number.fill(str(row["Phone Number"]))
        time.sleep(0.5)

        # Clique no botão Submit para finalizar
        click_submit.click()

def main():
    planilha = r"base/challenge.xlsx"
    df = ler_excel(planilha)

    with sync_playwright() as p:
        # Abrir o navegador
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://www.rpachallenge.com/")

        # Clique em "Começar"
        page.wait_for_selector("body > app-root > div.body.row1.scroll-y > app-rpa1 > div > div.instructions.col.s3.m3.l3.uiColorSecondary > div:nth-child(7) > button", state="visible")
        page.locator("body > app-root > div.body.row1.scroll-y > app-rpa1 > div > div.instructions.col.s3.m3.l3.uiColorSecondary > div:nth-child(7) > button").click()

        # Preencher o formulário com os dados da planilha
        preencher_formulario(page, df)

        # Fechar o navegador
        browser.close()

if __name__ == "__main__":
    main()
