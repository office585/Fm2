import os
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
from openpyxl import Workbook
from playwright.sync_api import sync_playwright
import time

def run_fmatrac_logic(company_name, login_user, login_pass):
    now = datetime.now()
    start_date = datetime(now.year, 1, 1)
    end_date = now
    
    print(f"\n--- {company_name} feldolgozása kezdődik ---")
    
    safe_name = company_name.replace(" ", "_").lower()
    output_file = f"fmatrac_{safe_name}_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    wb = Workbook()
    ws = wb.active
    ws.title = "adatok"
    ws.append(["Dátum", "Oszlop 1", "Oszlop 2", "Oszlop 3"])

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={'width': 1280, 'height': 800})
            page = context.new_page()

            page.goto("https://maverick-athenaeum.felhomatrac.com/")

            # LOGIN
            page.fill("#user", login_user)
            page.fill("#pwd", login_pass)
            page.keyboard.press("Enter")
            page.wait_for_load_state("networkidle")

            # NAVIGÁCIÓ
            page.locator("a.nav-link.dropdown-toggle", has_text="Beállítások").click()
            page.locator('a.dropdown-item', has_text="MEWS Számlázás").click()
            page.wait_for_load_state("networkidle")

            current_date = start_date
            while current_date <= end_date:
                date_str = current_date.strftime("%Y-%m-%d")
                
                date_input = page.locator("#billingDate")
                date_input.wait_for(state="visible")

                date_input.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                page.keyboard.type(date_str)
                page.keyboard.press("Enter")

                page.locator('div.size14.b.lh30', has_text="Számlázás Monitoring").click()
                
                time.sleep(1.5) # Biztonsági várakozás a frissülésre

                # Adatmentés
                rows = page.locator('div.row.size11.list-row')
                if rows.count() == 0:
                    ws.append([date_str, "NINCS ADAT", "", ""])
                else:
                    for i in range(rows.count()):
                        cols = rows.nth(i).locator("div.col-2")
                        row_data = [date_str]
                        for j in range(3):
                            row_data.append(cols.nth(j).inner_text().strip() if j < cols.count() else "")
                        ws.append(row_data)

                current_date += timedelta(days=1)

            browser.close()
            wb.save(output_file)
            return output_file, company_name
    except Exception as e:
        print(f"Hiba a(z) {company_name} futtatása során: {e}")
        return None, company_name

def send_email(file_path, company_name):
    email_user = os.environ.get('EMAIL_USER')
    email_pass = os.environ.get('EMAIL_PASS')
    recipient = os.environ.get('EMAIL_RECIPIENT')

    msg = EmailMessage()
    msg['Subject'] = f'FMATRAC Riport - {company_name} - {datetime.now().strftime("%Y-%m-%d")}'
    msg['From'] = email_user
    msg['To'] = recipient
    msg.set_content(f"Szia!\n\nMellékelten küldöm a(z) {company_name} adatait év elejétől a mai napig.")

    with open(file_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=file_path)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email_user, email_pass)
        smtp.send_message(msg)
    print(f"Email elküldve: {company_name}")

if __name__ == "__main__":
    # Az ÖSSZES cég listája (eredeti + 5 új)
    companies = [
        {"name": "Maverick Athenaeum",           "user_env": "LOGIN_0", "pass_env": "PASS_0"}, # Ez volt az eddigi
        {"name": "Maverick Downtown Apartment", "user_env": "LOGIN_1", "pass_env": "PASS_1"},
        {"name": "Maverick Budapest Soho",       "user_env": "LOGIN_2", "pass_env": "PASS_2"},
        {"name": "Giselle Vintage Doubles",      "user_env": "LOGIN_3", "pass_env": "PASS_3"},
        {"name": "Maverick Central Market",      "user_env": "LOGIN_4", "pass_env": "PASS_4"},
        {"name": "Giselle Buda Castle",          "user_env": "LOGIN_5", "pass_env": "PASS_5"},
    ]

    for comp in companies:
        u = os.environ.get(comp["user_env"])
        p = os.environ.get(comp["pass_env"])
        
        if u and p:
            file, c_name = run_fmatrac_logic(comp["name"], u, p)
            if file:
                send_email(file, c_name)
                os.remove(file)
        else:
            print(f"Hiba: Hiányzó login adatok ({comp['user_env']}) - Kihagyva.")
