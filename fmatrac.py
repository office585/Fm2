import os
import smtplib
import time
from datetime import datetime, timedelta
from email.message import EmailMessage
from openpyxl import Workbook
from playwright.sync_api import sync_playwright

def run_fmatrac_logic(company_name, url, login_user, login_pass):
    now = datetime.now()
    start_date = datetime(now.year, 1, 1)
    end_date = now
    
    print(f"\n--- {company_name} feldolgozása ---")
    print(f"URL: {url}")
    
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
            page.set_default_timeout(60000) # 1 perces türelmi idő

            # Oldal megnyitása
            page.goto(url, wait_until="networkidle")

            # Bejelentkezés
            page.fill("#user", login_user)
            page.fill("#pwd", login_pass)
            page.keyboard.press("Enter")
            page.wait_for_load_state("networkidle")

            # Navigáció a pontos menüpontra (link alapján, hogy ne tévedjen el)
            page.locator("a.nav-link.dropdown-toggle", has_text="Beállítások").click()
            page.locator('a[href*="billing"]').wait_for(state="visible")
            page.locator('a[href*="billing"]').click()
            page.wait_for_load_state("networkidle")

            current_date = start_date
            while current_date <= end_date:
                date_str = current_date.strftime("%Y-%m-%d")
                
                # Dátum mező megvárása
                date_input = page.locator("#billingDate")
                date_input.wait_for(state="visible")

                # Aktuális első sor elmentése ellenőrzéshez
                first_row_loc = page.locator('div.row.size11.list-row div.col-2').first
                old_val = first_row_loc.inner_text().strip() if first_row_loc.count() > 0 else "empty"

                # Dátum beírása
                date_input.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                page.keyboard.type(date_str)
                page.keyboard.press("Enter")

                # Kattintás a sarokba a fókusz elvételéhez (lekérdezés indítása)
                page.mouse.click(0, 0) 
                
                # Intelligens várakozás: Várjuk, amíg az adat frissül (vagy max 5 mp)
                try:
                    page.wait_for_function(
                        f'() => {{ const el = document.querySelector("div.row.size11.list-row div.col-2"); return !el || el.innerText.trim() !== "{old_val}"; }}',
                        timeout=5000
                    )
                except:
                    pass # Ha nincs adat vagy lassú, megyünk tovább

                # Rövid biztonsági pihenő (nem bugol be a naptár)
                time.sleep(0.5)

                # Adatok kimentése
                rows = page.locator('div.row.size11.list-row')
                count = rows.count()
                if count == 0:
                    ws.append([date_str, "NINCS ADAT", "", ""])
                else:
                    for i in range(count):
                        cols = rows.nth(i).locator("div.col-2")
                        row_data = [date_str]
                        col_count = cols.count()
                        for j in range(3):
                            val = cols.nth(j).inner_text().strip() if j < col_count else ""
                            row_data.append(val)
                        ws.append(row_data)

                current_date += timedelta(days=1)

            browser.close()
            wb.save(output_file)
            return output_file, company_name

    except Exception as e:
        print(f"Hiba történt ({company_name}): {e}")
        return None, company_name

def send_email(file_path, company_name, recipient_email):
    email_user = os.environ.get('EMAIL_USER')
    email_pass = os.environ.get('EMAIL_PASS')

    msg = EmailMessage()
    msg['Subject'] = f'FMATRAC Riport - {company_name} - {datetime.now().strftime("%Y-%m-%d")}'
    msg['From'] = email_user
    msg['To'] = recipient_email
    msg.set_content(f"Szia!\n\nMellékelten küldöm a(z) {company_name} kinyert adatait év elejétől a mai napig.")

    with open(file_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=file_path)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email_user, email_pass)
        smtp.send_message(msg)
    print(f"E-mail sikeresen elküldve: {company_name}")

if __name__ == "__main__":
    target_email = os.environ.get('EMAIL_RECIPIENT')

    companies = [
        {"name": "Maverick Athenaeum", "url": "https://maverick-athenaeum.felhomatrac.com/", "user_env": "LOGIN_0", "pass_env": "PASS_0"},
        {"name": "Maverick Downtown Apartment", "url": "https://maverick-apartments.felhomatrac.org/", "user_env": "LOGIN_1", "pass_env": "PASS_1"},
        {"name": "Maverick Budapest Soho", "url": "https://maverick-lodges.felhomatrac.org/", "user_env": "LOGIN_2", "pass_env": "PASS_2"},
        {"name": "Giselle Vintage Doubles", "url": "https://maverick.felhomatrac.org/", "user_env": "LOGIN_3", "pass_env": "PASS_3"},
        {"name": "Maverick Central Market", "url": "https://maverick-urban-lodge.felhomatrac.com/", "user_env": "LOGIN_4", "pass_env": "PASS_4"},
        {"name": "Giselle Buda Castle", "url": "https://maverick-buda-castle.felhomatrac.com/", "user_env": "LOGIN_5", "pass_env": "PASS_5"},
    ]

    for comp in companies:
        u = os.environ.get(comp["user_env"])
        p = os.environ.get(comp["pass_env"])
        
        if u and p:
            file, c_name = run_fmatrac_logic(comp["name"], comp["url"], u, p)
            if file:
                send_email(file, c_name, target_email)
                if os.path.exists(file):
                    os.remove(file)
        else:
            print(f"Kihagyva: {comp['name']} (Hiányzó környezeti változók)")
