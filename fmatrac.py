import os
import sys
import time
import smtplib
from datetime import datetime, timedelta
from email.message import EmailMessage
from openpyxl import Workbook
from playwright.sync_api import sync_playwright

def run_fmatrac_logic():
    # --- DÁTUMOK ---
    now = datetime.now()
    start_date = datetime(now.year, 1, 1)  
    end_date = now                         
    
    start_str = start_date.strftime("%Y-%m-%d")
    end_str = end_date.strftime("%Y-%m-%d")
    print(f"Lekérdezési időszak: {start_str} - {end_str}")
    
    output_file = f"fmatrac_export_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    wb = Workbook()
    ws = wb.active
    ws.title = "adatok"
    ws.append(["Dátum", "Oszlop 1", "Oszlop 2", "Oszlop 3"])

    email_login = "buki.bertold@mavericklodges.com"
    password_login = "Abc123"

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()

            page.goto("https://maverick-athenaeum.felhomatrac.com/")

            # LOGIN
            page.fill("#user", email_login)
            page.fill("#pwd", password_login)
            page.keyboard.press("Enter")
            page.wait_for_load_state("networkidle")

            # NAVIGÁCIÓ
            page.locator("a.nav-link.dropdown-toggle", has_text="Beállítások").click()
            page.locator('a.dropdown-item', has_text="MEWS Számlázás").click()
            page.wait_for_load_state("networkidle")

            current_date = start_date
            while current_date <= end_date:
                date_str = current_date.strftime("%Y-%m-%d")
                print(f"Feldolgozás: {date_str}")

                date_input = page.locator("#billingDate")
                date_input.wait_for(state="visible")

                # Előző adat mentése ellenőrzéshez (JS trim() használatával)
                first_row_locator = page.locator('div.row.size11.list-row div.col-2').first
                old_date_text = first_row_locator.inner_text().strip() if first_row_locator.count() > 0 else ""

                date_input.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                page.keyboard.type(date_str)
                page.keyboard.press("Enter")

                # Félrekattintás az aktiváláshoz
                page.locator('div.size14.b.lh30', has_text="Számlázás Monitoring").click()

                # Várakozás a frissülésre (.trim() javítva)
                try:
                    page.wait_for_function(
                        f'() => {{ const el = document.querySelector("div.row.size11.list-row div.col-2"); return !el || el.innerText.trim() !== "{old_date_text}"; }}',
                        timeout=5000
                    )
                except: pass

                time.sleep(1)

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
            return output_file
    except Exception as e:
        print(f"Hiba a futtatás során: {e}")
        return None

def send_email(file_path):
    email_user = os.environ.get('EMAIL_USER')
    email_pass = os.environ.get('EMAIL_PASS')
    recipient = os.environ.get('EMAIL_RECIPIENT')

    msg = EmailMessage()
    msg['Subject'] = f'FMATRAC Teljes Éves Riport - {datetime.now().strftime("%Y-%m-%d")}'
    msg['From'] = email_user
    msg['To'] = recipient
    msg.set_content("Szia!\n\nMellékelten küldöm a fmatrac rendszerből kinyert adatokat év elejétől a mai napig.")

    with open(file_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=file_path)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email_user, email_pass)
        smtp.send_message(msg)
    print("Email elküldve.")

if __name__ == "__main__":
    file = run_fmatrac_logic()
    if file:
        send_email(file)
