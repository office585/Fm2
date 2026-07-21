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

    safe_name = company_name.replace(" ", "_").lower()
    output_file = f"fmatrac_{safe_name}_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "adatok"
    ws.append(["Dátum", "Oszlop 1", "Oszlop 2", "Oszlop 3"])

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)

            context = browser.new_context(
                viewport={'width': 1280, 'height': 800}
            )

            page = context.new_page()
            page.set_default_timeout(60000)

            page.goto(url, wait_until="networkidle")

            # Bejelentkezés
            page.fill("#user", login_user)
            page.fill("#pwd", login_pass)
            page.keyboard.press("Enter")
            page.wait_for_load_state("networkidle")

            # Navigáció
            beallitasok = page.locator(
                "a.nav-link.dropdown-toggle",
                has_text="Beállítások"
            )

            beallitasok.wait_for(state="visible")
            beallitasok.click()

            # Pontosított link keresés
            mews_gomb = page.locator(
                'a[href*="billfuncs/billing"]'
            )

            mews_gomb.first.wait_for(state="visible")
            mews_gomb.first.click()

            page.wait_for_load_state("networkidle")

            # Csak a Maverick Athenaeumnál:
            # ellenőrizzük, hogy a billingDate mező valóban látható-e.
            # Ha a normál navigáció után nem látható, közvetlenül
            # megnyitjuk a számlamonitoring oldalt.
            if company_name == "Maverick Athenaeum":
                athenaeum_date_input = page.locator(
                    'input#billingDate'
                )

                try:
                    athenaeum_date_input.wait_for(
                        state="visible",
                        timeout=10000
                    )

                except Exception:
                    print(
                        "Athenaeum: a billingDate mező még nem látható, "
                        "közvetlen oldalmegnyitás következik."
                    )

                    billing_url = (
                        url.rstrip("/") + "/billfuncs/billing"
                    )

                    page.goto(
                        billing_url,
                        wait_until="domcontentloaded",
                        timeout=60000
                    )

                    athenaeum_date_input = page.locator(
                        'input#billingDate'
                    )

                    # Megvárjuk, hogy a mező bekerüljön az oldalba.
                    athenaeum_date_input.wait_for(
                        state="attached",
                        timeout=60000
                    )

                    # Szükség esetén odagörgetünk.
                    athenaeum_date_input.scroll_into_view_if_needed()

                    # Végül megvárjuk, hogy ténylegesen látható legyen.
                    athenaeum_date_input.wait_for(
                        state="visible",
                        timeout=60000
                    )

                print(
                    "Athenaeum: a billingDate mező látható, "
                    "a feldolgozás indul."
                )

            current_date = start_date

            while current_date <= end_date:
                date_str = current_date.strftime("%Y-%m-%d")

                date_input = page.locator("#billingDate")
                date_input.wait_for(state="visible")

                # Itt volt a hiba, most már stabil:
                first_row_loc = page.locator(
                    'div.row.size11.list-row div.col-2'
                ).first

                row_exists = first_row_loc.count()

                old_val = (
                    first_row_loc.inner_text().strip()
                    if row_exists > 0
                    else "empty"
                )

                date_input.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                page.keyboard.type(date_str)
                page.keyboard.press("Enter")

                page.mouse.click(0, 0)

                try:
                    page.wait_for_function(
                        f'''() => {{
                            const el = document.querySelector(
                                "div.row.size11.list-row div.col-2"
                            );
                            return !el || el.innerText.trim() !== "{old_val}";
                        }}''',
                        timeout=4000
                    )
                except Exception:
                    pass

                time.sleep(0.4)

                rows = page.locator(
                    'div.row.size11.list-row'
                )

                count = rows.count()

                if count == 0:
                    ws.append([
                        date_str,
                        "NINCS ADAT",
                        "",
                        ""
                    ])

                else:
                    for i in range(count):
                        cols = rows.nth(i).locator(
                            "div.col-2"
                        )

                        row_data = [date_str]
                        col_count = cols.count()

                        for j in range(3):
                            val = (
                                cols.nth(j).inner_text().strip()
                                if j < col_count
                                else ""
                            )

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
    email_app_password = os.environ.get(
        'EMAIL_APP_PASSWORD'
    )

    msg = EmailMessage()

    msg['Subject'] = (
        f'FMATRAC Riport - {company_name} - '
        f'{datetime.now().strftime("%Y-%m-%d")}'
    )

    msg['From'] = email_user
    msg['To'] = recipient_email

    msg.set_content(
        f"Szia!\n\n"
        f"Mellékelten küldöm a(z) "
        f"{company_name} kinyert adatait."
    )

    with open(file_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='octet-stream',
            filename=file_path
        )

    with smtplib.SMTP_SSL(
        'smtp.gmail.com',
        465
    ) as smtp:
        smtp.login(
            email_user,
            email_app_password
        )

        smtp.send_message(msg)

    print(f"E-mail elküldve: {company_name}")


if __name__ == "__main__":
    target_email = os.environ.get(
        'EMAIL_RECIPIENT'
    )

    companies = [
        {
            "name": "Maverick Athenaeum",
            "url": "https://maverick-athenaeum.felhomatrac.com/",
            "user_env": "LOGIN_0",
            "pass_env": "PASS_0"
        },
        {
            "name": "Maverick Downtown Apartment",
            "url": "https://maverick-apartments.felhomatrac.org/",
            "user_env": "LOGIN_1",
            "pass_env": "PASS_1"
        },
        {
            "name": "Maverick Budapest Soho",
            "url": "https://maverick-lodges.felhomatrac.org/",
            "user_env": "LOGIN_2",
            "pass_env": "PASS_2"
        },
        {
            "name": "Giselle Vintage Doubles",
            "url": "https://maverick.felhomatrac.org/",
            "user_env": "LOGIN_3",
            "pass_env": "PASS_3"
        },
        {
            "name": "Maverick Central Market",
            "url": "https://maverick-urban-lodge.felhomatrac.com/",
            "user_env": "LOGIN_4",
            "pass_env": "PASS_4"
        },
        {
            "name": "Giselle Buda Castle",
            "url": "https://maverick-buda-castle.felhomatrac.com/",
            "user_env": "LOGIN_5",
            "pass_env": "PASS_5"
        },
    ]

    for comp in companies:
        u = os.environ.get(
            comp["user_env"]
        )

        p = os.environ.get(
            comp["pass_env"]
        )

        if u and p:
            file, c_name = run_fmatrac_logic(
                comp["name"],
                comp["url"],
                u,
                p
            )

            if file:
                send_email(
                    file,
                    c_name,
                    target_email
                )

                if os.path.exists(file):
                    os.remove(file)
