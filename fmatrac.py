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

    print(f"\n--- {company_name} feldolgozása ---", flush=True)

    safe_name = (
        company_name
        .replace(" ", "_")
        .replace(".", "")
        .lower()
    )

    output_file = (
        f"fmatrac_{safe_name}_"
        f"{now.strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "adatok"
    ws.append([
        "Dátum",
        "Oszlop 1",
        "Oszlop 2",
        "Oszlop 3"
    ])

    browser = None

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True
            )

            context = browser.new_context(
                viewport={
                    "width": 1280,
                    "height": 800
                }
            )

            page = context.new_page()
            page.set_default_timeout(60000)

            print(
                f"{company_name}: oldal megnyitása: {url}",
                flush=True
            )

            page.goto(
                url,
                wait_until="networkidle",
                timeout=60000
            )

            # Bejelentkezés
            print(
                f"{company_name}: bejelentkezés...",
                flush=True
            )

            page.locator("#user").wait_for(
                state="visible"
            )

            page.fill(
                "#user",
                login_user
            )

            page.fill(
                "#pwd",
                login_pass
            )

            page.keyboard.press("Enter")

            page.wait_for_load_state(
                "networkidle"
            )

            print(
                f"{company_name}: bejelentkezés sikeres.",
                flush=True
            )

            # Navigáció a Beállítások menübe
            beallitasok = page.locator(
                "a.nav-link.dropdown-toggle",
                has_text="Beállítások"
            )

            beallitasok.wait_for(
                state="visible"
            )

            beallitasok.click()

            # Számlamonitoring / billing menüpont
            mews_gomb = page.locator(
                'a[href*="billfuncs/billing"]'
            )

            try:
                mews_gomb.first.wait_for(
                    state="visible",
                    timeout=15000
                )

                mews_gomb.first.click()

                page.wait_for_load_state(
                    "networkidle"
                )

            except Exception:
                print(
                    f"{company_name}: a menüpont nem volt "
                    "kattintható, közvetlen oldalmegnyitás.",
                    flush=True
                )

                billing_url = (
                    url.rstrip("/")
                    + "/billfuncs/billing"
                )

                page.goto(
                    billing_url,
                    wait_until="domcontentloaded",
                    timeout=60000
                )

            # Minden cégnél ellenőrizzük a dátummezőt.
            date_input = page.locator(
                "input#billingDate"
            )

            try:
                date_input.wait_for(
                    state="visible",
                    timeout=15000
                )

            except Exception:
                print(
                    f"{company_name}: a billingDate mező "
                    "nem látható, közvetlen oldalmegnyitás.",
                    flush=True
                )

                billing_url = (
                    url.rstrip("/")
                    + "/billfuncs/billing"
                )

                page.goto(
                    billing_url,
                    wait_until="domcontentloaded",
                    timeout=60000
                )

                date_input = page.locator(
                    "input#billingDate"
                )

                date_input.wait_for(
                    state="attached",
                    timeout=60000
                )

                date_input.scroll_into_view_if_needed()

                date_input.wait_for(
                    state="visible",
                    timeout=60000
                )

            print(
                f"{company_name}: a billingDate mező "
                "látható, feldolgozás indul.",
                flush=True
            )

            current_date = start_date

            while current_date <= end_date:
                date_str = current_date.strftime(
                    "%Y-%m-%d"
                )

                date_input = page.locator(
                    "#billingDate"
                )

                date_input.wait_for(
                    state="visible"
                )

                first_row_loc = page.locator(
                    "div.row.size11.list-row "
                    "div.col-2"
                ).first

                row_exists = (
                    first_row_loc.count()
                )

                old_val = (
                    first_row_loc
                    .inner_text()
                    .strip()
                    if row_exists > 0
                    else "empty"
                )

                date_input.click()

                page.keyboard.press(
                    "Control+A"
                )

                page.keyboard.press(
                    "Backspace"
                )

                page.keyboard.type(
                    date_str
                )

                page.keyboard.press(
                    "Enter"
                )

                page.mouse.click(
                    0,
                    0
                )

                try:
                    page.wait_for_function(
                        """
                        ([selector, oldValue]) => {
                            const el =
                                document.querySelector(
                                    selector
                                );

                            return (
                                !el
                                || el.innerText.trim()
                                !== oldValue
                            );
                        }
                        """,
                        arg=[
                            (
                                "div.row.size11.list-row "
                                "div.col-2"
                            ),
                            old_val
                        ],
                        timeout=4000
                    )

                except Exception:
                    pass

                time.sleep(0.4)

                rows = page.locator(
                    "div.row.size11.list-row"
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
                        cols = (
                            rows
                            .nth(i)
                            .locator("div.col-2")
                        )

                        row_data = [
                            date_str
                        ]

                        col_count = cols.count()

                        for j in range(3):
                            if j < col_count:
                                val = (
                                    cols
                                    .nth(j)
                                    .inner_text()
                                    .strip()
                                )
                            else:
                                val = ""

                            row_data.append(
                                val
                            )

                        ws.append(
                            row_data
                        )

                print(
                    f"{company_name}: {date_str} kész.",
                    flush=True
                )

                current_date += timedelta(
                    days=1
                )

            wb.save(
                output_file
            )

            print(
                f"{company_name}: Excel elkészült: "
                f"{output_file}",
                flush=True
            )

            browser.close()
            browser = None

            return output_file, company_name

    except Exception as error:
        print(
            f"Hiba történt ({company_name}): "
            f"{error}",
            flush=True
        )

        try:
            wb.close()
        except Exception:
            pass

        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass

        if os.path.exists(output_file):
            try:
                os.remove(output_file)
            except Exception:
                pass

        return None, company_name


def send_email(
    file_path,
    company_name,
    recipient_email
):
    email_user = os.environ.get(
        "EMAIL_USER"
    )

    email_app_password = os.environ.get(
        "EMAIL_APP_PASSWORD"
    )

    if not email_user:
        raise RuntimeError(
            "Hiányzik az EMAIL_USER környezeti változó."
        )

    if not email_app_password:
        raise RuntimeError(
            "Hiányzik az EMAIL_APP_PASSWORD "
            "környezeti változó."
        )

    if not recipient_email:
        raise RuntimeError(
            "Hiányzik az EMAIL_RECIPIENT "
            "környezeti változó."
        )

    msg = EmailMessage()

    msg["Subject"] = (
        f"FMATRAC Riport - {company_name} - "
        f"{datetime.now().strftime('%Y-%m-%d')}"
    )

    msg["From"] = email_user
    msg["To"] = recipient_email

    msg.set_content(
        "Szia!\n\n"
        f"Mellékelten küldöm a(z) "
        f"{company_name} kinyert adatait."
    )

    attachment_name = os.path.basename(
        file_path
    )

    with open(
        file_path,
        "rb"
    ) as attachment:
        msg.add_attachment(
            attachment.read(),
            maintype="application",
            subtype=(
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            filename=attachment_name
        )

    with smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465,
        timeout=120
    ) as smtp:
        smtp.login(
            email_user,
            email_app_password
        )

        smtp.send_message(
            msg
        )

    print(
        f"E-mail elküldve: {company_name}",
        flush=True
    )


def validate_environment(
    companies,
    target_email
):
    missing = []

    if not os.environ.get("EMAIL_USER"):
        missing.append("EMAIL_USER")

    if not os.environ.get(
        "EMAIL_APP_PASSWORD"
    ):
        missing.append(
            "EMAIL_APP_PASSWORD"
        )

    if not target_email:
        missing.append(
            "EMAIL_RECIPIENT"
        )

    for company in companies:
        if not os.environ.get(
            company["user_env"]
        ):
            missing.append(
                company["user_env"]
            )

        if not os.environ.get(
            company["pass_env"]
        ):
            missing.append(
                company["pass_env"]
            )

    if missing:
        unique_missing = list(
            dict.fromkeys(missing)
        )

        raise RuntimeError(
            "Hiányzó környezeti változók: "
            + ", ".join(unique_missing)
        )


if __name__ == "__main__":
    target_email = os.environ.get(
        "EMAIL_RECIPIENT"
    )

    companies = [
        {
            "name": "Maverick Athenaeum",
            "url": (
                "https://maverick-athenaeum."
                "felhomatrac.com/"
            ),
            "user_env": "LOGIN_0",
            "pass_env": "PASS_0"
        },
        {
            "name": "Maverick Downtown Apartment",
            "url": (
                "https://maverick-apartments."
                "felhomatrac.org/"
            ),
            "user_env": "LOGIN_1",
            "pass_env": "PASS_1"
        },
        {
            "name": "Maverick Budapest Soho",
            "url": (
                "https://maverick-lodges."
                "felhomatrac.org/"
            ),
            "user_env": "LOGIN_2",
            "pass_env": "PASS_2"
        },
        {
            "name": "Giselle Vintage Doubles",
            "url": (
                "https://maverick."
                "felhomatrac.org/"
            ),
            "user_env": "LOGIN_3",
            "pass_env": "PASS_3"
        },
        {
            "name": "Maverick Central Market",
            "url": (
                "https://maverick-urban-lodge."
                "felhomatrac.com/"
            ),
            "user_env": "LOGIN_4",
            "pass_env": "PASS_4"
        },
        {
            "name": "Giselle Buda Castle",
            "url": (
                "https://maverick-buda-castle."
                "felhomatrac.com/"
            ),
            "user_env": "LOGIN_5",
            "pass_env": "PASS_5"
        },
        {
            "name": "Amberlyn Management Kft.",
            "url": (
                "https://amberlyn."
                "felhomatrac.com/"
            ),
            "user_env": "LOGIN_6",
            "pass_env": "PASS_6"
        }
    ]

    validate_environment(
        companies,
        target_email
    )

    successful_companies = []
    failed_companies = []

    for comp in companies:
        username = os.environ.get(
            comp["user_env"]
        )

        password = os.environ.get(
            comp["pass_env"]
        )

        file_path, company_name = (
            run_fmatrac_logic(
                comp["name"],
                comp["url"],
                username,
                password
            )
        )

        if file_path:
            try:
                send_email(
                    file_path,
                    company_name,
                    target_email
                )

                successful_companies.append(
                    company_name
                )

            except Exception as error:
                print(
                    f"E-mail küldési hiba "
                    f"({company_name}): {error}",
                    flush=True
                )

                failed_companies.append(
                    company_name
                )

            finally:
                if os.path.exists(
                    file_path
                ):
                    os.remove(
                        file_path
                    )

                    print(
                        f"Ideiglenes fájl törölve: "
                        f"{file_path}",
                        flush=True
                    )

        else:
            failed_companies.append(
                company_name
            )

    print(
        "\n==============================",
        flush=True
    )

    print(
        "FMATRAC FELDOLGOZÁS ÖSSZESÍTÉSE",
        flush=True
    )

    for company_name in successful_companies:
        print(
            f"{company_name}: SIKER",
            flush=True
        )

    for company_name in failed_companies:
        print(
            f"{company_name}: HIBA",
            flush=True
        )

    print(
        f"Sikeres cégek száma: "
        f"{len(successful_companies)}",
        flush=True
    )

    print(
        f"Hibás cégek száma: "
        f"{len(failed_companies)}",
        flush=True
    )

    if failed_companies:
        raise RuntimeError(
            "Nem sikerült minden céget feldolgozni: "
            + ", ".join(failed_companies)
        )

    print(
        "MIND A 7 CÉG FELDOLGOZÁSA "
        "ÉS E-MAIL-KÜLDÉSE SIKERES.",
        flush=True
    )
