```python
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

            # CSAK AZ ATHENAEUMNÁL:
            # közvetlenül megnyitjuk a Számla Monitoring oldalt,
            # hogy biztosan betöltődjön a billingDate mező.
            if company_name == "Maverick Athenaeum":
                billing_url = url.rstrip("/") + "/billfuncs/billing"

                page.goto(
                    billing_url,
                    wait_until="domcontentloaded",
                    timeout=60000
                )

                page.wait_for_load_state("networkidle")

                ath_billing_date = page.locator(
                    'input#billingDate.form-control.datepicker'
                )

                # Először azt várjuk meg, hogy bekerüljön a HTML-be.
                ath_billing_date.wait_for(
                    state="attached",
                    timeout=60000
                )

                # Odaviszi a képernyőt a mezőhöz.
                ath_billing_date.scroll_into_view_if_needed()

                # Ezután megvárjuk, hogy valóban látható legyen.
                ath_billing_date.wait_for(
                    state="visible",
                    timeout=60000
                )

                print(
                    "Athenaeum: a billingDate mező sikeresen "
                    "betöltődött és látható."
                )

            else:
                # A többi cégnél marad az eredeti navigáció.
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
                        f'() => {{ const el = document.querySelector("div.row.size11.list-row div.col-2"); return !el || el.innerText.trim() !== "{old_val}"; }}',
                        timeout=4000
                    )
                except:
                    pass

                time.sleep(0.4)

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
```
