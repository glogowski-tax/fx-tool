import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import requests
from datetime import date, timedelta
from io import BytesIO


def fetch_nbp_rates(date_from: date, date_to: date, currency: str = "eur") -> dict[date, float]:
    """Pobiera wszystkie kursy waluty/PLN z NBP w podanym zakresie dat (jedno zapytanie)."""
    url = f"https://api.nbp.pl/api/exchangerates/rates/a/{currency.lower()}/{date_from}/{date_to}/?format=json"
    try:
        resp = requests.get(url, timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            return {
                date.fromisoformat(r["effectiveDate"]): r["mid"]
                for r in data["rates"]
            }
    except requests.RequestException:
        pass
    return {}


def fetch_ecb_rates(date_from: date, date_to: date) -> dict[date, float]:
    """Pobiera wszystkie kursy EUR/PLN z ECB w podanym zakresie dat (jedno zapytanie)."""
    url = (
        f"https://data-api.ecb.europa.eu/service/data/EXR/"
        f"D.PLN.EUR.SP00.A?startPeriod={date_from}&endPeriod={date_to}"
        f"&format=csvdata"
    )
    rates = {}
    try:
        resp = requests.get(url, timeout=30)
        if resp.status_code == 200 and "OBS_VALUE" in resp.text:
            lines = resp.text.strip().split("\n")
            if len(lines) >= 2:
                header = lines[0].split(",")
                obs_idx = header.index("OBS_VALUE")
                date_idx = header.index("TIME_PERIOD")
                for line in lines[1:]:
                    values = line.split(",")
                    try:
                        rates[date.fromisoformat(values[date_idx])] = float(values[obs_idx])
                    except (ValueError, IndexError):
                        continue
    except requests.RequestException:
        pass
    return rates


def find_previous_rate(target_date: date, all_rates: dict[date, float]) -> tuple[float | None, date | None]:
    """Znajduje kurs z ostatniego dnia roboczego PRZED podaną datą."""
    check_date = target_date - timedelta(days=1)
    for _ in range(10):
        if check_date in all_rates:
            return all_rates[check_date], check_date
        check_date -= timedelta(days=1)
    return None, None


def parse_date_value(val) -> date | None:
    """Próbuje sparsować wartość komórki jako datę."""
    if hasattr(val, "date"):  # datetime
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                from datetime import datetime
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def process_workbook(wb, sheet_name: str, col_idx: int, source: str, currency: str, amount_col_idx: int | None, progress_bar):
    """Przetwarza arkusz — wstawia kolumnę z kursami i opcjonalnie kolumnę PLN przeliczone."""
    ws = wb[sheet_name]
    insert_col = col_idx + 1

    # Wstaw kolumnę z kursem
    ws.insert_cols(insert_col)
    source_label = "kurs NBP" if source == "NBP" else "kurs EBC"
    ws.cell(row=1, column=insert_col, value=f"{source_label} {currency}/PLN")

    # Wstaw kolumnę PLN przeliczone (jeśli wybrano kolumnę z kwotami)
    pln_col = None
    if amount_col_idx is not None:
        pln_col = insert_col + 1
        ws.insert_cols(pln_col)
        ws.cell(row=1, column=pln_col, value="PLN przeliczone")
        # Skoryguj indeks kolumny kwot jeśli jest za wstawionymi kolumnami
        if amount_col_idx >= insert_col:
            amount_col_idx += 2  # przesunięcie o 2 nowe kolumny

    total_rows = ws.max_row - 1
    if total_rows <= 0:
        return wb

    # Zbierz wszystkie daty z arkusza
    all_dates = []
    for row_num in range(2, ws.max_row + 1):
        parsed = parse_date_value(ws.cell(row=row_num, column=col_idx).value)
        if parsed:
            all_dates.append(parsed)

    progress_bar.progress(0.05, "Pobieram kursy z API...")

    # Pobierz wszystkie kursy hurtowo (jeden zakres dat, 1 zapytanie)
    if all_dates:
        date_from = min(all_dates) - timedelta(days=15)
        date_to = max(all_dates)
        if source == "NBP":
            all_rates = fetch_nbp_rates(date_from, date_to, currency)
        else:
            all_rates = fetch_ecb_rates(date_from, date_to)
    else:
        all_rates = {}

    progress_bar.progress(0.3, "Wstawiam kursy do arkusza...")

    # Wstaw kursy do arkusza
    for i, row_num in enumerate(range(2, ws.max_row + 1)):
        cell_value = ws.cell(row=row_num, column=col_idx).value
        parsed_date = parse_date_value(cell_value)

        if parsed_date:
            rate, rate_date = find_previous_rate(parsed_date, all_rates)
            if rate is not None:
                ws.cell(row=row_num, column=insert_col, value=rate)
                # Przelicz kwotę na PLN
                if pln_col and amount_col_idx:
                    amount = ws.cell(row=row_num, column=amount_col_idx).value
                    if isinstance(amount, (int, float)):
                        pln_value = round(amount * rate)
                        ws.cell(row=row_num, column=pln_col, value=pln_value)
                    else:
                        ws.cell(row=row_num, column=pln_col, value="Brak kwoty")
            else:
                ws.cell(row=row_num, column=insert_col, value="Brak kursu")
                if pln_col:
                    ws.cell(row=row_num, column=pln_col, value="Brak kursu")
        else:
            ws.cell(row=row_num, column=insert_col, value="Błędna data")
            if pln_col:
                ws.cell(row=row_num, column=pln_col, value="Błędna data")

        progress_bar.progress(0.3 + 0.7 * (i + 1) / total_rows)

    return wb


# ---- STREAMLIT UI ----

st.set_page_config(page_title="FX Tool — Kursy walut", page_icon="💱", layout="wide")

CURRENCIES = {
    "EUR": "euro", "USD": "dolar amerykański", "GBP": "funt szterling",
    "CHF": "frank szwajcarski", "CZK": "korona czeska", "DKK": "korona duńska",
    "NOK": "korona norweska", "SEK": "korona szwedzka", "HUF": "forint (Węgry)",
    "RON": "lej rumuński", "UAH": "hrywna (Ukraina)", "TRY": "lira turecka",
    "CAD": "dolar kanadyjski", "AUD": "dolar australijski", "JPY": "jen (Japonia)",
    "CNY": "yuan renminbi (Chiny)", "HKD": "dolar Hongkongu", "SGD": "dolar singapurski",
    "NZD": "dolar nowozelandzki", "MXN": "peso meksykańskie", "BRL": "real (Brazylia)",
    "ZAR": "rand (RPA)", "KRW": "won (Korea Płd.)", "INR": "rupia indyjska",
    "ILS": "nowy izraelski szekel", "THB": "bat (Tajlandia)", "PHP": "peso filipińskie",
    "IDR": "rupia indonezyjska", "MYR": "ringgit (Malezja)", "CLP": "peso chilijskie",
    "ISK": "korona islandzka", "XDR": "SDR (MFW)",
}

st.title("FX Tool — Kursy walut do Excel")
st.markdown("Załaduj plik Excel, wskaż kolumnę z datami, a aplikacja wstawi kurs wybranej waluty/PLN z dnia poprzedzającego.")

# Upload pliku
uploaded_file = st.file_uploader("Wybierz plik Excel", type=["xlsx"])

if uploaded_file is not None:
    wb = openpyxl.load_workbook(BytesIO(uploaded_file.read()), data_only=False, rich_text=False)

    # Wybór arkusza
    sheet_names = wb.sheetnames
    if len(sheet_names) == 1:
        sheet_name = sheet_names[0]
    else:
        sheet_name = st.selectbox("Wybierz arkusz", sheet_names)

    ws = wb[sheet_name]

    # Podgląd danych
    st.subheader("Podgląd danych")
    preview_data = []
    headers = []
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        headers.append(str(val) if val else f"Kolumna {get_column_letter(col)}")

    for row in range(2, min(ws.max_row + 1, 12)):  # max 10 wierszy podglądu
        row_data = {}
        for col in range(1, ws.max_column + 1):
            row_data[headers[col - 1]] = ws.cell(row=row, column=col).value
        preview_data.append(row_data)

    st.dataframe(preview_data, use_container_width=True)

    # Wybór parametrów
    col1, col2 = st.columns(2)

    with col1:
        date_column = st.selectbox("Wskaż kolumnę z datami", headers)
        col_idx = headers.index(date_column) + 1

    with col2:
        currency_options = [f"{code} — {name}" for code, name in CURRENCIES.items()]
        currency_choice = st.selectbox("Waluta", currency_options)
        currency = currency_choice.split(" — ")[0]

    col3, col4 = st.columns(2)

    with col3:
        if currency == "EUR":
            source = st.radio("Źródło kursu", ["NBP", "ECB"], horizontal=True)
        else:
            source = "NBP"
            st.info("Źródło: **NBP** (EBC dostępne tylko dla EUR)")

    with col4:
        amount_options = ["— nie przeliczaj —"] + headers
        amount_column = st.selectbox(f"Kolumna z kwotami ({currency})", amount_options)
        amount_col_idx = headers.index(amount_column) + 1 if amount_column != "— nie przeliczaj —" else None

    # Info
    source_label = "NBP" if source == "NBP" else "EBC"
    info_text = (
        f"Kurs **{currency}/PLN** z **ostatniego dnia roboczego przed datą** w kolumnie **\"{date_column}\"** "
        f"zostanie pobrany z **{source_label}** i wstawiony w nowej kolumnie obok."
    )
    if amount_col_idx:
        info_text += f"\n\nKwoty z kolumny **\"{amount_column}\"** zostaną przeliczone na PLN (zaokrąglone do pełnych złotych)."
    st.info(info_text)

    # Przycisk generowania
    if st.button("Pobierz kursy i generuj plik", type="primary"):
        with st.spinner("Pobieram kursy walut..."):
            progress = st.progress(0)
            wb = process_workbook(wb, sheet_name, col_idx, source, currency, amount_col_idx, progress)

        st.success("Gotowe! Kursy zostały dodane.")

        # Zapis do bufora i przycisk pobierania
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        original_name = uploaded_file.name.rsplit(".", 1)[0]
        st.download_button(
            label="Pobierz plik Excel z kursami",
            data=output,
            file_name=f"{original_name}_z_kursami.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
