import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

CRED_JSON = "credenciais_sheets.json"
SHEET_ID = "1CFI3282Mx7MDw13RK5Vq7cA0BJxKsN3h7pwjBzFVogA"
WORKSHEET_TITLE = "Acompanhamento 2026"

COL_SEI = "SEI"
COL_STATUS = "STATUS"

REGEX_SEI = r"\d{7,}\.\d+\/\d{4}-\d+"

def normalize(s: str) -> str:
    return (s or "").strip().upper()


def extrair_ultimo_sei(texto):
    encontrados = re.findall(REGEX_SEI, texto)
    if not encontrados:
        return None
    return encontrados[-1]


def main():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        CRED_JSON, scope
    )
    client = gspread.authorize(creds)

    sh = client.open_by_key(SHEET_ID)
    ws = sh.worksheet(WORKSHEET_TITLE)

    values = ws.get_all_values()

    header = [normalize(x) for x in values[0]]

    idx_sei = header.index(COL_SEI)
    idx_status = header.index(COL_STATUS)

    seis = []

    for row in values[1:]:

        sei_raw = row[idx_sei] if idx_sei < len(row) else ""
        status = row[idx_status] if idx_status < len(row) else ""

        status = normalize(status)

        if "CONCLUID" in status:
            continue

        ultimo_sei = extrair_ultimo_sei(sei_raw)

        if ultimo_sei:
            seis.append(ultimo_sei)

    uniq = list(dict.fromkeys(seis))

    print(f"Total válidos: {len(seis)}")
    print(f"Únicos: {len(uniq)}")

    print("\n===== SEIs PROCESSADOS =====")
    for s in uniq:
        print(s)


if __name__ == "__main__":
    main()
