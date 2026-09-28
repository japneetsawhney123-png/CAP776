import openpyxl
from datetime import datetime


def energy_score(x):
    return {"high": 3, "medium": 2, "low": 1}.get(str(x).lower(), 0)


def avg(values):
    return sum(values) / len(values) if values else 0


def read_excel(file):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        rows = []

        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] is not None:
                rows.append(list(r))

        wb.close()
        return rows

    except FileNotFoundError:
        print("Excel file not found.")
        return []


def valid_rows(rows):
    valid = []

    for r in rows:
        try:
            for i in range(1, 8):
                if float(r[i]) < 0:
                    raise ValueError

            if not 1 <= float(r[10]) <= 5:
                raise ValueError

            if not 1 <= float(r[11]) <= 5:
                raise ValueError

            if str(r[12]).lower() not in ["high", "medium", "low"]:
                raise ValueError

            valid.append(r)

        except (ValueError, TypeError, IndexError):
            continue

    return valid


def calculate(rows):

    tpi = avg([float(r[4]) for r in rows])
    aai = avg([float(r[3]) + float(r[5]) for r in rows])
    phai = avg([float(r[2]) for r in rows])
    sri = avg([float(r[1]) for r in rows])
    abi = avg([float(r[9]) for r in rows])
    tui = avg([float(r[8]) for r in rows])

    ei = avg([
        (float(r[10]) + float(r[11]) + energy_score(r[12])) / 3
        for r in rows
    ])

    dci = len(rows) / 40 * 100

    pai = (
        0.15 * tpi +
        0.20 * aai +
        0.15 * phai +
        0.20 * sri +
        0.15 * tui +
        0.10 * ei +
        0.05 * dci
    )

    return tpi, aai, phai, sri, abi, tui, ei, dci, pai


def relationships(rows):

    sleep = avg([float(r[1]) for r in rows])
    study = avg([float(r[3]) for r in rows])
    coding = avg([float(r[4]) for r in rows])

    sleep_high = avg([
        energy_score(r[12]) for r in rows if float(r[1]) >= sleep
    ])
    sleep_low = avg([
        energy_score(r[12]) for r in rows if float(r[1]) < sleep
    ])

    study_high = avg([
        float(r[11]) for r in rows if float(r[3]) >= study
    ])
    study_low = avg([
        float(r[11]) for r in rows if float(r[3]) < study
    ])

    coding_high = avg([
        energy_score(r[12]) for r in rows if float(r[4]) >= coding
    ])
    coding_low = avg([
        energy_score(r[12]) for r in rows if float(r[4]) < coding
    ])

    return sleep_high, sleep_low, study_high, study_low, coding_high, coding_low


def main():

    rows = read_excel("1234567.xlsx")
    rows = valid_rows(rows)

    if not rows:
        print("No valid data found.")
        return

    tpi, aai, phai, sri, abi, tui, ei, dci, pai = calculate(rows)

    print("\n========== MY DATA, MY STORY ==========")

    print("\nACTIVITY SUMMARY")
    print("Expected Days       :", 40)
    print("Valid Days          :", len(rows))
    print("Missing Days        :", 40 - len(rows))
    print("Invalid Records     :", 0)
    print("Average Sleep       :", round(avg([r[1] for r in rows]), 2))
    print("Average Fitness     :", round(avg([r[2] for r in rows]), 2))
    print("Average Study       :", round(avg([r[3] for r in rows]), 2))
    print("Average Coding      :", round(avg([r[4] for r in rows]), 2))
    print("Average Class       :", round(avg([r[5] for r in rows]), 2))
    print("Average Other       :", round(avg([r[7] for r in rows]), 2))
    print("Average Free Time   :", round(avg([r[9] for r in rows]), 2))

    print("\nACTIVITY INDICES")
    print("PAI  =", round(pai, 2))
    print("TPI  =", round(tpi, 2), "min/day")
    print("AAI  =", round(aai, 2), "min/day")
    print("PhAI =", round(phai, 2), "min/day")
    print("SRI  =", round(sri, 2), "min/day")
    print("ABI  =", round(abi, 2), "min/day")
    print("TUI  =", round(tui, 2), "min/day")
    print("EI   =", round(ei, 2), "/5")
    print("DCI  =", round(dci, 2), "%")

    sh, sl, sth, stl, ch, cl = relationships(rows)

    print("\nKEY FINDINGS")
    print("Sleep-Energy          :", round(sh, 2), "vs", round(sl, 2))
    print("Study-Satisfaction   :", round(sth, 2), "vs", round(stl, 2))
    print("Coding-Energy        :", round(ch, 2), "vs", round(cl, 2))


if __name__ == "__main__":
    main()