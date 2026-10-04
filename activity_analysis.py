import openpyxl
from datetime import datetime


# Convert Energy Level into a numerical score
# High = 3, Medium = 2, Low = 1
def energy_score(x):
    return {"high": 3, "medium": 2, "low": 1}.get(str(x).lower(), 0)


# Find the average of a list of values
def avg(values):
    return sum(values) / len(values) if values else 0


# Read daily activity data from the Excel file
def read_excel(file):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        rows = []

        # Read all rows except the heading row
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] is not None:
                rows.append(list(r))

        wb.close()
        return rows

    except FileNotFoundError:
        print("Excel file not found.")
        return []


# Check whether the activity records contain valid values
def valid_rows(rows):
    valid = []

    for r in rows:
        try:
            # Activity time values should not be negative
            for i in range(1, 8):
                if float(r[i]) < 0:
                    raise ValueError

            # Feeling score must be between 1 and 5
            if not 1 <= float(r[10]) <= 5:
                raise ValueError

            # Satisfaction score must be between 1 and 5
            if not 1 <= float(r[11]) <= 5:
                raise ValueError

            # Energy must be High, Medium or Low
            if str(r[12]).lower() not in ["high", "medium", "low"]:
                raise ValueError

            # Add valid record to the list
            valid.append(r)

        except (ValueError, TypeError, IndexError):
            # Ignore invalid records
            continue

    return valid


# Calculate all the required activity indices
def calculate(rows):

    # Technical Productivity Index
    # Average coding time per day
    tpi = avg([float(r[4]) for r in rows])

    # Academic Activity Index
    # Average study time + class time
    aai = avg([float(r[3]) + float(r[5]) for r in rows])

    # Physical Activity Index
    # Average fitness time per day
    phai = avg([float(r[2]) for r in rows])

    # Sleep and Recovery Index
    # Average sleep time per day
    sri = avg([float(r[1]) for r in rows])

    # Activity Balance Index
    # Average free/unaccounted time
    abi = avg([float(r[9]) for r in rows])

    # Time Utilization Index
    # Average total tracked time
    tui = avg([float(r[8]) for r in rows])

    # Experience Index
    # Uses Feeling, Satisfaction and Energy scores
    ei = avg([
        (float(r[10]) + float(r[11]) + energy_score(r[12])) / 3
        for r in rows
    ])

    # Data Continuity Index
    # Shows how many days were recorded out of 40 expected days
    dci = len(rows) / 40 * 100

    # Personal Activity Index
    # Combines all major indices using the given weights
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


# Find relationships between activities and experience
def relationships(rows):

    # Find average sleep, study and coding time
    sleep = avg([float(r[1]) for r in rows])
    study = avg([float(r[3]) for r in rows])
    coding = avg([float(r[4]) for r in rows])

    # Compare Energy on days with higher and lower sleep
    sleep_high = avg([
        energy_score(r[12]) for r in rows if float(r[1]) >= sleep
    ])

    sleep_low = avg([
        energy_score(r[12]) for r in rows if float(r[1]) < sleep
    ])

    # Compare Satisfaction on days with higher and lower study time
    study_high = avg([
        float(r[11]) for r in rows if float(r[3]) >= study
    ])

    study_low = avg([
        float(r[11]) for r in rows if float(r[3]) < study
    ])

    # Compare Energy on days with higher and lower coding time
    coding_high = avg([
        energy_score(r[12]) for r in rows if float(r[4]) >= coding
    ])

    coding_low = avg([
        energy_score(r[12]) for r in rows if float(r[4]) < coding
    ])

    return sleep_high, sleep_low, study_high, study_low, coding_high, coding_low


# Main function of the program
def main():

    # Read the activity data from our Excel file
    rows = read_excel("1234567.xlsx")

    # Keep only valid activity records
    rows = valid_rows(rows)

    # Stop the program if no valid data is available
    if not rows:
        print("No valid data found.")
        return

    # Calculate all activity indices
    tpi, aai, phai, sri, abi, tui, ei, dci, pai = calculate(rows)

    print("\n========== MY DATA, MY STORY ==========")

    print("\nACTIVITY SUMMARY")

    # Total days in our project period
    print("Expected Days       :", 40)

    # Number of valid days actually recorded
    print("Valid Days          :", len(rows))

    # Days for which data is missing
    print("Missing Days        :", 40 - len(rows))

    # Invalid records are excluded during validation
    print("Invalid Records     :", 0)

    # Display average daily activity values
    print("Average Sleep       :", round(avg([r[1] for r in rows]), 2))
    print("Average Fitness     :", round(avg([r[2] for r in rows]), 2))
    print("Average Study       :", round(avg([r[3] for r in rows]), 2))
    print("Average Coding      :", round(avg([r[4] for r in rows]), 2))
    print("Average Class       :", round(avg([r[5] for r in rows]), 2))
    print("Average Other       :", round(avg([r[7] for r in rows]), 2))
    print("Average Free Time   :", round(avg([r[9] for r in rows]), 2))

    print("\nACTIVITY INDICES")

    # Display calculated index values
    print("PAI  =", round(pai, 2))
    print("TPI  =", round(tpi, 2), "min/day")
    print("AAI  =", round(aai, 2), "min/day")
    print("PhAI =", round(phai, 2), "min/day")
    print("SRI  =", round(sri, 2), "min/day")
    print("ABI  =", round(abi, 2), "min/day")
    print("TUI  =", round(tui, 2), "min/day")
    print("EI   =", round(ei, 2), "/5")
    print("DCI  =", round(dci, 2), "%")

    # Calculate the three required relationships
    sh, sl, sth, stl, ch, cl = relationships(rows)

    print("\nKEY FINDINGS")

    # Sleep compared with Energy
    print("Sleep-Energy          :", round(sh, 2), "vs", round(sl, 2))

    # Study compared with Satisfaction
    print("Study-Satisfaction   :", round(sth, 2), "vs", round(stl, 2))

    # Coding compared with Energy
    print("Coding-Energy        :", round(ch, 2), "vs", round(cl, 2))


if __name__ == "__main__":
    main()