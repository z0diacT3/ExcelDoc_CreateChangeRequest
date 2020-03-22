# Change Request Final (from V2) includes resource name selection
# Created by Eugene Davies on 20 March 2020
import openpyxl

wb = openpyxl.load_workbook("ChangeRequestv1.xlsx")
changeSheet = wb["Change"]
infosheet = wb["Info"]
newFile = input("New File Name: ") + ".xlsx"


def get_resource(team, begin_row, end_row, begin_column, end_column):
    resource = ""
    print(f"Choose the {team} resource (Y/N)")
    for row in infosheet.iter_rows(min_row=begin_row, min_col=begin_column, max_col=end_column, max_row=end_row):
        for cell in row:
            if cell.value != None:
                if resource == "":
                    answer = input(f"{cell.value}: ").lower()
                    if answer == "y":
                        resource = cell.value
            else:
                if cell.value == None and resource == "":
                    resource = input(f"Enter {team} resource name: ").title()
    return resource


    # Get info from input and file (prompt for selection)
correctDate = False
while not correctDate:
    startDate = input("Enter Start Date (dd/mm/yyyy): ")
    startTime = input("Enter Start Time (HH:mm): ")
    endDate = input("Enter End Date (dd/mm/yyyy): ")
    endTime = input("Enter End Time (HH:mm): ")
    seDate = (f"{startDate} @ {startTime} to {endDate} @ {endTime}")
    inputDate = input(f"** {seDate} **\n\tIs this date correct Y/N: ").lower()
    if inputDate == "y":
        correctDate = True
        changeSheet["B4"] = seDate

deploymentPlan = input("Has the deployment plan been completed? (Y/N): ").lower()
while deploymentPlan != "y" and deploymentPlan != "n":
    deploymentPlan = input("Has the deployment plan been completed? (Y/N): ").lower()
if deploymentPlan == "y":
    changeSheet["B5"] = "Yes - Change may be logged"
else:
    changeSheet["B5"] = "No - Change may no be logged yet"

migration = input("Migration Strategy?\nA: Workload\nB: Lift and Shift\nEnter A or B: ").lower()
while migration != "a" and migration != "b":
    migration = input("Migration Strategy?\nA: Workload\nB: Lift and Shift\nEnter A or B: ").lower()
if migration == "a":
    changeSheet["B6"] = "Workload"
else:
    changeSheet["B6"] = "Lift and Shift"

changeImplementer = input("MS Change implementer: ").title()
changeSheet["B7"] = changeImplementer
changeImplementerBackup = input("MS Change implementer backup: ").title()
changeSheet["B25"] = changeImplementerBackup
downTime = input("Downtime: ")
changeSheet["B22"] = downTime
impact = input("Change Impact: ").capitalize()
changeSheet["B23"] = impact
risk = input("Change Risk (Low/Med/High): ").capitalize()
changeSheet["B24"] = risk

soShared = input("SO Shared Hosting (Y/N): ").lower()
while soShared != "y" and soShared != "n":
    soShared = input("SO Shared Hosting (Y/N): ").lower()
if soShared == "y":
    changeSheet["B8"] = "Yes"
    changeSheet["C8"] = get_resource("SO Shared Hosting", 8, 8, 2, 5)
else:
    changeSheet["B8"] = "No"

soRetail = input("SO Retail Hosting (Y/N): ").lower()
while soRetail != "y" and soRetail != "n":
    soRetail = input("SO Retail Hosting (Y/N): ").lower()
if soRetail == "y":
    changeSheet["B9"] = "Yes"
    changeSheet["C9"] = get_resource("SO Retail Hosting", 9, 9, 2, 5)
else:
    changeSheet["B9"] = "No"

soCCG = input("SO CC&G Hosting (Y/N): ").lower()
while soCCG != "y" and soCCG != "n":
    soCCG = input("SO CC&G Hosting (Y/N): ").lower()
if soCCG == "y":
    changeSheet["B10"] = "Yes"
    changeSheet["C10"] = get_resource("SO Retail Hosting", 10, 10, 2, 5)
else:
    changeSheet["B10"] = "No"

wintel = input("Wintel Africa (Y/N): ").lower()
while wintel != "y" and wintel != "n":
    wintel = input("Wintel Africa (Y/N): ").lower()
if wintel == "y":
    changeSheet["B11"] = "Yes"
    changeSheet["C11"] = get_resource("SO Retail Hosting", 11, 11, 2, 5)
else:
    changeSheet["B11"] = "No"

soKZN = input("SO KZN (Y/N): ").lower()
while soKZN != "y" and soKZN != "n":
    soKZN = input("SO KZN (Y/N): ").lower()
if soKZN == "y":
    changeSheet["B12"] = "Yes"
    changeSheet["C12"] = get_resource("SO Retail Hosting", 12, 12, 2, 5)
else:
    changeSheet["B12"] = "No"

f5 = input("f5 Support (Y/N): ").lower()
while f5 != "y" and f5 != "n":
    f5 = input("f5 Support (Y/N): ").lower()
if f5 == "y":
    changeSheet["B13"] = "Yes"
    changeSheet["C13"] = get_resource("SO Retail Hosting", 13, 13, 2, 5)
else:
    changeSheet["B13"] = "No"

sqlSupport = input("SQL Support (Y/N): ").lower()
while sqlSupport != "y" and sqlSupport != "n":
    sqlSupport = input("SQL Support (Y/N): ").lower()
if sqlSupport == "y":
    changeSheet["B14"] = "Yes"
    changeSheet["C14"] = get_resource("SO Retail Hosting", 14, 14, 2, 5)
else:
    changeSheet["B14"] = "No"

dnsSO = input("DNS / SO Infra Support (Y/N): ").lower()
while dnsSO != "y" and dnsSO != "n":
    dnsSO = input("DNS / SO Infra Support (Y/N): ").lower()
if dnsSO == "y":
    changeSheet["B15"] = "Yes"
    changeSheet["C15"] = get_resource("SO Retail Hosting", 15, 15, 2, 5)
else:
    changeSheet["B15"] = "No"

firewall = input("Firewall Support (Y/N): ").lower()
while firewall != "y" and firewall != "n":
    firewall = input("Firewall Support (Y/N): ").lower()
if firewall == "y":
    changeSheet["B16"] = "Yes"
    changeSheet["C16"] = get_resource("SO Retail Hosting", 16, 16, 2, 5)
else:
    changeSheet["B16"] = "No"

tsm = input("TSM Support (Backups) (Y/N): ").lower()
while tsm != "y" and tsm != "n":
    tsm = input("TSM Support (Backups) (Y/N): ").lower()
if tsm == "y":
    changeSheet["B17"] = "Yes"
    changeSheet["C17"] = get_resource("SO Retail Hosting", 17, 17, 2, 5)
else:
    changeSheet["B17"] = "No"

aix = input("AIX Scheduling (Y/N): ").lower()
while aix != "y" and aix != "n":
    aix = input("AIX Scheduling (Y/N): ").lower()
if aix == "y":
    changeSheet["B18"] = "Yes"
    changeSheet["C18"] = get_resource("SO Retail Hosting", 18, 18, 2, 5)
else:
    changeSheet["B18"] = "No"

missciss = input("MISS and/or CISS (Y/N): ").lower()
while missciss != "y" and missciss != "n":
    missciss = input("MISS and/or CISS (Y/N): ").lower()
if missciss == "y":
    changeSheet["B19"] = "Yes"
    changeSheet["C19"] = get_resource("SO Retail Hosting", 19, 19, 2, 5)
else:
    changeSheet["B19"] = "No"

proxyemail = input("Proxy, Email and Web Security (Y/N): ").lower()
while proxyemail != "y" and proxyemail != "n":
    proxyemail = input("Proxy, Email and Web Security (Y/N): ").lower()
if proxyemail == "y":
    changeSheet["B20"] = "Yes"
    changeSheet["C20"] = get_resource("SO Retail Hosting", 20, 20, 2, 5)
else:
    changeSheet["B20"] = "No"

datapower = input("Datapower (Y/N): ").lower()
while datapower != "y" and datapower != "n":
    datapower = input("Datapower (Y/N): ").lower()
if datapower == "y":
    changeSheet["B21"] = "Yes"
    changeSheet["C21"] = get_resource("SO Retail Hosting", 21, 21, 2, 5)
else:
    changeSheet["B21"] = "No"

wb.save(newFile)
