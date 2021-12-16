import cx_Oracle
import datetime
import xlsxwriter
import sys


def convertStr(stringToCheck):
    if stringToCheck is None:
        return ""
    return str(stringToCheck)


def convertStrToDate(stringToCheck, csuID):
    if (stringToCheck == None) or (stringToCheck == ""):
        return None
    else:
        try:
            dateReturn = datetime.datetime.strptime(stringToCheck, "%m/%d/%Y")
        except:
            try:
                dateReturn = datetime.datetime.strptime(stringToCheck, "%d-%b-%Y")
            except:
                try:
                    dateReturn = datetime.datetime.strptime(stringToCheck, "%d-%b-%y")
                except:
                    print("Failed string to date conversion:", str(stringToCheck), "(" + str(findCSUID(csuID)) + ")")
                    return None

    return dateReturn


def writeImportFile(dict, fileName):
    dateToday = datetime.date.today()

    if (sys.argv[1] == "PreCensus"):
        writeFile = open(FILEPATH + dateToday.strftime("%Y%m%d") + "-" + fileName + "-PreCensus.txt", "wt")
    else:
        writeFile = open(FILEPATH + dateToday.strftime("%Y%m%d") + "-" + fileName + "-PostCensus.txt", "wt")

    for key, student in dict.items():
        if student.newInsurancePlan != None:
            writeFile.write(convertStr(student.csuID) + "\t")
            writeFile.write(convertStr(student.dateOfBirth) + "\t")
            writeFile.write(convertStr(student.newInsurancePlan) + "\t")
            writeFile.write(convertStr(student.newInsuranceEffectiveDate) + "\t")
            writeFile.write(convertStr(student.newInsuranceExpirationDate) + "\t")
            writeFile.write("\t")
            # writeFile.write(convertStr(GROUP_NUMBER))
            writeFile.write(convertStr(student.groupNumber) + "\t")
            # writeFile.write("\t" + "1")	# Set the Priority
            writeFile.write("\n")

    writeFile.close()


def createWorkbook(dict, fields, fileName):
    workbook = xlsxwriter.Workbook(FILEPATH + datetime.date.today().strftime("%Y%m%d") + "-" + fileName + ".xlsx")
    worksheet = workbook.add_worksheet()

    row = 0
    column = 0
    for x in range(0, len(fields)):
        worksheet.write(row, column, fields[x][1])
        column += 1

    row += 1
    for key, value in dict.items():
        column = 0
        for x in range(0, len(fields)):
            worksheet.write(row, column, convertStr(value.__dict__[fields[x][0]]))
            column += 1
        row += 1

    workbook.close()


class studentData:
    def __init__(self, pidm):
        self.pidm = pidm
        self.csuID = None
        self.dateOfBirth = None
        self.pidm = None
        self.hasHEALTHINS = None
        self.insuranceBilled = None
        self.internationalStatus = None
        self.creditsRI = None
        self.campus = None
        self.publicHealth = None
        self.employeeBenefits = None
        self.thirdPartyWaiver = None
        self.thirdStartDate = None
        self.thirdEndDate = None
        self.internationalWaiver = None
        self.internationalSource = None
        self.internationalStartDate = None
        self.internationalEndDate = None
        self.healthNetworkWaiver = None
        self.csuhnStartDate = None
        self.csuhnEndDate = None
        self.monthly = None
        self.primaryMajor = None
        self.email = None

        self.studentHealthInsurancePlan = None
        self.insuranceEffectiveDate = None
        self.insuranceExpirationDate = None
        self.insuranceMultiplePlans = None
        self.matchedDataCode = None
        self.internationalPatient = None
        self.academicMajor = None

        self.newInsurancePlan = None
        self.newInsuranceEffectiveDate = None
        self.newInsuranceExpirationDate = None

        self.groupNumber = None
        self.SISPending = None

    fields = {0: ["csuID", "CSU ID"],
              1: ["pidm", "PIDM"],
              2: ["dateOfBirth", "DOB"],
              3: ["hasHEALTHINS", "Data Code"],
              4: ["insuranceBilled", "Health Insurance Billed"],
              5: ["internationalStatus", "International Status"],
              6: ["creditsRI", "RI Credits"],
              7: ["campus", "Campus"],
              8: ["publicHealth", "CSPH"],
              9: ["employeeBenefits", "Eligible Employee"],
              10: ["thirdPartyWaiver", "3rd Party Waiver"],
              11: ["thirdStartDate", "3rd Start Date"],
              12: ["thirdEndDate", "3rd End Date"],
              13: ["internationalWaiver", "OIP Waiver"],
              14: ["internationalSource", "OIP Source"],
              15: ["internationalStartDate", "OIP Start Date"],
              16: ["internationalEndDate", "OIP End Date"],
              17: ["healthNetworkWaiver", "CSUHN Waiver"],
              18: ["csuhnStartDate", "CSUHN Start Date"],
              19: ["csuhnEndDate", "CSUHN End Date"],
              20: ["monthly", "Monthly Charge"],
              21: ["studentHealthInsurancePlan", "Student Insurance"],
              22: ["insuranceEffectiveDate", "Effective Date"],
              23: ["insuranceExpirationDate", "Expiration Date"],
              24: ["insuranceMultiplePlans", "Multiple Plans"],
              25: ["matchedDataCode", "Found Data Code"],
              26: ["newInsurancePlan", "New Plan"],
              27: ["newInsuranceEffectiveDate", "New Effective Date"],
              28: ["newInsuranceExpirationDate", "New Expiration Date"],
              29: ["internationalPatient", "InternationalFlag"],
              30: ["academicMajor", "Academic Major"],
              31: ["groupNumber", "Group Number"],
              32: ["primaryMajor", "Primary Major"],
              33: ["email", "E-Mail"],
              }


def findSISPending(dict):
    files = ["N:\\NET\\Source Code\\Insurance\\Dashboard-PA-CSU_18-19_Fall_Waiver-09-11-2018.txt",
             "N:\\NET\\Source Code\\Insurance\\Dashboard-PE-CSU_18-19_Fall_Waiver-09-11-2018.txt"]

    # for filename in files:

    for file in files:

        reader = open(file, "r")

        # This reads the file given to us by the Ins team, flags the SIS pending as True, skips them
        for line in reader:
            # line = line.replace("\n", "")
            studentID = line.split(",")[0]
            if studentID in dict:
                print("changing " + studentID + "to true")
                dict[studentID].SISPending = True

    return dict


def getDataCodeData(internationalOnly):
    dict = {}

    if (len(sys.argv) == 1):
        print("No argument provided")
        exit

    # Pulling from banner - ODS was not current

    sql = (
                "SELECT SWRGPCD_PIDM, NULL, dc.SWRGPCD_ATTR1, dc.SWRGPCD_ATTR2, dc.SWRGPCD_ATTR3, dc.SWRGPCD_ATTR4, dc.SWRGPCD_ATTR5, dc.SWRGPCD_ATTR6, dc.SWRGPCD_ATTR7, dc.SWRGPCD_ATTR8, dc.SWRGPCD_ATTR9, dc.SWRGPCD_ATTR10, "
                "dc.SWRGPCD_ATTR11, dc.SWRGPCD_ATTR12, dc.SWRGPCD_ATTR13, dc.SWRGPCD_ATTR14, dc.SWRGPCD_ATTR15, dc.SWRGPCD_ATTR16, dc.SWRGPCD_ATTR17, dc.SWRGPCD_ATTR18, dc.SWRGPCD_ATTR19, dc.SWRGPCD_ATTR20, NULL, NULL "
                "FROM SWRGPCD dc "
                "WHERE SWRGPCD_GPCD_CODE = 'HEALTHINS' "
                "AND SWRGPCD_TERM = '" + TERM + "' "
                # "and swrgpcd_PIDM = '11743608' "

                )

    if internationalOnly == True:
        sql += "AND dc.SWRGPCD_ATTR1 = 'ASHI' "

    CreateDSN = cx_Oracle.makedsn(BANProdServer, BANProdServerPort, BANProdServerDBName)
    BannerConnection = cx_Oracle.Connection(BANProdServerUser, BANProdServerPass, CreateDSN)
    BannerCursor = cx_Oracle.Cursor(BannerConnection)
    BannerCursor.execute(sql)
    results = BannerCursor.fetchall()
    BannerConnection.close()

    print("Length of results: ", len(results))
    if results:
        for row in results:
            student = studentData(row[0])
            student.pidm = row[0]
            student.dateOfBirth = row[1]
            student.hasHEALTHINS = True
            student.insuranceBilled = row[2]
            student.internationalStatus = row[3]
            student.creditsRI = row[4]
            student.campus = row[5]
            student.publicHealth = row[6]
            student.employeeBenefits = row[7]
            student.thirdPartyWaiver = row[8]
            student.thirdStartDate = row[9]
            student.thirdEndDate = row[10]
            student.internationalWaiver = row[11]
            student.internationalSource = row[12]
            student.internationalStartDate = row[13]
            student.internationalEndDate = row[14]
            student.healthNetworkWaiver = row[15]
            student.csuhnStartDate = row[16]
            student.csuhnEndDate = row[17]
            student.monthly = row[18]
            student.primaryMajor = row[22]
            student.email = row[23]
            dict[student.pidm] = student

    print("getDataCodeData:", len(dict))
    return dict


def getStudentHealthInsurancePlanData():
    dict = {}

    sql = ("SELECT PAT_Institutionid, PAT_Date_Of_Birth, PPLN_Plan_Id_Text, PPLN_Effective_Date, PPLN_Expiration_Date "
           "FROM V_Patient_Plan "
           "JOIN V_Patient ON PPLN_Patient_Id = PAT_Patient_Id "
           "WHERE PPLN_Plan_Id_Text = '" + PLAN_NAME + "' "
           # "AND PAT_institutionid = '832562998' "
           )

    CreateDSN = cx_Oracle.makedsn(PyraMEDServerIP, PyraMEDServerPort, PyraMEDServerDBName)
    PyraMEDConnection = cx_Oracle.Connection(PyraMEDServerUser, PyraMEDServerPass, CreateDSN)
    Cursor = cx_Oracle.Cursor(PyraMEDConnection)
    Cursor.execute(sql)
    results = Cursor.fetchall()
    PyraMEDConnection.close()

    if results:
        for row in results:
            pidm = getPIDM(row[0])
            if row[0] not in dict:
                student = studentData(pidm)
                student.pidm = pidm
                student.csuID = row[0]
                student.dateOfBirth = row[1]
                student.studentHealthInsurancePlan = row[2]
                student.insuranceEffectiveDate = row[3]
                student.insuranceExpirationDate = row[4]
                dict[student.pidm] = student
            else:
                print("Individual has multiple " + PLAN_NAME + " plans:", row[0])
                student = studentData(pidm)
                student.pidm = pidm
                student.csuID = row[0]
                student.dateOfBirth = row[1]
                student.studentHealthInsurancePlan = row[2]
                student.insuranceEffectiveDate = row[3]
                student.insuranceExpirationDate = row[4]
                student.insuranceMultiplePlans = True
                dict[student.pidm] = student

    print("getStudentHealthInsurancePlanData:", len(dict))
    # print ("STUDENT IN QUESTION", dict['831630164'])
    return dict


def internationalDataFromPyraMED(dict):
    dictReturn = {}
    print("Length of dict going in to international data: ", len(dict))
    sql = ("SELECT PAT_International_Patient, PAT_Academic_Major "
           "FROM V_Patient "
           "WHERE PAT_Institutionid = :searchID "
           "AND PAT_Date_Of_Birth = :searchDOB "
           )

    CreateDSN = cx_Oracle.makedsn(PyraMEDServerIP, PyraMEDServerPort, PyraMEDServerDBName)
    PyraMEDConnection = cx_Oracle.Connection(PyraMEDServerUser, PyraMEDServerPass, CreateDSN)
    Cursor = cx_Oracle.Cursor(PyraMEDConnection)

    for key, student in dict.items():
        Cursor.execute(sql, {"searchID": student.csuID, "searchDOB": student.dateOfBirth})
        results = Cursor.fetchall()

        if results:
            if len(results) > 1:
                print("Individual has multiple PyraMED accounts", student.csuID)
            else:
                for row in results:
                    student.internationalPatient = row[0]
                    student.academicMajor = row[1]

        if student.internationalPatient == True:
            student.groupNumber = INTERNATIONAL_GROUP_NUMBER
        elif (student.academicMajor != None) and (student.academicMajor[:4] == "INTO"):
            student.groupNumber = INTERNATIONAL_GROUP_NUMBER
        else:
            student.groupNumber = DOMESTIC_GROUP_NUMBER

        dictReturn[student.csuID] = student

    PyraMEDConnection.close()

    print("internationalDataFromPyraMED:", len(dictReturn))
    return dictReturn


def compareDataSets(dataCode, shiPlan):
    dictReturn = {}

    for shipKey, shipStudent in shiPlan.items():
        if shipKey in dataCode:
            # print ("Entered")
            if shipStudent.dateOfBirth == dataCode[shipKey].dateOfBirth:
                shipStudent.matchedDataCode = True
            else:
                print("(SHIP) Matched ID but not date of birth:", shipKey)

    for shipKey, shipStudent in shiPlan.items():
        if shipStudent.matchedDataCode != True:
            print("(SHIP) Did not find a data code:", shipKey)
            shipStudent.matchedDataCode = False
            dictReturn[shipStudent.pidm] = shipStudent

    for dcKey, dcStudent in dataCode.items():
        if dcKey in shiPlan:
            if shiPlan[dcKey].insuranceMultiplePlans == True:

                pass  # Has to be dealt with manually
            elif dcStudent.dateOfBirth == shiPlan[dcKey].dateOfBirth:
                student = studentData(dcStudent.pidm)
                student.pidm = dcStudent.pidm
                student.csuID = dcStudent.csuID
                student.dateOfBirth = dcStudent.dateOfBirth
                student.hasHEALTHINS = dcStudent.hasHEALTHINS
                student.insuranceBilled = dcStudent.insuranceBilled
                student.internationalStatus = dcStudent.internationalStatus
                student.creditsRI = dcStudent.creditsRI
                student.campus = dcStudent.campus
                student.publicHealth = dcStudent.publicHealth
                student.employeeBenefits = dcStudent.employeeBenefits
                student.thirdPartyWaiver = dcStudent.thirdPartyWaiver
                student.thirdStartDate = dcStudent.thirdStartDate
                student.thirdEndDate = dcStudent.thirdEndDate
                student.internationalWaiver = dcStudent.internationalWaiver
                student.internationalSource = dcStudent.internationalSource
                student.internationalStartDate = dcStudent.internationalStartDate
                student.internationalEndDate = dcStudent.internationalEndDate
                student.healthNetworkWaiver = dcStudent.healthNetworkWaiver
                student.csuhnStartDate = dcStudent.csuhnStartDate
                student.csuhnEndDate = dcStudent.csuhnEndDate
                student.monthly = dcStudent.monthly
                student.primaryMajor = dcStudent.primaryMajor
                student.email = dcStudent.email
                student.studentHealthInsurancePlan = shiPlan[dcKey].studentHealthInsurancePlan
                student.insuranceEffectiveDate = shiPlan[dcKey].insuranceEffectiveDate
                student.insuranceExpirationDate = shiPlan[dcKey].insuranceExpirationDate
                student.matchedDataCode = shiPlan[dcKey].matchedDataCode
                dictReturn[student.pidm] = student
            else:
                print("(DataCode) Matched ID but not date of birth:", dcKey)
        else:
            student = studentData(dcStudent.pidm)
            student.pidm = dcStudent.pidm
            student.csuID = dcStudent.csuID
            student.dateOfBirth = dcStudent.dateOfBirth
            student.hasHEALTHINS = dcStudent.hasHEALTHINS
            student.insuranceBilled = dcStudent.insuranceBilled
            student.internationalStatus = dcStudent.internationalStatus
            student.creditsRI = dcStudent.creditsRI
            student.campus = dcStudent.campus
            student.publicHealth = dcStudent.publicHealth
            student.employeeBenefits = dcStudent.employeeBenefits
            student.thirdPartyWaiver = dcStudent.thirdPartyWaiver
            student.thirdStartDate = dcStudent.thirdStartDate
            student.thirdEndDate = dcStudent.thirdEndDate
            student.internationalWaiver = dcStudent.internationalWaiver
            student.internationalSource = dcStudent.internationalSource
            student.internationalStartDate = dcStudent.internationalStartDate
            student.internationalEndDate = dcStudent.internationalEndDate
            student.healthNetworkWaiver = dcStudent.healthNetworkWaiver
            student.csuhnStartDate = dcStudent.csuhnStartDate
            student.csuhnEndDate = dcStudent.csuhnEndDate
            student.monthly = dcStudent.monthly
            student.primaryMajor = dcStudent.primaryMajor
            student.email = dcStudent.email
            dictReturn[student.pidm] = student

    print("compareDataSets:", len(dictReturn))
    return dictReturn


def addPlan(student):
    if student.insuranceBilled == "ASHD" or student.healthNetworkWaiver == "VO":
        student.newInsurancePlan = PLAN_NAME
        student.newInsuranceEffectiveDate = EFFECTIVE_DATE
        student.newInsuranceExpirationDate = EXPIRATION_DATE
    # brandNewPlans += 1
    elif student.insuranceBilled == "ASHI":
        student.newInsurancePlan = PLAN_NAME
        student.newInsuranceEffectiveDate = INTERNATIONAL_EFFECTIVE_DATE
        student.newInsuranceExpirationDate = INTERNATIONAL_EXPIRATION_DATE
    # brandNewPlans += 1
    else:
        print(student.csuID, "has an ATTR1_VALUE of something other than ASHD/ASHI:", student.insuranceBilled)


def updatePlan(student):
    brandNewPlans = 0
    extendExistingPlans = 0
    newOnPreviouslyExpiredPlans = 0
    removingPlans = 0
    monthlyOnMonthlyPlans = 0
    monthlyOnTermPlans = 0
    noPlans = 0
    removedPlans = {}

    if student.insuranceBilled == "ASHD" or student.healthNetworkWaiver == "VO":
        # Probably more for Spring, had full term insurance last term, or active insurance, use same effective date, populate new expiration date (extend current plan)
        if (student.insuranceExpirationDate == EFFECTIVE_DATE - datetime.timedelta(days=1)) or (
                student.insuranceExpirationDate > EFFECTIVE_DATE):
            student.newInsurancePlan = PLAN_NAME
            student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
            student.newInsuranceExpirationDate = EXPIRATION_DATE
            extendExistingPlans += 1
        # Termed plan, reactiveate using effective and expiration date
        elif student.insuranceEffectiveDate == student.insuranceExpirationDate:
            student.newInsurancePlan = PLAN_NAME
            student.newInsuranceEffectiveDate = EFFECTIVE_DATE
            student.newInsuranceExpirationDate = EXPIRATION_DATE
            newOnPreviouslyExpiredPlans += 1
        # Not sure if we need this based on the first condition checked above.
        elif student.insuranceEffectiveDate < EXPIRATION_DATE:
            if student.insuranceExpirationDate == EFFECTIVE_DATE - datetime.timedelta(days=1):
                print("CHECK THESE:", student.csuID)
                student.newInsurancePlan = PLAN_NAME
                student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                student.newInsuranceExpirationDate = EXPIRATION_DATE
                newOnPreviouslyExpiredPlans += 1
            # not sure what this is catching, eitheer
            else:
                print("CHECK THESE:", student.csuID)
                student.newInsurancePlan = PLAN_NAME
                student.newInsuranceEffectiveDate = EFFECTIVE_DATE
                student.newInsuranceExpirationDate = EXPIRATION_DATE
                newOnPreviouslyExpiredPlans += 1
    elif student.insuranceBilled == "ASHI":
        # Same as above checks but with internationals (could we combine these in to one "if" since international and domestic start on the same day??)
        if (student.insuranceExpirationDate == INTERNATIONAL_EFFECTIVE_DATE - datetime.timedelta(days=1)) or (
                student.insuranceExpirationDate > INTERNATIONAL_EFFECTIVE_DATE):
            student.newInsurancePlan = PLAN_NAME
            student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
            student.newInsuranceExpirationDate = INTERNATIONAL_EXPIRATION_DATE
            extendExistingPlans += 1
        elif student.insuranceEffectiveDate == student.insuranceExpirationDate:
            student.newInsurancePlan = PLAN_NAME
            student.newInsuranceEffectiveDate = INTERNATIONAL_EFFECTIVE_DATE
            student.newInsuranceExpirationDate = INTERNATIONAL_EXPIRATION_DATE
            newOnPreviouslyExpiredPlans += 1
        elif student.insuranceEffectiveDate < INTERNATIONAL_EFFECTIVE_DATE:
            if student.insuranceExpirationDate == INTERNATIONAL_EFFECTIVE_DATE - datetime.timedelta(days=1):
                print("CHECK THESE:", student.csuID)
                student.newInsurancePlan = PLAN_NAME
                student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                student.newInsuranceExpirationDate = INTERNATIONAL_EXPIRATION_DATE
                newOnPreviouslyExpiredPlans += 1
            else:
                print("CHECK THESE:", student.csuID)
                student.newInsurancePlan = PLAN_NAME
                student.newInsuranceEffectiveDate = INTERNATIONAL_EFFECTIVE_DATE
                student.newInsuranceExpirationDate = INTERNATIONAL_EXPIRATION_DATE
                newOnPreviouslyExpiredPlans += 1
        else:
            print(student.csuID, "has an ATTR1_VALUE of something other than ASHD/ASHI:", student.insuranceBilled)
    elif (student.insuranceBilled == None) and (student.monthly != None):
        if student.monthly == "MONTHLY":
            monthlyOnMonthlyPlans += 1
    else:
        monthlyOnTermPlans += 1


def determinePlanUpdates(dict):
    dictReturn = {}
    brandNewPlans = 0
    extendExistingPlans = 0
    newOnPreviouslyExpiredPlans = 0
    removingPlans = 0
    monthlyOnMonthlyPlans = 0
    monthlyOnTermPlans = 0
    noPlans = 0
    removedPlans = {}

    for key, student in dict.items():

        if student.csuID in HARD_CODE_EXPIRES:
            student.newInsurancePlan = PLAN_NAME
            student.newInsuranceEffectiveDate = EFFECTIVE_DATE
            student.newInsuranceExpirationDate = EFFECTIVE_DATE
            noPlans += 1
        elif student.csuID in HARD_CODE_ADDITIONS:
            student.newInsurancePlan = PLAN_NAME
            student.newInsuranceEffectiveDate = INTERNATIONAL_EFFECTIVE_DATE
            student.newInsuranceExpirationDate = INTERNATIONAL_EXPIRATION_DATE
            extendExistingPlans += 1

        # Adding this Feb 2021, lots of students with plan but no data code. Had plan in fall but not enrolled for Spring and had dates pulled forward. Need to pull back expiration date
        elif student.matchedDataCode == False and TERM[-2:] == "10":
            student.newInsurancePlan = PLAN_NAME
            student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
            if student.insuranceExpirationDate <= PRIOR_EXPIRATION_DATE:
                student.newInsuranceExpirationDate = student.insuranceExpirationDate
            elif student.insuranceExpirationDate > PRIOR_EXPIRATION_DATE:
                student.newInsuranceExpirationDate = PRIOR_EXPIRATION_DATE
                removedPlans[student.csuID] = student
                removingPlans += 1

        # DO ONE THING IF PRECENSUS
        elif (sys.argv[1] == "PreCensus"):
            # student has OI, VO or ASHI and no student health plan, add it with an ASHD/ASHI
            # elif (student.insuranceBilled != None) and (student.studentHealthInsurancePlan == None):
            if (((student.insuranceBilled != None) and (
                    (student.thirdPartyWaiver == "OI") or (student.healthNetworkWaiver == "VO") or (
                    student.insuranceBilled == "ASHI"))) and (student.studentHealthInsurancePlan == None)):
                addPlan(student)
                brandNewPlans += 1
            # elif (student.insuranceBilled != None) and (student.studentHealthInsurancePlan != None):
            # student has OI, VO or ASHI and has a student health plan, change the dates
            elif (((student.insuranceBilled != None) and (
                    (student.thirdPartyWaiver == "OI") or (student.healthNetworkWaiver == "VO") or (
                    student.insuranceBilled == "ASHI"))) and (student.studentHealthInsurancePlan != None)):
                print("H!!ERE")
                updatePlan(student)
            elif ((student.insuranceBilled == None) and (student.monthly != None)):
                if student.monthly == "MONTHLY":
                    monthlyOnMonthlyPlans += 1
                else:
                    monthlyOnTermPlans += 1
            # Adding this line in to add plan with VO without ASHD  	Might be a bad idea
            elif (student.healthNetworkWaiver == "VO"):
                if student.studentHealthInsurancePlan != None:
                    updatePlan(student)
                else:
                    addPlan(student)
            if TERM[-2:] == "10":
                # run this only in Spring. This will expire plans for students who had SHIP in fall but submitted a waiver for Spring. So their exp date should be 12/31
                if ((
                        student.thirdPartyWaiver == 'AP' or student.healthNetworkWaiver == "AP" or student.internationalWaiver == "AP") and student.thirdPartyWaiver != "OI" and student.healthNetworkWaiver != "VO" and student.insuranceBilled == None and TERM[
                                                                                                                                                                                                                                                             -2:] == '10' and student.insuranceExpirationDate == EXPIRATION_DATE):
                    if student.insuranceEffectiveDate == EFFECTIVE_DATE and student.insuranceExpirationDate != student.insuranceEffectiveDate:
                        student.newInsurancePlan = PLAN_NAME
                        student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                        student.newInsuranceExpirationDate = student.insuranceEffectiveDate
                        removingPlans += 1
                        removedPlans[student.csuID] = student
                    elif (
                            student.insuranceEffectiveDate == PRIOR_EFFECTIVE_DATE and student.insuranceEffectiveDate != student.insuranceExpirationDate):
                        student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                        student.newInsurancePlan = PLAN_NAME
                        student.newInsuranceExpirationDate = PRIOR_EXPIRATION_DATE
                        removingPlans += 1
                        removedPlans[student.csuID] = student

                # Life event, leave it alone
                elif (
                        student.insuranceEffectiveDate != EFFECTIVE_DATE and student.insuranceEffectiveDate != PRIOR_EFFECTIVE_DATE) or (
                        student.insuranceExpirationDate == student.insuranceEffectiveDate):
                    pass
            # else:
            # print("NEW DATE: ", student.newInsuranceExpirationDate)
            # student.newInsurancePlan = PLAN_NAME
            # student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
            # student.newInsuranceExpirationDate = PRIOR_EXPIRATION_DATE
            # removingPlans += 1

            # Only run this on or after the first day of classes, will remove plan from OI with no credits

            elif (
                    student.thirdPartyWaiver == "OI" and student.creditsRI == None and student.insuranceBilled == None and student.studentHealthInsurancePlan != None and student.healthNetworkWaiver != 'AP' and student.internationalWaiver != 'AP'):
                #	if TERM[-2:] == "90":
                #		if (student.insuranceEffectiveDate == student.insuranceExpirationDate):
                #			pass
                #		if (student.insuranceEffectiveDate != EFFECTIVE_DATE):
                #			pass
                # Life event, don't do anything
                # student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                # student.newInsuranceExpirationDate = student.insuranceEffectiveDate
                #		else:
                #			student.newInsurancePlan = PLAN_NAME
                #			student.newInsuranceEffectiveDate = EFFECTIVE_DATE
                #			student.newInsuranceExpirationDate = EFFECTIVE_DATE
                #			removingPlans += 1
                #			removedPlans[student.csuID] = student

                if TERM[-2:] == "10":
                    # @rint ("Credits: ", student.creditsRI)
                    if (student.insuranceEffectiveDate == student.insuranceExpirationDate):
                        pass
                    elif (
                            student.insuranceEffectiveDate != EFFECTIVE_DATE and student.insuranceEffectiveDate != PRIOR_EFFECTIVE_DATE):
                        pass
                    elif (
                            student.insuranceEffectiveDate == PRIOR_EFFECTIVE_DATE and student.insuranceExpirationDate == PRIOR_EXPIRATION_DATE):
                        pass
                    # Commenting out for COVID spring 2020 - leave everyone who had the plan in fall, only add in oi's and vo's
                    elif (student.insuranceEffectiveDate == PRIOR_EFFECTIVE_DATE):
                        student.newInsurancePlan = PLAN_NAME
                        student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                        student.newInsuranceExpirationDate = PRIOR_EXPIRATION_DATE
                        removingPlans += 1
                        removedPlans[student.csuID] = student
                    # END OF PART I COMMENTED OUT FOR COVID SPRING
                    elif (student.insuranceEffectiveDate == EFFECTIVE_DATE):
                        student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                        student.newInsuranceExpirationDate = student.insuranceEffectiveDate
                    else:
                        print("Weird dates, check this: ", student.csuID)

            # take this out for spring go off term code....?? if
            elif (student.insuranceBilled == None) and (student.studentHealthInsurancePlan != None):

                if student.insuranceEffectiveDate == student.insuranceExpirationDate:
                    noPlans += 1
                # Take this out for spring
                # Term plan if they no longer have an ASHI/ASHD and have an active plan in PyraMED
                # elif (student.insuranceEffectiveDate == EFFECTIVE_DATE) and (student.insuranceExpirationDate == EXPIRATION_DATE):
                #	student.newInsurancePlan = PLAN_NAME
                #	student.newInsuranceEffectiveDate = EFFECTIVE_DATE
                #	student.newInsuranceExpirationDate = EFFECTIVE_DATE
                #	removingPlans += 1
                #	removedPlans[student.csuID] = student
                # take out for spring

                # elif (student.insuranceEffectiveDate == INTERNATIONAL_EFFECTIVE_DATE) and (student.insuranceExpirationDate == INTERNATIONAL_EXPIRATION_DATE):
                #	student.newInsurancePlan = PLAN_NAME
                #	student.newInsuranceEffectiveDate = INTERNATIONAL_EFFECTIVE_DATE
                #	student.newInsuranceExpirationDate = INTERNATIONAL_EFFECTIVE_DATE
                #	removingPlans += 1
                #	removedPlans[student.csuID] = student
                # take out for spring
                # Had plan in Fall and now doesn't need it for spring
                # This might be okay for Spring but the above will term the plan
                elif (student.insuranceEffectiveDate == PRIOR_EFFECTIVE_DATE) and (
                        student.insuranceExpirationDate == EXPIRATION_DATE):
                    student.newInsurancePlan = PLAN_NAME
                    student.newInsuranceEffectiveDate = PRIOR_EFFECTIVE_DATE
                    student.newInsuranceExpirationDate = EFFECTIVE_DATE - datetime.timedelta(days=1)
                    removingPlans += 1
                    removedPlans[student.csuID] = student
                elif (student.insuranceEffectiveDate == INTERNATIONAL_PRIOR_EFFECTIVE_DATE) and (
                        student.insuranceExpirationDate == INTERNATIONAL_EXPIRATION_DATE):
                    student.newInsurancePlan = PLAN_NAME
                    student.newInsuranceEffectiveDate = INTERNATIONAL_PRIOR_EFFECTIVE_DATE
                    student.newInsuranceExpirationDate = INTERNATIONAL_EFFECTIVE_DATE - datetime.timedelta(days=1)
                    removingPlans += 1
                    removedPlans[student.csuID] = student
            else:
                noPlans += 1
        # DO ONE THING IF POST CENSUS
        elif (sys.argv[1] == "PostCensus"):

            if ((student.insuranceBilled != None) and (student.studentHealthInsurancePlan == None)):
                addPlan(student)
            elif ((student.insuranceBilled != None) and (student.monthly != None)):
                updatePlan(student)
            elif ((student.insuranceBilled == None) and (student.monthly != None)):
                if student.monthly == "MONTHLY":
                    monthlyOnMonthlyPlans += 1
                else:
                    monthlyOnTermPlans += 1
            elif (student.insuranceBilled == None) and (student.studentHealthInsurancePlan != None):

                # print("entered here insuranceBilled none and has plan " + student.csuID )
                if student.insuranceEffectiveDate == student.insuranceExpirationDate:
                    noPlans += 1

                elif (student.healthNetworkWaiver == "VO"):
                    student.newInsurancePlan = PLAN_NAME
                    student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                    student.newInsuranceExpirationDate = EXPIRATION_DATE
                # Rewriting this line in Feb 2020- This was not catching life event students whose effective date != 8/1
                # elif (student.insuranceEffectiveDate == EFFECTIVE_DATE) and (student.insuranceExpirationDate == EXPIRATION_DATE):
                elif (student.insuranceExpirationDate == EXPIRATION_DATE) and (
                        student.insuranceExpirationDate != student.insuranceEffectiveDate):
                    student.newInsurancePlan = PLAN_NAME
                    if student.insuranceEffectiveDate == EFFECTIVE_DATE:

                        student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                        student.newInsuranceExpirationDate = student.insuranceEffectiveDate
                    # Weird dates probably a life event, leave it alone
                    elif (
                            student.insuranceEffectiveDate != PRIOR_EFFECTIVE_DATE and student.insuranceEffectiveDate != EFFECTIVE_DATE):
                        student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                        student.newInsuranceExpirationDate = student.insuranceExpirationDate
                    # Review this logic, only works in Spring
                    # student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                    # if student.insuranceEffectiveDate > PRIOR_EXPIRATION_DATE:
                    #	student.newInsuranceExpirationDate = student.insuranceEffectiveDate
                    # else:
                    #	student.newInsuranceExpirationDate = PRIOR_EXPIRATION_DATE
                    # removingPlans += 1
                    # removedPlans[student.id] = student
                    # elif (student.insuranceEffectiveDate >= EFFECTIVE_DATE) and (student.insuranceExpirationDate == EXPIRATION_DATE):
                    #	student.newInsurancePlan = PLAN_NAME
                    #	student.newInsuranceEffectiveDate = EFFECTIVE_DATE
                    #	student.newInsuranceExpirationDate = EFFECTIVE_DATE
                    elif (student.insuranceEffectiveDate >= PRIOR_EFFECTIVE_DATE) and (
                            student.insuranceExpirationDate == EXPIRATION_DATE):
                        if student.healthNetworkWaiver == "VO":
                            pass
                        else:
                            student.newInsurancePlan = PLAN_NAME
                            student.newInsuranceEffectiveDate = student.insuranceEffectiveDate
                            student.newInsuranceExpirationDate = PRIOR_EXPIRATION_DATE
                            removingPlans += 1
                            removedPlans[student.csuID] = student
                    elif (student.insuranceEffectiveDate == INTERNATIONAL_EFFECTIVE_DATE) and (
                            student.insuranceExpirationDate == INTERNATIONAL_EXPIRATION_DATE):
                        student.newInsurancePlan = PLAN_NAME
                        student.newInsuranceEffectiveDate = INTERNATIONAL_EFFECTIVE_DATE
                        student.newInsuranceExpirationDate = INTERNATIONAL_EFFECTIVE_DATE
                        removingPlans += 1
                        removedPlans[student.csuID] = student
                    elif (student.insuranceEffectiveDate == PRIOR_EFFECTIVE_DATE) and (
                            student.insuranceExpirationDate == EXPIRATION_DATE):
                        if student.healthNetworkWaiver == "VO":
                            pass
                        else:
                            student.newInsurancePlan = PLAN_NAME
                            student.newInsuranceEffectiveDate = PRIOR_EFFECTIVE_DATE
                            student.newInsuranceExpirationDate = PRIOR_EXPIRATION_DATE
                            removingPlans += 1
                            removedPlans[student.csuID] = student
                    elif (student.insuranceEffectiveDate == INTERNATIONAL_PRIOR_EFFECTIVE_DATE) and (
                            student.insuranceExpirationDate == INTERNATIONAL_EXPIRATION_DATE):
                        student.newInsurancePlan = PLAN_NAME
                        student.newInsuranceEffectiveDate = INTERNATIONAL_PRIOR_EFFECTIVE_DATE
                        student.newInsuranceExpirationDate = INTERNATIONAL_EFFECTIVE_DATE - datetime.timedelta(days=1)
                        removingPlans += 1
                        removedPlans[student.csuID] = student
                # No ASHI/ASHD and had plan last term. THis is fine, pass. MIght only work for Spring
                elif (
                        student.insuranceEffectiveDate == PRIOR_EFFECTIVE_DATE and student.insuranceExpirationDate == PRIOR_EXPIRATION_DATE):
                    pass
                else:
                    print("Weird plan dates: ", student.csuID)
                    print(student.insuranceEffectiveDate, student.insuranceExpirationDate)
            elif (student.insuranceBilled != None) and (student.studentHealthInsurancePlan != None):
                # student.newInsuranceEffectiveDate = EFFECTIVE_DATE
                # student.newInsuranceExpirationDate = EXPIRATION_DATE
                updatePlan(student)
            else:
                noPlans += 1

        dictReturn[student.csuID] = student

    print("Brand New Plans:", brandNewPlans)
    print("Extending Existing Plans:", extendExistingPlans)
    print("Had Expired Plan Previously, New Plans:", newOnPreviouslyExpiredPlans)
    print("Removing Plans:", removingPlans)
    print("Monthly Month-to-Month Plans:", monthlyOnMonthlyPlans)
    print("Monthly Term Plans:", monthlyOnTermPlans)
    print("No Plans:", noPlans)

    print("determinePlanUpdates:", len(dictReturn))
    return dictReturn, removedPlans


def getCSUIDandDOB(dict):
    sql = ("SELECT CSU_ID, BIRTH_DATE, EMAIL  "
           "FROM CSUG_GP_DEMO "
           "WHERE PIDM = :searchID ")

    CreateDSN = cx_Oracle.makedsn(ODSServer, ODSServerPort, ODSServerDBName)
    ODSConnection = cx_Oracle.Connection(ODSServerUser, ODSServerPass, CreateDSN)
    ODSCursor = cx_Oracle.Cursor(ODSConnection)

    for key, student in dict.items():

        ODSCursor.execute(sql, {"searchID": student.pidm})
        results = ODSCursor.fetchall()

        if results:
            for row in results:
                dict[student.pidm].csuID = row[0]
                dict[student.pidm].dateOfBirth = row[1]
                dict[student.pidm].email = row[2]

    print("Length of dict to find CSU IDs: ", len(dict))
    return dict


def getPIDM(csuID):
    sql = ("SELECT PIDM "
           "FROM CSUG_GP_DEMO "
           "WHERE CSU_ID = '" + csuID + "' ")

    CreateDSN = cx_Oracle.makedsn(ODSServer, ODSServerPort, ODSServerDBName)
    ODSConnection = cx_Oracle.Connection(ODSServerUser, ODSServerPass, CreateDSN)
    ODSCursor = cx_Oracle.Cursor(ODSConnection)
    ODSCursor.execute(sql)
    results = ODSCursor.fetchall()

    for row in results:
        pidm = row[0]

    return pidm


def process():
    Students = {}

    if (len(sys.argv) == 1):
        print("Error: Please run script again using either 'PreCensus' or 'PostCensus' arguments ")
        exit()

    print("Begin:", datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S"))

    dataCode = getDataCodeData(INTERNATIONAL_ONLY)
    dataCode = getCSUIDandDOB(dataCode)
    shiPlan = getStudentHealthInsurancePlanData()

    comparisonResults = compareDataSets(dataCode, shiPlan)

    determinationResults, removedPlans = determinePlanUpdates(internationalDataFromPyraMED(comparisonResults))

    createWorkbook(determinationResults, studentData.fields, "testingResults-")
    writeImportFile(determinationResults, "planImport")

    createWorkbook(removedPlans, studentData.fields, "removedPlans-")
    print("Length of removed Plans: ", len(removedPlans))
    # print ("Length of removed Plans: ", len(removedPlans))
    print("End:", datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S"))


INTERNATIONAL_ONLY = True
INTERNATIONAL_ONLY = False

TERM = "202210"
PLAN_NAME = "2122StuIns"
GROUP_NUMBER = "0817013"

DOMESTIC_GROUP_NUMBER = "196519M001"
INTERNATIONAL_GROUP_NUMBER = "196519M003"

EFFECTIVE_DATE = datetime.datetime(2022, 1, 1)
INTERNATIONAL_EFFECTIVE_DATE = datetime.datetime(2022, 1, 1)
EXPIRATION_DATE = datetime.datetime(2022, 7, 31)
INTERNATIONAL_EXPIRATION_DATE = datetime.datetime(2022, 7, 31)
PRIOR_EFFECTIVE_DATE = datetime.datetime(2021, 8, 1)
INTERNATIONAL_PRIOR_EFFECTIVE_DATE = datetime.datetime(2021, 8, 1)
PRIOR_EXPIRATION_DATE = datetime.datetime(2021, 12, 31)

print ("test change")
# if(sys.argv[1] == "PreCensus"):
#	FILEPATH = 'N:\\NET\\Source Code\\Insurance\\Insurance Data\\202190\\PreCensus\\'
# else:
#	FILEPATH = 'N:\\NET\\Source Code\\Insurance\\Insurance Data\\202190\\PostCensus\\'
FILEPATH = 'N:\\NET\\Source Code\\Insurance\\Insurance Data\\202210\\PreCensus\\'

PyraMEDServerIP = "hhs215.hhs.colostate.edu"
PyraMEDServerPort = 1521
PyraMEDServerDBName = "p5prod"
# PyraMEDServerDBName = "p5test"
PyraMEDServerUser = "amullen_ro"
PyraMEDServerPass = "pyramed2010"

ODSServer = "dbodsprod.is.colostate.edu"
ODSServerPort = 1526
ODSServerDBName = "odsprod"
ODSServerUser = "mooman"
ODSServerPass = "wrucru9a"

BANProdServer = "dbbanprod.is.colostate.edu"
BANProdServerPort = 1526
BANProdServerDBName = "banprod"
BANProdServerUser = "HN_WEB"
BANProdServerPass = "cnfdtddcbt5fhjp8"

HARD_CODE_ADDITIONS = []  # ["831656050"]
HARD_CODE_EXPIRES = []

global extendExistingPlans
global newOnPreviouslyExpiredPlans

process()
