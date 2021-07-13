import shutil
from openpyxl.worksheet.datavalidation import DataValidation


class Utils:

    @staticmethod
    def duplicateTemplateLTGC(tempLTGC_Path, out, outputFilename):
        companyDir = out + "/"

        shutil.copy(tempLTGC_Path,
                    companyDir + outputFilename + ".xlsx")

        return companyDir + outputFilename + ".xlsx"

    @staticmethod
    def addingDataValidation(currentSheet, numrows):
        # print("start Init Data validation")
        # create data validation
        Category_data_val = DataValidation(type="list", formula1="=LOVCategories")
        currentSheet.add_data_validation(Category_data_val)

        CategoryID_data_val = DataValidation(type="list", formula1="=LOVCategoryID")
        currentSheet.add_data_validation(CategoryID_data_val)

        Suffix_data_val = DataValidation(type="list", formula1="=LOVSuffix")
        currentSheet.add_data_validation(Suffix_data_val)

        C_residence_region_data_val = DataValidation(type="list", formula1="=Region")
        currentSheet.add_data_validation(C_residence_region_data_val)

        C_residence_province_data_val = DataValidation(type="list", formula1="=INDIRECT(L3)")
        currentSheet.add_data_validation(C_residence_province_data_val)

        C_residence_municipality_data_val = DataValidation(type="list", formula1="=INDIRECT(M3)")
        currentSheet.add_data_validation(C_residence_municipality_data_val)

        C_residence_Barangay_data_val = DataValidation(type="list", formula1="=INDIRECT(N3)")
        currentSheet.add_data_validation(C_residence_Barangay_data_val)

        sex_data_val = DataValidation(type="list", formula1="=LOVSex")
        currentSheet.add_data_validation(sex_data_val)

        civilStatus_data_val = DataValidation(type="list", formula1="=LOVCivilStatus")
        currentSheet.add_data_validation(civilStatus_data_val)

        employmentStatus_data_val = DataValidation(type="list", formula1="=LOVEmploymentStatus")
        currentSheet.add_data_validation(employmentStatus_data_val)

        Directly_in_interaction_with_COVID_patient_data_val = DataValidation(type="list", formula1="=LOVDirectCovid")
        currentSheet.add_data_validation(Directly_in_interaction_with_COVID_patient_data_val)

        Profession_data_val = DataValidation(type="list", formula1="=LOVProfession")
        currentSheet.add_data_validation(Profession_data_val)

        ICC_of_Employer_data_val = DataValidation(type="list", formula1="=LOVProvinceHUCICCofEmployer")
        currentSheet.add_data_validation(ICC_of_Employer_data_val)

        Pregnancy_status_data_val = DataValidation(type="list", formula1="=LOVPregnancyStatus")
        currentSheet.add_data_validation(Pregnancy_status_data_val)

        YesNo_data_val = DataValidation(type="list", formula1="=LOVYesNo")
        currentSheet.add_data_validation(YesNo_data_val)

        With_Comorbidity_data_val = DataValidation(type="list", formula1="=LOVYesNone")
        currentSheet.add_data_validation(With_Comorbidity_data_val)

        Classification_of_COVID_19_data_val = DataValidation(type="list", formula1="=LOVCovidClass")
        currentSheet.add_data_validation(Classification_of_COVID_19_data_val)

        Willing_to_be_Vaccinated_data_val = DataValidation(type="list", formula1="=LOVConsent")
        currentSheet.add_data_validation(Willing_to_be_Vaccinated_data_val)

        Working_from_Home_data_val = DataValidation(type="list", formula1="=LOVWFH")
        currentSheet.add_data_validation(Working_from_Home_data_val)

        A1_Health_Worker_data_val = DataValidation(type="list", formula1="=A1LOV")
        currentSheet.add_data_validation(A1_Health_Worker_data_val)

        A2_Senior_data_val = DataValidation(type="list", formula1="=A2LOV")
        currentSheet.add_data_validation(A2_Senior_data_val)

        A3_With_Co_morbidity_data_val = DataValidation(type="list", formula1="=A3LOV")
        currentSheet.add_data_validation(A3_With_Co_morbidity_data_val)

        Risk_of_Exposure_data_val = DataValidation(type="list", formula1="=RiskOfExposure")
        currentSheet.add_data_validation(Risk_of_Exposure_data_val)

        Business_Continuity_data_val = DataValidation(type="list", formula1="=BusinessContinuity")
        currentSheet.add_data_validation(Business_Continuity_data_val)

        Type_of_Employees_data_val = DataValidation(type="list", formula1="=TypeOfEmployees")
        currentSheet.add_data_validation(Type_of_Employees_data_val)

        Public_Image_Impact_data_val = DataValidation(type="list", formula1="=PublicImage")
        currentSheet.add_data_validation(Public_Image_Impact_data_val)

        Age_Risk_Factor_data_val = DataValidation(type="list", formula1="=AgeRiskFactor")
        currentSheet.add_data_validation(Age_Risk_Factor_data_val)

        Confirmed_Vaccination_Site_data_val = DataValidation(type="list", formula1="=VaccinationSites")
        currentSheet.add_data_validation(Confirmed_Vaccination_Site_data_val)
        # print("Done Init Data validation")
        #
        # print("Start assigning Data validation")
        row = numrows+3
        Confirmed_Vaccination_Site_data_val.add("BV3:BV"+str(row))
        Category_data_val.add("A3:A" + str(row))
        CategoryID_data_val.add("B3:B" + str(row))
        Suffix_data_val.add("I3:I" + str(row))
        C_residence_region_data_val.add("L3:L" + str(row))
        C_residence_province_data_val.add("M3:M" + str(row))
        C_residence_municipality_data_val.add("N3:N" + str(row))
        C_residence_Barangay_data_val.add("O3:O" + str(row))
        sex_data_val.add("P3:P" + str(row))
        civilStatus_data_val.add("R3:R" + str(row))
        employmentStatus_data_val.add("S3:S" + str(row))
        Directly_in_interaction_with_COVID_patient_data_val.add("T3:T" + str(row))
        Profession_data_val.add("U3:U" + str(row))
        ICC_of_Employer_data_val.add("W3:W" + str(row))
        Pregnancy_status_data_val.add("Z3:Z" + str(row))
        YesNo_data_val.add("AA3:AA" + str(row))
        YesNo_data_val.add("AB3:AB" + str(row))
        YesNo_data_val.add("AC3:AC" + str(row))
        YesNo_data_val.add("AD3:AD" + str(row))
        YesNo_data_val.add("AE3:AE" + str(row))
        YesNo_data_val.add("AF3:AF" + str(row))
        YesNo_data_val.add("AG3:AG" + str(row))
        With_Comorbidity_data_val.add(("AH3:AH" + str(row)))
        YesNo_data_val.add("AI3:AI" + str(row))
        YesNo_data_val.add("AJ3:AJ" + str(row))
        YesNo_data_val.add("AK3:AK" + str(row))
        YesNo_data_val.add("AL3:AL" + str(row))
        YesNo_data_val.add("AM3:AM" + str(row))
        YesNo_data_val.add("AN3:AN" + str(row))
        YesNo_data_val.add("AO3:AO" + str(row))
        YesNo_data_val.add("AP3:AP" + str(row))
        YesNo_data_val.add("AQ3:AQ" + str(row))
        Classification_of_COVID_19_data_val.add("AS3:AS" + str(row))
        Willing_to_be_Vaccinated_data_val.add("AT3:AT" + str(row))
        A1_Health_Worker_data_val.add("BF3:BF" + str(row))
        A2_Senior_data_val.add("BG3:BG" + str(row))
        A3_With_Co_morbidity_data_val.add("BH3:BH" + str(row))
        Risk_of_Exposure_data_val.add("BI3:BI" + str(row))
        Business_Continuity_data_val.add("BJ3:BJ" + str(row))
        Type_of_Employees_data_val.add("BK3:BK" + str(row))
        Public_Image_Impact_data_val.add("BL3:BL" + str(row))
        Age_Risk_Factor_data_val.add(("BM3:BM" + str(row)))

        # print("Done assigning Data validation")
        pass
