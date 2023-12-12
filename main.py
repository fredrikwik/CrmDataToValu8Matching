import pandas as pd
from valu8search import getV8CompanybyRegNoCC, getV8CompanybyTextCC, getV8CompanybyURLCC, getV8CompanybyNameURLCC

# Reading the Excel file
excel_file = 'Raw data SoverResults5.xlsx'
df = pd.read_excel(excel_file)

df['V8 Valu8ID'] = df["V8 Valu8ID"].astype(str)
df['V8 Company Name'] = df["V8 Company Name"].astype(str)
df['V8 Registration number'] = df["V8 Registration number"].astype(str)
df['V8 Website'] = df["V8 Website"].astype(str)

def step1OrgNoCheck(regNo, CountryCode):

    companyResultList = getV8CompanybyRegNoCC(regNo, CountryCode)

    if companyResultList["ReturnedHits"] == 1:
        print("Perfect match, return result")

        valu8ID = companyResultList["SearchResults"][0]["Valu8Id"]
        OrgNo = companyResultList["SearchResults"][0]["OrgNo"]
        CompanyName = companyResultList["SearchResults"][0]["CompanyName"]
        Website = companyResultList["SearchResults"][0]["Website"]


        output = {
            "valu8ID": valu8ID,
            "orgNo": OrgNo,
            "companyName": CompanyName,
            "website": Website,
        }

        return output


    elif companyResultList["ReturnedHits"] == 0:
        print("no match, going to step 2")

        print(regNo)
        # Step 2.
        if regNo != None:
            try:
                shortOrgNo = regNo.split(CountryCode, 1)[1]
                print(shortOrgNo)

                ## Kolla om shortOrgNo funkar

                companyResultList = getV8CompanybyRegNoCC(shortOrgNo, CountryCode)

                if companyResultList["ReturnedHits"] == 1:
                    print("Perfect match, return result")

                    valu8ID = companyResultList["SearchResults"][0]["Valu8Id"]
                    OrgNo = companyResultList["SearchResults"][0]["OrgNo"]
                    CompanyName = companyResultList["SearchResults"][0]["CompanyName"]
                    Website = companyResultList["SearchResults"][0]["Website"]

                    output = {
                        "valu8ID": valu8ID,
                        "orgNo": OrgNo,
                        "companyName": CompanyName,
                        "website": Website,
                    }

                    return output

                elif companyResultList["ReturnedHits"] > 1:
                    print("Multi results level 2")
                    return "Multi results level 2"

                else:
                    output = {
                        "valu8ID": None,
                        "orgNo": None,
                        "companyName": None,
                        "website": None,
                    }

                    return output


            except IndexError:
                print("ERROR, ingen match, ERROR")
                output = {
                    "valu8ID": None,
                    "orgNo": None,
                    "companyName": None,
                    "website": None
                }

                return output




        else:
            output = {
                "valu8ID": None,
                "orgNo": None,
                "companyName": None,
                "website": None
            }

            return output




    else:
        print("MULTI results, need to look into these in more detail")

        output = {
            "valu8ID": None,
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

def step2CompanyNameCheck(companyName, CountryCode):

    companyResultList = getV8CompanybyTextCC(companyName, CountryCode)

    if companyResultList["ReturnedHits"] == 1:
        print("Perfect match, return result")

        valu8ID = companyResultList["SearchResults"][0]["Valu8Id"]
        OrgNo = companyResultList["SearchResults"][0]["OrgNo"]
        CompanyName = companyResultList["SearchResults"][0]["CompanyName"]
        Website = companyResultList["SearchResults"][0]["Website"]


        output = {
            "valu8ID": valu8ID,
            "orgNo": OrgNo,
            "companyName": CompanyName,
            "website": Website,
        }

        return output


    elif companyResultList["ReturnedHits"] == 0:
        print("no match")
        output = {
            "valu8ID": None,
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

    elif companyResultList["ReturnedHits"] > 1:
        print("Multi match")

        output = {
            "valu8ID": "Multi name match",
            "orgNo": "Multi name match",
            "companyName": "Multi name match",
            "website": "Multi name match"
        }

        return output

    else:
        print("Sometihg starange here")

        output = {
            "valu8ID": "Something strange",
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

def step3CompanyURLCheck(url, CountryCode):

    companyResultList = getV8CompanybyURLCC(url, CountryCode)


    if companyResultList["ReturnedHits"] == 1:
        print("Perfect match, return result")

        valu8ID = companyResultList["SearchResults"][0]["Valu8Id"]
        OrgNo = companyResultList["SearchResults"][0]["OrgNo"]
        CompanyName = companyResultList["SearchResults"][0]["CompanyName"]
        Website = companyResultList["SearchResults"][0]["Website"]


        output = {
            "valu8ID": valu8ID,
            "orgNo": OrgNo,
            "companyName": CompanyName,
            "website": Website,
        }
        print(output)

        return output


    elif companyResultList["ReturnedHits"] == 0:
        print("no match")
        output = {
            "valu8ID": None,
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

    elif companyResultList["ReturnedHits"] > 1:
        print("Multi match")

        output = {
            "valu8ID": "Multi name match",
            "orgNo": "Multi name match",
            "companyName": "Multi name match",
            "website": "Multi name match",
        }

        return output

    else:
        print("Sometihg starange here")

        output = {
            "valu8ID": "Something strange",
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

def step4CompanyNameURLCheck(companyName, website, CountryCode):

    companyResultList = getV8CompanybyNameURLCC(companyName, website, CountryCode)


    if companyResultList["ReturnedHits"] == 1:
        print("Perfect match, return result")

        valu8ID = companyResultList["SearchResults"][0]["Valu8Id"]
        OrgNo = companyResultList["SearchResults"][0]["OrgNo"]
        CompanyName = companyResultList["SearchResults"][0]["CompanyName"]
        Website = companyResultList["SearchResults"][0]["Website"]


        output = {
            "valu8ID": valu8ID,
            "orgNo": OrgNo,
            "companyName": CompanyName,
            "website": Website,
        }
        print(output)

        return output


    elif companyResultList["ReturnedHits"] == 0:
        print("no match")
        output = {
            "valu8ID": None,
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

    elif companyResultList["ReturnedHits"] > 1:
        print("Multi match")

        output = {
            "valu8ID": "Multi name match",
            "orgNo": "Multi name match",
            "companyName": "Multi name match",
            "website": "Multi name match",
        }

        return output

    else:
        print("Sometihg starange here")

        output = {
            "valu8ID": "Something strange",
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

def step5CompanyNameURLCheck(companyName, website, CountryCode):

    companyResultList = getV8CompanybyNameURLCC(companyName, website, CountryCode)


    if companyResultList["ReturnedHits"] > 0:
        print("match, return result")

        valu8ID = companyResultList["SearchResults"][0]["Valu8Id"]
        OrgNo = companyResultList["SearchResults"][0]["OrgNo"]
        CompanyName = companyResultList["SearchResults"][0]["CompanyName"]
        Website = companyResultList["SearchResults"][0]["Website"]


        output = {
            "valu8ID": valu8ID,
            "orgNo": OrgNo,
            "companyName": CompanyName,
            "website": Website,
        }
        print(output)

        return output


    elif companyResultList["ReturnedHits"] == 0:
        print("no match")
        output = {
            "valu8ID": None,
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

    else:
        print("Sometihg starange here")

        output = {
            "valu8ID": "Something strange",
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output


def step6CompanyNameCheck(companyName, CountryCode):

    companyResultList = getV8CompanybyTextCC(companyName, CountryCode)


    if companyResultList["ReturnedHits"] > 0:
        print("match, return result")

        valu8ID = companyResultList["SearchResults"][0]["Valu8Id"]
        OrgNo = companyResultList["SearchResults"][0]["OrgNo"]
        CompanyName = companyResultList["SearchResults"][0]["CompanyName"]
        Website = companyResultList["SearchResults"][0]["Website"]


        output = {
            "valu8ID": valu8ID,
            "orgNo": OrgNo,
            "companyName": CompanyName,
            "website": Website,
        }
        print(output)

        return output


    elif companyResultList["ReturnedHits"] == 0:
        print("no match")
        output = {
            "valu8ID": None,
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

    else:
        print("Sometihg starange here")

        output = {
            "valu8ID": "Something strange",
            "orgNo": None,
            "companyName": None,
            "website": None
        }

        return output

def main():


    # Looping through each row
    for index, row in df.iterrows():
        print(index)
        # 'row' is a Series representing the row data
        # You can access column data by the column name, e.g., row['column_namåäe']
        regNo = row["Company Reg"]
        countryCode = row["CountryCode"]


        if pd.isna(regNo) == False and pd.isna(countryCode) == False:

            output = step1OrgNoCheck(regNo, countryCode)
            df.at[index, "V8 Valu8ID"] = output["valu8ID"]
            df.at[index, "V8 Company Name"] = output["companyName"]
            df.at[index, "V8 Registration number"] = output["orgNo"]
            df.at[index, "V8 Website"] = output["website"]


        else:
            print("No match")
            df.at[index, "V8 Valu8ID"] = None
            df.at[index, "V8 Company Name"] = None
            df.at[index, "V8 Registration number"] = None
            df.at[index, "V8 Website"] = None



def main2():

    # Looping through each row again
    for index, row in df.iterrows():
        print(index)
        # 'row' is a Series representing the row data
        # You can access column data by the column name, e.g., row['column_namåäe']
        companyName = row["Name (ID)"]
        countryCode = row["CountryCode"]


        if pd.isna(companyName) == False and pd.isna(countryCode) == False:

            output = step2CompanyNameCheck(companyName, countryCode)
            df.at[index, "V8 Valu8ID"] = output["valu8ID"]
            df.at[index, "V8 Company Name"] = output["companyName"]
            df.at[index, "V8 Registration number"] = output["orgNo"]
            df.at[index, "V8 Website"] = output["website"]



        else:
            print("No match")
            df.at[index, "V8 Valu8ID"] = None
            df.at[index, "V8 Company Name"] = None
            df.at[index, "V8 Registration number"] = None
            df.at[index, "V8 Website"] = None




def main3():
    urls = [
        "www.123sonography.com",
        "http://www.200oksolutions.com",
        "https://2021.ai/",
        "https://www.25delta.es",
        "https://2bm.com/",
        "https://www.4net.ch/",
        "www.4xxi.com",
    ]

    # Looping through each row again
    for index, row in df.iterrows():
        print(index)
        # 'row' is a Series representing the row data
        # You can access column data by the column name, e.g., row['column_namåäe']
        URL = row["Website"]
        countryCode = row["CountryCode"]


        if pd.isna(URL) == False and pd.isna(countryCode) == False:

            output = step3CompanyURLCheck(URL, countryCode)
            df.at[index, "V8 Valu8ID"] = output["valu8ID"]
            df.at[index, "V8 Company Name"] = output["companyName"]
            df.at[index, "V8 Registration number"] = output["orgNo"]
            df.at[index, "V8 Website"] = output["website"]


        else:
            print("No match")
            df.at[index, "V8 Valu8ID"] = None
            df.at[index, "V8 Company Name"] = None
            df.at[index, "V8 Registration number"] = None
            df.at[index, "V8 Website"] = None



def main4():

    # Looping through each row again
    for index, row in df.iterrows():
        print(index)
        # 'row' is a Series representing the row data
        # You can access column data by the column name, e.g., row['column_namåäe']
        companyName = row["Name (ID)"]
        URL = row["Website"]
        countryCode = row["CountryCode"]


        if pd.isna(URL) == False and pd.isna(countryCode) == False:

            output = step4CompanyNameURLCheck(companyName, URL, countryCode)
            df.at[index, "V8 Valu8ID"] = output["valu8ID"]
            df.at[index, "V8 Company Name"] = output["companyName"]
            df.at[index, "V8 Registration number"] = output["orgNo"]
            df.at[index, "V8 Website"] = output["website"]


        else:
            print("No match")
            df.at[index, "V8 Valu8ID"] = None
            df.at[index, "V8 Company Name"] = None
            df.at[index, "V8 Registration number"] = None
            df.at[index, "V8 Website"] = None

def main5():

    # Looping through each row again
    for index, row in df.iterrows():
        print(index)
        # 'row' is a Series representing the row data
        # You can access column data by the column name, e.g., row['column_namåäe']
        companyName = row["Name (ID)"]
        URL = row["Website"]
        countryCode = row["CountryCode"]


        if pd.isna(URL) == False and pd.isna(countryCode) == False:

            output = step5CompanyNameURLCheck(companyName, URL, countryCode)
            df.at[index, "V8 Valu8ID"] = output["valu8ID"]
            df.at[index, "V8 Company Name"] = output["companyName"]
            df.at[index, "V8 Registration number"] = output["orgNo"]
            df.at[index, "V8 Website"] = output["website"]


        else:
            print("No match")
            df.at[index, "V8 Valu8ID"] = None
            df.at[index, "V8 Company Name"] = None
            df.at[index, "V8 Registration number"] = None
            df.at[index, "V8 Website"] = None

def main6():

    # Looping through each row again
    for index, row in df.iterrows():
        print(index)
        # 'row' is a Series representing the row data
        # You can access column data by the column name, e.g., row['column_namåäe']
        companyName = row["Name (ID)"]
        URL = row["Website"]
        countryCode = row["CountryCode"]


        if pd.isna(URL) == False and pd.isna(countryCode) == False:

            output = step6CompanyNameCheck(companyName, countryCode)
            df.at[index, "V8 Valu8ID"] = output["valu8ID"]
            df.at[index, "V8 Company Name"] = output["companyName"]
            df.at[index, "V8 Registration number"] = output["orgNo"]
            df.at[index, "V8 Website"] = output["website"]


        else:
            print("No match")
            df.at[index, "V8 Valu8ID"] = None
            df.at[index, "V8 Company Name"] = None
            df.at[index, "V8 Registration number"] = None
            df.at[index, "V8 Website"] = None



if __name__ == "__main__":
    # main()
    # main2()
    # main3()
    # main4()
    # main5()
    main6()

    df.to_excel('Raw data SoverResults6.xlsx')
