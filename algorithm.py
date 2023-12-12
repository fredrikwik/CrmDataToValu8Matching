from valu8search import getV8CompanybyRegNoCC


def step1OrgNoCheck(orgNo, CountryCode):
    """  """

    companyResultList = getV8CompanybyRegNoCC(orgNo, CountryCode)

    if companyResultList["ReturnedHits"] == 0:
        print("no match, step 2")

        print(orgNo)
        # Step 2.
        if orgNo != None:
            try:
                shortOrgNo = orgNo.split(CountryCode, 1)[1]
                print(shortOrgNo)

                ## Kolla om shortOrgNo funkar

                companyResultList2 = getV8CompanybyRegNoCC(shortOrgNo, CountryCode)




            except IndexError:
                print("ingen match")
                pass



        return None


    elif companyResultList["ReturnedHits"] == 1:
        print("Perfect match, return result")

        valu8ID = companyResultList["SearchResults"][0]["Valu8ID"]

        return valu8ID

    else:
        print("MULTI results, need to look into these in more detail")


step1OrgNoCheck("", "AT")