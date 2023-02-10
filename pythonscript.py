import openpyxl


wb = openpyxl.load_workbook('Leads - Sept.xlsx')


sheet = wb['Leads 120 days']



for row in sheet.rows:
    
    col10 = row[10].value
    
   
    
    if col10 in ["VA", "VA COMP", "VA Model C", "VA Model C(125-250)", "VA Model C(125k-250k)","VA Model C (250k+)", "VA Model C(250)", "VA Model C(250k+)", "VA Model C LLA", "VA Model C (125-250)"]:
        row[11].value = '490'
        row[12].value = "VA_CompLLA_DirectMail"

    elif col10 in ["VA AVEQU RAVEQ", "VA Model C (250+)", "VA Model C (250k+)"]:
        row[11].value = '490'
        row[12].value = "VA_CompLLA_DirectMail"

    elif col10 in ["VA TT", "VA TT(125-250)", "VA TT (250+)", "VA TT LLA", "VA TT (250k+)", "VA TT (250k+)"]:
        row[11].value = "489"
        row[12].value = "VA_TTLLA_DirectMail"
        # VA INTERNET 
    elif col10 in ["VA Int (250k+)", "VA TT(125-250)", "VA TT(250+)", "VA TT LLA", "VA TT (250+)", "VA TT (125-250)"]:
        row[11].value = "489"
        row[12].value = "VA_TTLLA_DirectMail"
    elif col10 in ["CONV", "Conventional", "CONV - 2018 Refi", "CONV - 2019 Refi", "CONV - MI Removal", "CONV 2018 Refi", "CONV 2019 Refi", "CONV Adverse Market Fee", "CONV AMRF NOFR", "Conv AMRF-NOFR"]:
        row[11].value = "492"
        row[12].value = "CONV_TTLLA_DirectMail"
   
    elif col10 in ["CONV HELOC", "CONV Internet", "CONV MI Removal", "CONV Credit"]:
        row[11].value = "492" 
        row[12].value = "CONV_TTLLA_DirectMail"

    elif col10 in ["CONV Model C", "Conv Model C(647+)", "Conv Model C Refi", "Conv Model C (647+)"]:
        row[11].value = "493"
        row[12].value = "CONV_ModelCLLA_DirectMail"

    elif col10 in ["", "Conv New Cashout", "Conv New Cashout LLA","NEW CASHOUT", "Conv Model C Refi","CONV FNMA-NOFE", "Conv New Cashout LLA", "Conv New CashoutLLA", "NEW CASH", "NEW CASH LLA", "Conv Property Cashout"]:
        row[11].value = "473"
        row[12].value = "CONV_NCO_DirectMail"

    elif col10 in ["Conv Property Cashout LLA", "Property Cashout", "Conv Property Cashout LLA", "PROPERTY CASHOUT LLA", "PROPERTY CASHOUT", "Conv PROPERTY CASHOUT"]:
        row[11].value = "482"
        row[12].value = "CONV_PCO_DirectMail"
        
    elif col10 in ["Conv Pur 5+ refi", "Conv Pur 5+ no refi", "Conv Purchase", "Conv Refi"]:
        row[11].value = "492"
        row[12].value = "CONV_TTLLA_DirectMail"

    elif col10 in ["CONV TT", "CONV TT - 2018 Refi", "CONV TT - ARM", "CONV TT - HELOC", "CONV TT - 2019 Refi"]:
        row[11].value = "492"
        row[12].value = "CONV_TTLLA_DirectMail"

    elif col10 in ["FHA"]:
        row[11].value = "487"
        row[12].value = "FHA_TTLLA_DirectMail"

    elif col10 in ["FHA COMP", "FHA COMP(125-250)", "FHA COMP LLA", "FHA Comp LLA", "FHA Comp", "FHA Comp (125-250)", "FHA Comp (125k-250k)"]:
        row[11].value = "486"
        row[12].value = "FHA_CompLLA_DirectMail"

    elif col10 in ["FHA Credit", "FHA Equity", "FHA Equity Access", "FHA Equity Notice", "FHA EQUITY RESERVES", "FHA Equity Reserves"]:
        row[11].value = "486"
        row[12].value = "FHA_CompLLA_DirectMail"

    elif col10 in ["FHA Int", "FHA Int(125k-250K)", "FHA Int LLA", "FHA Internet", "FHA Internet Comp", "FHA Int (125k-250k)"]:
        row[11].value = "485"
        row[12].value = "FHA_IntLLA_DirectMail"        

    elif col10 in ["FHA Model B", "FHA Model B LLA"]:
        row[11].value = "484"
        row[12].value = "FHA_ModBLLA_DirectMail"   

    elif col10 in ["FHA Model C", "FHA Model C LLA", "FHA Model C(125-250)", "FHA Model C(647+)", "FHA Model C LLA", "CONV  Model C - Co-Brwr", "FHA Model C (125-250)", "FHA Model C (647+)"]:
        row[11].value = "488"
        row[12].value = "FHA_ModelCLLA_DirectMail" 

    elif col10 in ["FHA TT", "FHA TT(125-250)", "FHA TT(125k-250k)", "FHA TT LLA", "FHA TT2020", "FHA TT (647+)", "FHA TT (125-250)", "FHA TT (125k-250k)"]:
        row[11].value = "487"
        row[12].value = "FHA_TTLLA_DirectMail" 

    elif col10 in ["No Category", "Trigger", "UWM Ops", "FHA TT LLA", "FHA TT2020", "FHA TT (647+)"]:
        row[11].value = "487"
        row[12].value = "FHA_TTLLA_DirectMail" 

    else:
        row[11].value = ""
        row[12].value = "" 



    




wb.save('Leads - Sept.xlsx')
