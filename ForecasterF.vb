Sub ForecasterFwithsetRevised()

'Revised (not tested but with improvements to the original) Macro for updating Forecast file
'This defines the dimensions of the data for variables that will be defined
Dim forecastrng, incomerng, forallorng, inallorng As String
Dim myOrder,myOrder2, myOrder3 As Variant
Dim Income, Forecast As Workbook
Dim ws As Worksheet
Dim i As Integer
Dim x As Long
  
'Set workbooks to open files. File name will be editing each run. Make sure these files are both open when you run the macro
Set Income = Workbooks("2020-03 CM Detail IS (T-13 Actual) for final")
Set Forecast = Workbooks("3+9 Forecast (4.14.20) V2")

'Set Ranges we taking and applying from the income statement and forecast file. Will most likely need to edit forecastrng and forallorng column letter for each new month
forecastrng = "AM7:AM243"
incomerng = "N7:N243"
forallorng = "AM8:AM32"
inallorng = "N198:N222"

'This is an array of the desired order needed in the income Statement. It removes unwanted tabs for the entry into the forecast file
myOrder = Array("11XX-F", "Watch Stores-F", "Same-stores-F", "Retail Management-F", "Donation Acq-F", "Outlets and Recycling-F", "Logistics-F", "ADCs-F", "Property Services-F", "Misc Vehicles-F", "Buildings-F", "Private Temps-F", "State Temps-F", "Private Contract-F", "State Contracts-F", "Management Contracts-F", "Federal Contracts-F", "CS Other-F", "7580-F", "Business Solutions Misc-F", "GCTA Admin-F", "CSN-F", "GCTA Tuition Class-F", "GCTA Other-F", "WFA Management-F", "Match-F", "Flat Fee-F", "Pay For Performance-F", "Cost Reimbursement-F", "WFA GW Programs-F", "WFA Closed-F", "6160-F", "6061-F", "6219-F", "Accounting-F", "6060-F", "6064-F", "6066-F", "6070-F", "IT-F", "Asset Protection-F", "FP&A-F", "Endowment Funds-F", "7198-F", "Summary EXCEL-F", "Public Relations-F", "Lobbying-F", "Development-F", "6230-F", "EMS-F", "6200-F", "Legacy Accounts-F", "GWC Contingency-F", "0000-F")

'Array order that defines do not edit or non entry worksheets in Forecast file
myOrder2 = Array("Tieout", "World View", "Summary", "Mark's World", "Retail ", "All Stores", "Post-Retail & Ops", "GSG", "Comm Services", "Blue Solutions", "WFA Total", "GCTA", "Workforce Advancement", "One-Time", "Paula's World", "9998", "Traci's World", "Steve's World", "Kenny's World", "Allocated Expenses", "Capital", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference", "Capital ", "Comparison", "64XX", "Rents", "Instructions", "Initiative")

'Reorders Forecast file to orginal order
myOrder3 = Array("World View", "Summary", "Mark's World", "Retail ", "ECommerce", "All Stores", "Watch Stores", "Same Stores", "Retail Admin", "Post-Retail & Ops", "Don Acq", "Outlets and Recycling", "Logistics", "ADCs", "Property Svcs", "Misc Vehicles", "Buildings", "GSG", "Private Temps", "State Temps", "Comm Services", "Private Contracts", "State Contracts", "Blue Solutions", "Mgmt Contracts", "Fed Contracts", "CS Other", "Bus Dev", "Bus Sol Misc", "WFA Total", "GCTA", "GCTA Admin", "CSN", "GCTA Tuition Class", "GCTA Other", "Workforce Advancement", "WFA Management", "Match", "Flat Fee", "Pay For Performance", "Cost Reimbursement", "WFA GW Programs", "One-Time", "Other", "GCC Building", "Paula's World", "Training", "EE", "Accounting", "HR", "Safety", "Employee Benefits", "Mgmt Svcs", "IT", "AP", "FP&A", "Endowment", "WFS Lease", "9998", "Traci's World", "EXCEL", "Lobbying", "Steve's World", "PR", "Development", "CC", "Kenny's World", "EMS", "6200", "Legacy", "Contingency", _
"Eliminations", "Allocated Expenses", "Capital ", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")

'This speeds up code and reduces screen flickering
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
    
'This is the code that puts array in order from left until it ends. Unused wksheets will be to the right of last wksht in array
    On Error Resume Next
        For x = UBound(myOrder) To LBound(myOrder) Step -1
        Income.Worksheets(myOrder(x)).Move Before:=Income.Worksheets(1)
        Next x
    On Error GoTo 0
    
'This is the code that puts array in order from left until it ends. Unused wksheets will be to the right of last wksht in array
    On Error Resume Next
        For x = UBound(myOrder2) To LBound(myOrder2) Step -1
        Forecast.Worksheets(myOrder2(x)).Move After:=Forecast.Worksheets(Worksheets.Count)
        Next x
    On Error GoTo 0

'Loop. takes desired range in income statement and puts it into the forecast file.
    For i = 1 To 54

        Forecast.Sheets(i).Range(forecastrng).Value = Income.Sheets(i).Range(incomerng).Value

    Next i
 
'Allocated Expenses is different than the other worksheets; therefore, it gets a different calculation
Forecast.Sheets("Allocated Expenses").Range(forallorng).Value = Income.Sheets("SUMMARY-F").Range(inallorng).Value 

'This is the code that puts array in order from left until it ends. Unused wksheets will be to the right of last wksht in array
    On Error Resume Next
        For x = UBound(myOrder3) To LBound(myOrder2) Step -1
        Forecast.Worksheets(myOrder3(x)).Move Before:=Forecast.Worksheets(1)
        Next x
    On Error GoTo 0
    
'Enables turned off Excel functions
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
    
End Sub