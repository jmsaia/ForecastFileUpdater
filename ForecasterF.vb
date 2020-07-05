Sub ForecasterFwithset()

'Orginal file without new edits (The untested one is in the brach with edits should be faster and better. I've improved since this code)
'Tested and working Macro for updating Forecast file
    'This defines the dimensions of the data for variables that will be defined
    'This is the most recent version of the Forecast Macro that has been tested
    Dim ws As Worksheet
    Dim Income As Workbook
    Dim Forecast As Workbook
    Dim i As Integer
    Dim x As Long
    Dim myOrder As Variant
    Dim myOrder2 As Variant
    Dim myOrder3 As Variant
    
    
    'Change set workbooks to match files open
    Set Income = Workbooks("2020-03 CM Detail IS (T-13 Actual) for final")
    Set Forecast = Workbooks("3+9 Forecast (4.14.20) V2")
    
    'This speeds up code and reduces screen flickering
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'This activates the open workbook (All workbooks we are using should be open)
    Income.Activate
    
    'This is an array of the desired order needed in the CM IS. It removes unwanted tabs for the entry into the forecast file
    myOrder = Array("11XX-F", "Watch Stores-F", "Same-stores-F", "Retail Management-F", "Donation Acq-F", "Outlets and Recycling-F", "Logistics-F", "ADCs-F", "Property Services-F", "Misc Vehicles-F", "Buildings-F", "Private Temps-F", "State Temps-F", "Private Contract-F", "State Contracts-F", "Management Contracts-F", "Federal Contracts-F", "CS Other-F", "7580-F", "Business Solutions Misc-F", "GCTA Admin-F", "CSN-F", "GCTA Tuition Class-F", "GCTA Other-F", "WFA Management-F", "Match-F", "Flat Fee-F", "Pay For Performance-F", "Cost Reimbursement-F", "WFA GW Programs-F", "WFA Closed-F", "6160-F", "6061-F", "6219-F", "Accounting-F", "6060-F", "6064-F", "6066-F", "6070-F", "IT-F", "Asset Protection-F", "FP&A-F", "Endowment Funds-F", "7198-F", "Summary EXCEL-F", "Public Relations-F", "Lobbying-F", "Development-F", "6230-F", "EMS-F", "6200-F", "Legacy Accounts-F", "GWC Contingency-F", "0000-F")
    
    'This is the code that puts array in order from left until it ends. Unused wksheets will be to the right of last wksht in array
    On Error Resume Next
    For x = UBound(myOrder) To LBound(myOrder) Step -1
    Worksheets(myOrder(x)).Move Before:=Worksheets(1)
    Next x
    On Error GoTo 0

    'This defines the end of the wkshts we want to keep aka end of the array
    lStart = 54
   
    'This For Each loop finds the index number and if it is larger than lStart's index # it deletes it
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Index >= Sheets(lStart).Index Then ws.Visible = False
    
    Next
    
    'Brings worksheet to the beginning for when we excute copy paste loop
    Worksheets("11XX-F").Activate
    
    'Activates Forecast file
    Forecast.Activate
    
    'Array order that defines do not edit or non entry worksheets
    myOrder2 = Array("Tieout", "World View", "Summary", "Mark's World", "Retail ", "All Stores", "Post-Retail & Ops", "GSG", "Comm Services", "Blue Solutions", "WFA Total", "GCTA", "Workforce Advancement", "One-Time", "Paula's World", "9998", "Traci's World", "Steve's World", "Kenny's World", "Allocated Expenses", "Capital", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference", "Capital ", "Comparison", "64XX", "Rents", "Instructions", "Initiative")
    
    'This is the code that puts array in order from left until it ends. Unused wksheets will be to the right of last wksht in array
    On Error Resume Next
    For x = UBound(myOrder2) To LBound(myOrder2) Step -1
    Worksheets(myOrder2(x)).Move After:=Worksheets(Worksheets.Count)
    Worksheets(myOrder2(x)).Visible = True
    Next x
    On Error GoTo 0
    
    'Brings worksheet to the beginning for when we excute copy paste loop
    Worksheets("ECommerce").Activate
    
'Copy/Paste loop. Performs copy paste for a set number of rows. They will be equal. You will end up on the tab past elimination on the Forecast file
For i = 1 To 54
    If Worksheets(i).Visible = True Then
    Income.Activate
    Range("N7:N243").Select
    Selection.Copy
    Forecast.Activate
    Range("AM7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Next.Activate
    Income.Activate
    ActiveSheet.Next.Activate
    
    End If

Next i
 
 'Allocated Expenses is different;therefore, it gets a different calculation
    Income.Sheets("SUMMARY-F").Range("N198:N222").Copy
    Forecast.Sheets("Allocated Expenses").Range("AM8").PasteSpecial Paste:=xlPasteValues

    
    'Activates Forecast file
    Forecast.Activate
    
    'Reorders Forecast file to orginal order
    myOrder3 = Array("World View", "Summary", "Mark's World", "Retail ", "ECommerce", "All Stores", "Watch Stores", "Same Stores", "Retail Admin", "Post-Retail & Ops", "Don Acq", "Outlets and Recycling", "Logistics", "ADCs", "Property Svcs", "Misc Vehicles", "Buildings", "GSG", "Private Temps", "State Temps", "Comm Services", "Private Contracts", "State Contracts", "Blue Solutions", "Mgmt Contracts", "Fed Contracts", "CS Other", "Bus Dev", "Bus Sol Misc", "WFA Total", "GCTA", "GCTA Admin", "CSN", "GCTA Tuition Class", "GCTA Other", "Workforce Advancement", "WFA Management", "Match", "Flat Fee", "Pay For Performance", "Cost Reimbursement", "WFA GW Programs", "One-Time", "Other", "GCC Building", "Paula's World", "Training", "EE", "Accounting", "HR", "Safety", "Employee Benefits", "Mgmt Svcs", "IT", "AP", "FP&A", "Endowment", "WFS Lease", "9998", "Traci's World", "EXCEL", "Lobbying", "Steve's World", "PR", "Development", "CC", "Kenny's World", "EMS", "6200", "Legacy", "Contingency", _
    "Eliminations", "Allocated Expenses", "Capital ", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")
    
    'This is the code that puts array in order from left until it ends. Unused wksheets will be to the right of last wksht in array
    On Error Resume Next
    For x = UBound(myOrder3) To LBound(myOrder2) Step -1
    Worksheets(myOrder3(x)).Move Before:=Worksheets(1)
    Worksheets(myOrder3(x)).Visible = True
    Next x
    On Error GoTo 0
    
    'Enables turned off Excel functions
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub