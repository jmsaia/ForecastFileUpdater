Sub PrelimForecastSheetDivider()

'This Macro is used to divide the prelim forecast file for forecast entries by dept. owners for review. One should hide cross-dept tabs if the owner doesn't forecast those depts
'This macro saves the new workbooks out to a specificied location. You need to change them each month

Dim x As Long
Dim myOrder, myOrder2, myOrder3, myOrder4 As Variant
Dim wb As Workbook

'Change to match most recent forcast file

Set wb = Workbooks("3+9 Forecast (4.8.20)")

'Setting Arrays for particular groups. myOrder2 is Comm Services, myOrder3 is IT, myOrder4 is PR

myOrder = Array("Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")
myOrder2 = Array("Telephone 738", "Printing 756", "Advertising 758", "Internet 786", "Reference", "State Temps", "Private Temps")
myOrder3 = Array("Printing 756", "Building Repairs 742", "Advertising 758", "Reference")
myOrder4 = Array("Telephone 738", "Building Repairs 742", "Internet 786", "Reference")

'Speeds up macros. Must turn back on at the end of the macro

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

'You will need to change file names and paths if they change/move
    
'Commercial Services. State and Private Temps added bc they are linked. If not included and hidden, one would have Value/ref errors in Payroll Accural forecast
wb.Sheets(Array("Comm Services", "Private Contracts", "State Contracts", "Blue Solutions", "Mgmt Contracts", "Fed Contracts", "CS Other", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference", "State Temps", "Private Temps")).Copy

    For x = UBound(myOrder2) To LBound(myOrder2) Step -1
        Worksheets(myOrder2(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder2(x)).Visible = False
    Next x
    On Error GoTo 0
    
Workbooks("Book1").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Mark\Commercial Services\2020\Commercial Services 3+9.xlsx"
    
'GSG
wb.Sheets(Array("GSG", "Private Temps", "State Temps", "Bus Dev", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0

Workbooks("Book2").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Mark\GSG\2020\GSG 3+9.xlsx"

'Post Retail
wb.Sheets(Array("Post-Retail & Ops", "Don Acq", "Outlets and Recycling", "Logistics", "ADCs", "Property Svcs", "Misc Vehicles", "Buildings", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0
    
Workbooks("Book3").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Mark\Post Retail & Ops\2020\Post-Retail 3+9.xlsx"
    
'Retail
wb.Sheets(Array("Retail ", "ECommerce", "All Stores", "Watch Stores", "Same Stores", "Retail Admin", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0
    
Workbooks("Book4").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Mark\Retail\2020\Retail 3+9.xlsx"
    
'WFA GCTA
wb.Sheets(Array("WFA Total", "GCTA", "GCTA Admin", "CSN", "GCTA Tuition Class", "GCTA Other", "Workforce Advancement", "WFA Management", "Match", "Flat Fee", "Pay For Performance", "Cost Reimbursement", "WFA GW Programs", "One-Time", "Other", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0
    
Workbooks("Book5").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Mark\WFA GCTA\2020\WFA GCTA 3+9.xlsx"
    
'Dodie
wb.Sheets(Array("Mgmt Svcs", "GCC Building", "Accounting", "Employee Benefits", "FP&A", "Endowment", "WFS Lease", "9998", "Legacy", "Contingency", "Eliminations", "Allocated Expenses", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0
    
Workbooks("Book6").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Paula\Dodie\2020\Dodie 3+9.xlsx"
    
'HR Safety
wb.Sheets(Array("HR", "Safety", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0

Workbooks("Book7").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Paula\HR\2020\HR Safety 3+9.xlsx"

'IT
wb.Sheets(Array("IT", "AP", "Telephone 738", "Internet 786", "Building Repairs 742", "Printing 756", "Advertising 758", "Reference")).Copy

    For x = UBound(myOrder3) To LBound(myOrder3) Step -1
        Worksheets(myOrder3(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder3(x)).Visible = False
    Next x
    On Error GoTo 0
    
Workbooks("Book8").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Paula\IT\2020\IT 3+9.xlsx"

'Training
wb.Sheets(Array("EE", "Training", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0

Workbooks("Book9").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Paula\Training\2020\Training 3+9.xlsx"


'Excel
wb.Sheets(Array("EXCEL", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0

Workbooks("Book10").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Traci\Excel\2020\Excel 3+9.xlsx"

'Lobbying
wb.Sheets(Array("Lobbying", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0

Workbooks("Book11").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Traci\Lobbying\2020\Lobbying 3+9.xlsx"

'Steve
wb.Sheets(Array("Steve's World", "Development", "CC", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0

Workbooks("Book12").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Steve\2020\Steve 3+9.xlsx"

'EMS
wb.Sheets(Array("EMS", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder) To LBound(myOrder) Step -1
        Worksheets(myOrder(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder(x)).Visible = False
    Next x
    On Error GoTo 0

Workbooks("Book13").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Kenny\EMS\2020\EMS 3+9.xlsx"

'PR
wb.Sheets(Array("6200", "Telephone 738", "Building Repairs 742", "Printing 756", "Advertising 758", "Internet 786", "Reference")).Copy

    For x = UBound(myOrder4) To LBound(myOrder4) Step -1
        Worksheets(myOrder4(x)).Move After:=Worksheets(Worksheets.Count)
        Worksheets(myOrder4(x)).Visible = False
    Next x
    On Error GoTo 0

Workbooks("Book14").SaveAs Filename:="T:\Forecast\Prelim Forecast Adjustment\Kenny\PR\2020\PR 3+9.xlsx"

'Turn Settings back on
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

End Sub