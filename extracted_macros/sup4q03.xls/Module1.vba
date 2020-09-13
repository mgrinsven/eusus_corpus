Attribute VB_Name = "Module1"







' Print final supplemental package
'
Sub Macro1()
Attribute Macro1.VB_Description = "Print supplemental package"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("COVERPAGE").Select
    ActiveSheet.PageSetup.PrintArea = "COVER"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    
     Sheets("EARNINGS").Select
    ActiveSheet.PageSetup.PrintArea = "EARNINGS"
    With ActiveSheet.PageSetup
        .CenterFooter = "1"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    With ActiveSheet.PageSetup
        .CenterFooter = ""
    End With
    
    Sheets("ANNUAL EARNINGS").Select
    ActiveSheet.PageSetup.PrintArea = "ANNEARN"
    With ActiveSheet.PageSetup
        .CenterFooter = "2"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    With ActiveSheet.PageSetup
        .CenterFooter = ""
    End With
        
    Sheets("P&C").Select
    ActiveSheet.PageSetup.PrintArea = "HIGHLIGHTS"
    With ActiveSheet.PageSetup
        .CenterFooter = "3"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    With ActiveSheet.PageSetup
        .CenterFooter = ""
    End With
    
    Sheets("QTR SUMMARY").Select
    ActiveSheet.PageSetup.PrintArea = "SUMMARIES"
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterFooter = "4"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
        
    Sheets("ANNUAL SUMMARY").Select
    ActiveSheet.PageSetup.PrintArea = "YRSUMM"
    With ActiveSheet.PageSetup
        .CenterFooter = "5"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    With ActiveSheet.PageSetup
        .CenterFooter = ""
    End With
    
    Sheets("Specialty Premium Distribution").Select
    ActiveSheet.PageSetup.PrintArea = "SPECPREM"
    With ActiveSheet.PageSetup
        .CenterFooter = "6"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    With ActiveSheet.PageSetup
        .CenterFooter = ""
    End With
            
    Sheets("Specialty Supplemental Ops Info").Select
    ActiveSheet.PageSetup.PrintArea = "MJRSUMM"
    With ActiveSheet.PageSetup
        .CenterFooter = "7"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    With ActiveSheet.PageSetup
        .CenterFooter = ""
    End With
        
    Sheets("QTR COR").Select
    ActiveSheet.PageSetup.PrintArea = "GRAPHICS"
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterFooter = "8"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
        
    Sheets("CORVSINDUSTRY").Select
    ActiveSheet.PageSetup.PrintArea = "AFGIND"
    With ActiveSheet.PageSetup
        .CenterFooter = "9"
    End With
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    With ActiveSheet.PageSetup
        .CenterFooter = ""
    End With
                                                   
    Sheets("COVERPAGE").Select
    Range("A1").Select

End Sub


' Go to summary of earnings
'
Sub Macro2()
Attribute Macro2.VB_Description = "Go to summary of earnings page"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Earnings").Select
    Range("A1").Select
End Sub


' Go to conlidated balance sheet data
'
Sub Macro3()
Attribute Macro3.VB_Description = "Go to balance sheet data page"
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Balance Sheet").Select
    Range("A1").Select
End Sub


' Go to property & casualty insurance highlights
'
Sub Macro4()
Attribute Macro4.VB_Description = "Go to p&c insurance highlights page"
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("P&C").Select
    Range("A1").Select
End Sub


' Go to property & casualty underwriting summaries
'
Sub Macro5()
Attribute Macro5.VB_Description = "Go to p&c underwriting summaries page"
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("QTR SUMMARY").Select
    Range("A1").Select
End Sub


' Go to annuity, life & health insurance hightlights
'
Sub Macro6()
Attribute Macro6.VB_Description = "Go to annuity, life & health highlights page"
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Annuity").Select
    Range("A1").Select
End Sub


' Go to investment portfolio
'
Sub Macro7()
Attribute Macro7.VB_Description = "Go to investment portfolio page"
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Investment Portfolio").Select
    Range("A1").Select
End Sub


' Print summary of earnings
'
Sub Macro8()
Attribute Macro8.VB_Description = "Print summary of earnings "
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("EARNINGS").Select
    ActiveSheet.PageSetup.PrintArea = "EARNINGS"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Print balance sheet data
'
Sub Macro9()
Attribute Macro9.VB_Description = "Print balance sheet data "
Attribute Macro9.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("BALANCE SHEET").Select
    ActiveSheet.PageSetup.PrintArea = "CAPITAL"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
    
    
' Print p&c hightlights
'
Sub Macro10()
Attribute Macro10.VB_Description = "Print p&c insurance highlights"
Attribute Macro10.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("P&C").Select
    ActiveSheet.PageSetup.PrintArea = "HIGHLIGHTS"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
    

' Print p&c underwriting summaries
'
Sub Macro11()
Attribute Macro11.VB_Description = "Print p&c underwriting summaries"
Attribute Macro11.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("QTR SUMMARY").Select
    ActiveSheet.PageSetup.PrintArea = "SUMMARIES"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Print annuity highlights
'
Sub Macro12()
Attribute Macro12.VB_Description = "Print annuity, life & health highlights"
Attribute Macro12.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("ANNUITY").Select
    ActiveSheet.PageSetup.PrintArea = "ANNUITY"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Print investment portfolio
'
Sub Macro13()
Attribute Macro13.VB_Description = "Print investment portfolio"
Attribute Macro13.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("INVESTMENT PORTFOLIO").Select
    ActiveSheet.PageSetup.PrintArea = "PORTFOLIO"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to cover page
'
Sub Macro14()
Attribute Macro14.VB_Description = "Go to table of contents"
Attribute Macro14.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("COVERPAGE").Select
    Range("A1").Select
End Sub


' Print graphics
'
Sub Macro15()
Attribute Macro15.VB_Description = "print graphics"
Attribute Macro15.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("QTR COR").Select
    ActiveSheet.PageSetup.PrintArea = "GRAPHICS"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to graphics
'
Sub Macro16()
Attribute Macro16.VB_Description = "go to graphics"
Attribute Macro16.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Qtr COR").Select
    Range("A1").Select
End Sub


' Print yearly summary
'
Sub Macro20()
    Sheets("ANNUAL SUMMARY").Select
    ActiveSheet.PageSetup.PrintArea = "YRSUMM"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to yearly summary
'
Sub Macro21()
    Sheets("Annual Summary").Select
    Range("A1").Select
End Sub

' Print investments
'
Sub Macro22()
    Sheets("INVESTMENT SCHEDULE").Select
    ActiveSheet.PageSetup.PrintArea = "INVESTMENT"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to investments
'
Sub Macro23()
    Sheets("Investment Schedule").Select
    Range("A1").Select
End Sub


' Print qtr annuity
'
Sub Macro24()
    Sheets("QTR ANNUITY").Select
    ActiveSheet.PageSetup.PrintArea = "QTRANN"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to qtr annuity
'
Sub Macro25()
    Sheets("Qtr Annuity").Select
    Range("A1").Select
End Sub


' Print year annuity
'
Sub Macro26()
    Sheets("ANNUAL ANNUITY").Select
    ActiveSheet.PageSetup.PrintArea = "YRANN"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to year annuity
'
Sub Macro27()
    Sheets("Annual Annuity").Select
    Range("A1").Select
End Sub


' Print annual earnings
'
Sub Macro28()
    Sheets("ANNUAL EARNINGS").Select
    ActiveSheet.PageSetup.PrintArea = "ANNEARN"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to annual earnings
'
Sub Macro29()
    Sheets("Annual Earnings").Select
    Range("A1").Select
End Sub


' Print qtr afg vs ind
'
Sub Macro30()
    Sheets("CORVSINDUSTRY").Select
    ActiveSheet.PageSetup.PrintArea = "AFGIND"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to qtr afg vs ind
'
Sub Macro31()
    Sheets("CORvsIndustry").Select
    Range("A1").Select
End Sub


' Print year specialty premium
'
Sub Macro32()
    Sheets("SPECIALTY PREMIUM DISTRIBUTION").Select
    ActiveSheet.PageSetup.PrintArea = "SPECPREM"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub


' Go to year specialty premium
'
Sub Macro33()
    Sheets("Specialty Premium Distribution").Select
    Range("A1").Select
End Sub


' Go to Major Bus Ann Summary
'
Sub Macro34()
    Sheets("Specialty Supplemental Ops Info").Select
    Range("A1").Select
End Sub


' Print Major Bus Ann Summary
'
Sub Macro35()
    Sheets("Specialty Supplemental Ops Info").Select
    ActiveSheet.PageSetup.PrintArea = "MJRSUMM"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub

