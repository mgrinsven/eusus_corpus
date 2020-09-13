Attribute VB_Name = "Module1"



'********************************************************************************
'***************************** GLOBAL VARIABLES *********************************
'********************************************************************************
Private wrkbk As Workbook     'Workbook type variable
Private APath As String       'application path
Private mnuItem As MenuItem   'MenuItem type variable
Private mn As Menu            'Menu type variable
Private msg As String         'stores messages displaying in dialog boxes
Private resp As Integer       'response from dialog boxes
Private NL As String          'contains "new line" characters
Private mnuIV As Variant
Private cap As String

'********************************************************************************
'***** MACRO THAT CREATES CUSTOM MENU FOR OVERVIEW.XLS WORKBOOK ******************
'********************************************************************************
Sub CreateCustomMenu()
Attribute CreateCustomMenu.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    For Each mnuIV In MenuBars(xlWorksheet).Menus("Edit").MenuItems  ' new code 22/7/98
        If mnuIV.Caption = "&Protection" Then mnuIV.Delete           ' new code
    Next                                                             ' new code
    For Each mn In MenuBars(xlWorksheet).Menus
        mn.Delete
    Next
    With MenuBars(xlWorksheet).Menus
        .Add "&File", Restore:=True
        .Add "&Edit", Restore:=True
        .Add "&Insert", Restore:=True
        .Add "&Sectors"
        '.Add "T&itle Block"
        .Add "&Long Summary"
        .Add "&Short Summary"
        .Add "&Uncertainty"
        '.Add "&Protection"
        .Add "&Window", Restore:=True
        .Add "&Help", Restore:=True
    End With
    For Each mnuIV In MenuBars(xlWorksheet).Menus("Insert").MenuItems
        If mnuIV.Index > 3 Then mnuIV.Delete
    Next
    For Each mnuIV In MenuBars(xlWorksheet).Menus("Edit").MenuItems
        cap = mnuIV.Caption
        Select Case cap
            Case "&Delete...", "De&lete Sheet", "Lin&ks...", "&Object"
                mnuIV.Delete
        End Select
    Next
    MenuBars(xlWorksheet).Menus("Edit").MenuItems.AddMenu ("&Protection")
    With MenuBars(xlWorksheet).Menus("Edit").MenuItems("&Protection").MenuItems
        Set mnuItem = .Add("Sheet")
        mnuItem.OnAction = "ProtectSheet"
        mnuItem.StatusBar = "Protects/Unprotects current sheet..."
        If ThisWorkbook.ActiveSheet.ProtectContents = True Then
            mnuItem.Caption = "&Unprotect sheet"
         Else
            mnuItem.Caption = "&Protect sheet"
        End If
        Set mnuItem = .Add("workbook")
        mnuItem.OnAction = "ProtectWorkbook"
        mnuItem.StatusBar = "Protects/Unprotects current workbook..."
        If ThisWorkbook.ProtectStructure = True Then
            mnuItem.Caption = "&Unprotect workbook"
         Else
            mnuItem.Caption = "&Protect workbook"
        End If
    End With
    With MenuBars(xlWorksheet).Menus("Window")
        .MenuItems("New Window").Delete
        .MenuItems("Arrange...").Delete
        .MenuItems("Hide").Delete
        .MenuItems("Unhide...").Delete
        .MenuItems("-").Delete
'        .MenuItems("Split").Delete
'        .MenuItems("Freeze Panes").Delete
'        .MenuItems("-").Delete
    End With
    Err = 0
    With MenuBars(xlWorksheet).Menus("Sectors").MenuItems
        Set mnuItem = .Add("&Energy")
        mnuItem.OnAction = "OpenModule1"
        mnuItem.StatusBar = "Opens Module1.xls with custom menu..."
        Set mnuItem = .Add("&Industrial Processes")
        mnuItem.OnAction = "OpenModule2"
        mnuItem.StatusBar = "Opens Module2.xls with custom menu..."
        Set mnuItem = .Add("&Agriculture")
        mnuItem.OnAction = "OpenModule4"
        mnuItem.StatusBar = "Opens Module4.xls with custom menu..."
        Set mnuItem = .Add("&Land-use Change and Forestry")
        mnuItem.OnAction = "OpenModule5"
        mnuItem.StatusBar = "Opens Module5.xls with custom menu..."
        Set mnuItem = .Add("&Waste")
        mnuItem.OnAction = "OpenModule6"
        mnuItem.StatusBar = "Opens Module6.xls with custom menu..."
    End With
    'Set mnuItem = MenuBars(xlWorksheet).Menus("Title Block").MenuItems.Add("&Show")
    'mnuItem.OnAction = "TitleBlock"
    'mnuItem.StatusBar = "Shows title of inventory..."
    With MenuBars(xlWorksheet).Menus("Long Summary").MenuItems
        Set mnuItem = .Add("Sheet &1 of 3")
        mnuItem.OnAction = "LongSummary1"
        mnuItem.StatusBar = "Energy, Industry..."
        Set mnuItem = .Add("Sheet &2 of 3")
        mnuItem.OnAction = "LongSummary2"
        mnuItem.StatusBar = "Solvents, Agriculture, Land-Use Change & Forestry, Waste..."
        Set mnuItem = .Add("Sheet &3 of 3")
        mnuItem.OnAction = "LongSummary3"
        mnuItem.StatusBar = "International Bunkers, Biomass..."
    End With
    Set mnuItem = MenuBars(xlWorksheet).Menus("Short Summary").MenuItems.Add("&Show")
    mnuItem.OnAction = "ShortSummary"
    mnuItem.StatusBar = "All Sectors..."
    With MenuBars(xlWorksheet).Menus("Uncertainty").MenuItems
        Set mnuItem = .Add("Sheet &1 of 3")
        mnuItem.OnAction = "OU1"
        mnuItem.StatusBar = "Energy, Industry..."
        Set mnuItem = .Add("Sheet &2 of 3")
        mnuItem.OnAction = "OU2"
        mnuItem.StatusBar = "Agriculture, Land-Use Change & Forestry..."
        Set mnuItem = .Add("Sheet &3 of 3")
        mnuItem.OnAction = "OU3"
        mnuItem.StatusBar = "Waste, Bunkers, Biomass..."
    End With
End Sub

'********************************************************************************
'********* MACRO THAT EXECUTES AUTOMATICALLY AFTER OPENING OVERVIEW.XLS *********
'********************************************************************************
Sub Auto_Open()
Attribute Auto_Open.VB_ProcData.VB_Invoke_Func = " \n14"
    If Right(Application.Path, 1) <> Chr(92) Then
        APath = ThisWorkbook.Path & Chr(92)
     Else
        APath = ThisWorkbook.Path
    End If
    NL = Chr(13) & Chr(10) 'sets "new line" character
    Sheets("head").Activate
    Range("C16").Select
    Application.Windows("OVERVIEW.XLS").OnWindow = "ActivateOverView" 'Macro "ActivateOverView" executes each time overview.xls is activated
    ThisWorkbook.Saved = True
    ThisWorkbook.OnSheetActivate = "Module1.SheetActivate"
End Sub

'********************************************************************************
'****** MACRO THAT EXECUTES AUTOMATICALLY AFTER ACTIVATING THIS WORKBOOK ********
'********************************************************************************
Sub ActivateOverView()
Attribute ActivateOverView.VB_ProcData.VB_Invoke_Func = " \n14"
    CreateCustomMenu 'calls macro that creates custom menu for this workbook
    If ThisWorkbook.ProtectStructure = True Then
        MenuBars(xlWorksheet).Menus("Edit").MenuItems("Protection").MenuItems(2).Caption = "&Unprotect workbook"
     Else
        MenuBars(xlWorksheet).Menus("Edit").MenuItems("Protection").MenuItems(2).Caption = "&Protect workbook"
    End If
End Sub

'********************************************************************************
'****** MACRO THAT CHECKS IF CHANGES WERE MADE TO THIS WORKBOOK *****************
'********************************************************************************
'If changes were made to this workbook dialog box appears that asks the user    |
'if he wants to save these changes.                                             |
'-------------------------------------------------------------------------------|
Sub AskToSaveWorkbook()
Attribute AskToSaveWorkbook.VB_ProcData.VB_Invoke_Func = " \n14"
    If Not ThisWorkbook.Saved Then
        msg = "Changes have been made to OVERVIEW.XLS! Would you like to save these changes?"
        resp = MsgBox(msg, 36, "IPCC")
        If resp = vbYes Then
            ThisWorkbook.Save
         Else
            ThisWorkbook.Saved = True
        End If
    End If
End Sub

'********************************************************************************
'****** MACRO THAT EXECUTES AUTOMATICALLY WHEN THIS WORKBOOK IS CLOSED **********
'********************************************************************************
'This macro checks if some of the modules (module1.xls - module6.xls) are open. |
'If some of the modules exist overview.xls closes them automatically asking a   |
'to save changes (if made)                                                      |
'-------------------------------------------------------------------------------|
Sub Auto_Close()
Attribute Auto_Close.VB_ProcData.VB_Invoke_Func = " \n14"
    For Each wrkbk In Application.Workbooks
        Select Case LCase(wrkbk.Name)
            Case "module1.xls"
                Application.Run "MODULE1.XLS!Module1.AskToSaveWorkbook"
                wrkbk.Close
            Case "module2.xls"
                Application.Run "MODULE2.XLS!Module1.AskToSaveWorkbook"
                wrkbk.Close
            Case "module4.xls"
                Application.Run "MODULE4.XLS!Module1.AskToSaveWorkbook"
                wrkbk.Close
            Case "module5.xls"
                Application.Run "MODULE5.XLS!Module1.AskToSaveWorkbook"
                wrkbk.Close
            Case "module6.xls"
                Application.Run "MODULE6.XLS!Module1.AskToSaveWorkbook"
                wrkbk.Close
        End Select
    Next
    AskToSaveWorkbook
    MenuBars(xlWorksheet).Reset
End Sub

'********************************************************************************
'********* MACRO FOR OPENING AND ACTIVATING MODULE1.XLS *************************
'********************************************************************************
Sub OpenModule1()
Attribute OpenModule1.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim eflag As Boolean
    On Error Resume Next
    eflag = False
    For Each wrkbk In Workbooks
        If LCase(wrkbk.Name) = "module1.xls" Then eflag = True
    Next
    If Not eflag Then
        Workbooks.Open (APath & "module1.xls")
        Application.Run "MODULE1.XLS!Module1.Auto_Open"
        Application.Run "MODULE1.XLS!Module1.CreateCustomMenu"
     Else
        Workbooks("MODULE1.XLS").Activate
        Application.Run "MODULE1.XLS!Module1.CreateCustomMenu"
    End If
End Sub

'********************************************************************************
'********* MACRO FOR OPENING AND ACTIVATING MODULE2.XLS *************************
'********************************************************************************
Sub OpenModule2()
Attribute OpenModule2.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim eflag As Boolean
    On Error Resume Next
    eflag = False
    For Each wrkbk In Workbooks
        If LCase(wrkbk.Name) = "module2.xls" Then eflag = True
    Next
    If Not eflag Then
        Workbooks.Open (APath & "module2.xls")
        Application.Run "MODULE2.XLS!Module1.Auto_Open"
        Application.Run "MODULE2.XLS!Module1.CreateCustomMenu"
     Else
        Workbooks("MODULE2.XLS").Activate
        Application.Run "MODULE2.XLS!Module1.CreateCustomMenu"
    End If
End Sub

'********************************************************************************
'********* MACRO FOR OPENING AND ACTIVATING MODULE4.XLS *************************
'********************************************************************************
Sub OpenModule4()
Attribute OpenModule4.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim eflag As Boolean
    On Error Resume Next
    eflag = False
    For Each wrkbk In Workbooks
        If LCase(wrkbk.Name) = "module4.xls" Then eflag = True
    Next
    If Not eflag Then
        Workbooks.Open (APath & "module4.xls")
        Application.Run "MODULE4.XLS!Module1.Auto_Open"
        Application.Run "MODULE4.XLS!Module1.CreateCustomMenu"
     Else
        Workbooks("MODULE4.XLS").Activate
        Application.Run "MODULE4.XLS!Module1.CreateCustomMenu"
    End If
End Sub

'********************************************************************************
'********* MACRO FOR OPENING AND ACTIVATING MODULE5.XLS *************************
'********************************************************************************
Sub OpenModule5()
Attribute OpenModule5.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim eflag As Boolean
    On Error Resume Next
    eflag = False
    For Each wrkbk In Workbooks
        If LCase(wrkbk.Name) = "module5.xls" Then eflag = True
    Next
    If Not eflag Then
        Workbooks.Open (APath & "module5.xls")
        Application.Run "MODULE5.XLS!Module1.Auto_Open"
        Application.Run "MODULE5.XLS!Module1.CreateCustomMenu"
     Else
        Workbooks("MODULE5.XLS").Activate
        Application.Run "MODULE5.XLS!Module1.CreateCustomMenu"
    End If
End Sub

'********************************************************************************
'********* MACRO FOR OPENING AND ACTIVATING MODULE6.XLS *************************
'********************************************************************************
Sub OpenModule6()
Attribute OpenModule6.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim eflag As Boolean
    On Error Resume Next
    eflag = False
    For Each wrkbk In Workbooks
        If LCase(wrkbk.Name) = "module6.xls" Then eflag = True
    Next
    If Not eflag Then
        Workbooks.Open (APath & "module6.xls")
        Application.Run "MODULE6.XLS!Module1.Auto_Open"
        Application.Run "MODULE6.XLS!Module1.CreateCustomMenu"
     Else
        Workbooks("MODULE6.XLS").Activate
        Application.Run "MODULE6.XLS!Module1.CreateCustomMenu"
    End If
End Sub

'********************************************************************************
'********* MACRO FOR OPENING TITLE OF INVENTORY SHEET OF OVERVIEW.XLS ***********
'********************************************************************************
Sub TitleBlock()
Attribute TitleBlock.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("head").Activate
    Range("C16").Select
End Sub

'********************************************************************************
'********* MACRO FOR OPENING SHEET1 OF LONG SUMMARY *****************************
'********************************************************************************
Sub LongSummary1()
Attribute LongSummary1.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Table7As1").Activate
    Range("A1").Select
End Sub

'********************************************************************************
'********* MACRO FOR OPENING SHEET2 OF LONG SUMMARY *****************************
'********************************************************************************
Sub LongSummary2()
Attribute LongSummary2.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Table7As2").Activate
    Range("A1").Select
End Sub

'********************************************************************************
'********* MACRO FOR OPENING SHEET3 OF LONG SUMMARY *****************************
'********************************************************************************
Sub LongSummary3()
Attribute LongSummary3.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Table7As3").Activate
    Range("A1").Select
End Sub

'********************************************************************************
'****************** MACRO FOR OPENING SHORT SUMMARY *****************************
'********************************************************************************
Sub ShortSummary()
Attribute ShortSummary.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Table7Bs1").Activate
    Range("A1").Select
End Sub

'********************************************************************************
'********* MACRO FOR OPENING SHEET1 OF OVERVIEW/UNCERTAINTY *********************
'********************************************************************************
Sub OU1()
Attribute OU1.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Table8As1").Activate
    Range("A1").Select
End Sub

'********************************************************************************
'********* MACRO FOR OPENING SHEET2 OF OVERVIEW/UNCERTAINTY *********************
'********************************************************************************
Sub OU2()
Attribute OU2.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Table8As2").Activate
    Range("A1").Select
End Sub

'********************************************************************************
'********* MACRO FOR OPENING SHEET3 OF OVERVIEW/UNCERTAINTY *********************
'********************************************************************************
Sub OU3()
Attribute OU3.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Table8As3").Activate
    Range("A1").Select
End Sub

Sub ProtectSheet()
Attribute ProtectSheet.VB_ProcData.VB_Invoke_Func = " \n14"
    With ThisWorkbook.ActiveSheet
        If .ProtectContents = True Then
            .Unprotect
            resp = MsgBox("The contents of this sheet have been unprotected!", vbInformation, "IPCC")
            MenuBars(xlWorksheet).Menus("Edit").MenuItems("Protection").MenuItems(1).Caption = "&Protect sheet"
         Else
            .Protect
            resp = MsgBox("The contents of this sheet have been protected!", vbInformation, "IPCC")
            MenuBars(xlWorksheet).Menus("Edit").MenuItems("Protection").MenuItems(1).Caption = "&Unprotect sheet"
        End If
    End With
End Sub

Sub ProtectWorkbook()
Attribute ProtectWorkbook.VB_ProcData.VB_Invoke_Func = " \n14"
    
    With ThisWorkbook
        If .ProtectStructure = True Then
            .Unprotect
            resp = MsgBox("The contents of this workbook have been unprotected!", vbInformation, "IPCC")
            MenuBars(xlWorksheet).Menus("Edit").MenuItems("Protection").MenuItems(2).Caption = "&Protect workbook"
         Else
            .Protect
            resp = MsgBox("The contents of this workbook have been protected!", vbInformation, "IPCC")
            MenuBars(xlWorksheet).Menus("Edit").MenuItems("Protection").MenuItems(2).Caption = "&Unprotect workbook"
        End If
    End With
End Sub

Sub SheetActivate()
Attribute SheetActivate.VB_ProcData.VB_Invoke_Func = " \n14"
    If ThisWorkbook.ActiveSheet.ProtectContents = True Then
        MenuBars(xlWorksheet).Menus("Edit").MenuItems("Protection").MenuItems(1).Caption = "&Unprotect sheet"
     Else
        MenuBars(xlWorksheet).Menus("Edit").MenuItems("Protection").MenuItems(1).Caption = "&Protect sheet"
    End If
End Sub

