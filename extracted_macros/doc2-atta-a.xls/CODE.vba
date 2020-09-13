Attribute VB_Name = "CODE"
Option Explicit
Option Compare Text
Option Base 1

Global Const gcVersion = "5.0.2.9.4"
Global Const gGLDIINIKey = "SOFTWARE\ORACLE\GLDI"
Global Const gORACLEINIKey = "SOFTWARE\ORACLE"
Global Const HKEY_CURRENT_USER = &H80000001
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const REG_SZ = 1
Public Const ERROR_SUCCESS = 0&

Type OFSTRUCT
    cBytes As String * 1
    fFixedDisk As String * 1
    nErrCode As Integer
    rserved As String * 4
    szPathName As String * 128
End Type

Public Const OF_EXIST = &H4000

Declare Function RegOpenKeyEx Lib "ADVAPI32.DLL" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "ADVAPI32.DLL" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegCloseKey Lib "ADVAPI32.DLL" (ByVal hKey As Long) As Long
Declare Function OpenFile32 Lib "KERNEL32" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Dim smSectionName As String
Dim smIniFileName As String
Dim nmPrivInit As Integer

Global gOracleHome As String
Global gApplicationPathName As String

Const Max_EntryBuffer = 255
Const gFSGImagePath = "FSGIMAGE"

Const mcRight = "RIGHT"
Const mcLeft = "LEFT"

Function CodeVersion() As Variant
    CodeVersion = gcVersion
End Function

Sub StartGLDIAddinCode(ApplicationFilename As Variant)
Dim Result As Integer
Dim IsRunning As Boolean
Dim Dummy As Object

    On Error Resume Next
    Application.Run "LoadINISettings"
    
    Err = 0
    Set Dummy = CreateObject("oraGLDI32.IsRunning")
    If Err = 0 Then
        DoEvents
        Set Dummy = Nothing
        IsRunning = True
    Else
        IsRunning = False
    End If
        
    If Not IsRunning Then
        Result = Shell(ApplicationFilename, 4)
        DoEvents
    End If
    
End Sub

Sub Auto_Open()
Dim FullAddinFilename As Variant
Dim ApplicationPathname As Variant
Dim MS As Object
Dim testGLSA As Variant
Dim Counter1 As Integer
Dim IniFileName As Variant
Dim RunGLDIOnWorkbookOpen As Variant
Dim ApplicationFilename As Variant
Dim CurrentCellValue As Variant

Static AlreadyRun As Boolean

    On Error Resume Next

    If Not AlreadyRun Then

        Set MS = ActiveWorkbook.Modules("CODE")
        If MS.Visible = True Then
            MS.Visible = False
        End If

        If (InStr(Application.OperatingSystem, "95") > 0) Or (InStr(Application.OperatingSystem, "WINDOWS (32-BIT) 4") > 0) Then
            ApplicationPathname = ReadRegistryKeyString(gGLDIINIKey & "\GLDI90", "APPLICATION PATH", "C:\ORAWIN95\GLDI90")
        Else
            ApplicationPathname = ReadRegistryKeyString(gGLDIINIKey & "\GLDI90", "APPLICATION PATH", "C:\ORANT\GLDI90")
        End If
        RunGLDIOnWorkbookOpen = ReadRegistryKeyString(gGLDIINIKey & "\GLDI90", "RUN GLDI ON WORKBOOK OPEN", "TRUE")
        ApplicationFilename = ReadRegistryKeyString(gGLDIINIKey & "\GLDI90", "APPLICATION EXE", "GLDI9032.EXE")
        gOracleHome = ReadRegistryKeyString(gORACLEINIKey, "ORACLE_HOME", "")
        gApplicationPathName = ApplicationPathname
        FullAddinFilename = ApplicationPathname & "\" & AddinFileName()

        testGLSA = Run(AddinFileName() & "!IsGLSALoaded")
        If TypeName(testGLSA) = "EMPTY" Then
            Application.StatusBar = "Loading ORACLE GL Desktop Integrator Add-in..."
            Workbooks.Open (FullAddinFilename)
        End If
        If RunGLDIOnWorkbookOpen = "TRUE" Then
            StartGLDIAddinCode ApplicationFilename
        End If

        For Counter1 = 1 To ActiveWorkbook.Worksheets.Count
            With ActiveWorkbook.Worksheets(Counter1)
                
                CurrentCellValue = .Cells(1, 2).Value
            
                Select Case CurrentCellValue
                Case "ASSET"
                    .OnEntry = "Flag_Journal_Row"
                    .OnDoubleClick = "JournalDoubleClick"
            
                Case "INVENTORY"
                    .OnEntry = "Flag_Journal_Row"
                    .OnDoubleClick = "JournalDoubleClick"
            
                Case "JOURNAL"
                    .OnEntry = "Flag_Journal_Row"
                    .OnDoubleClick = "JournalDoubleClick"
            
                Case "BUDGET"
                    .OnEntry = "Flag_Row"
                    .OnDoubleClick = "BudgetDoubleClick"
            
                Case "REPORT"
                    .OnEntry = "RWOnEntry"
                    .OnDoubleClick = "ReportWizardDoubleClick"
                    .OnSheetActivate = "RWActivateSheet"
                    .OnSheetDeactivate = "RWDeactivateSheet"
                
                    If Int(.Cells(1, 4) * 100000) = Int(.Cells(1, 6) * 100000) Then
                        Counter1 = 100
                        While Counter1 > 0 And ((Int(.Cells(1, 4).Value * 100000) <> Int(.Cells(1, 3).Value * 100000)) Or (Int(.Cells(1, 6).Value * 100000) <> Int(.Cells(1, 3).Value * 100000)))
                            .Cells(1, 7).Value = 101 - Counter1
                            .Cells(1, 4).Value = .Cells(1, 3).Value
                            .Cells(1, 6).Value = .Cells(1, 3).Value
                            Counter1 = Counter1 - 1
                        Wend
                    End If
            
                Case "REPORT OUTPUT"
                    If .Cells(1, 3) = "TEMPLATE" Then
                        .OnDoubleClick = "RWOutputTemplateDoubleClick"
                    Else
                        .OnDoubleClick = "RWFSGOutputDoubleClick"
                    End If
                    
                End Select
            
            End With
    
        Next Counter1
    
        Set MS = Nothing
        Application.StatusBar = False

        AlreadyRun = True

    End If

End Sub

Sub Auto_Close()
Dim Counter1 As Integer
Dim tmpValue As Variant
    On Error Resume Next
    Counter1 = 0
    Counter1 = ActiveWorkbook.Worksheets.Count
    If Counter1 > 0 Then
        For Counter1 = 1 To ActiveWorkbook.Worksheets.Count
            With ActiveWorkbook.Worksheets(Counter1)
                tmpValue = ""
                tmpValue = .Cells(1, 2).Value
                If tmpValue = "REPORT" Then
                    If (Int(.Cells(1, 3).Value * 100000) <= Int(.Cells(1, 4).Value * 100000)) Then
                        Counter1 = 100
                        While Counter1 > 0 And ((Int(.Cells(1, 4).Value * 100000) <> Int(.Cells(1, 3).Value * 100000)) Or (Int(.Cells(1, 6).Value * 100000) <> Int(.Cells(1, 3).Value * 100000)))
                            .Cells(1, 7).Value = 101 - Counter1
                            .Cells(1, 4).Value = .Cells(1, 3).Value
                            .Cells(1, 6).Value = .Cells(1, 3).Value
                            Counter1 = Counter1 - 1
                        Wend
                    Else
                        .Cells(1, 6).Value = .Cells(1, 3).Value
                    End If
                End If
            End With
        Next Counter1
    End If
End Sub

Sub RWFSGOutputDoubleClick()
    On Error Resume Next
    Application.Run "GLDIRWOutputLOV"
End Sub

Sub ReportWizardDoubleClick()
    On Error Resume Next
    Application.Run "GLDIRWLOV"
End Sub

Sub RWOnEntry()
    On Error Resume Next
    Application.Run "GLDIRWOnEntry"
End Sub

Sub CodeRWRefreshSampleValues()
    On Error Resume Next
    Application.Run "RWRefreshSampleValues"
End Sub

Sub CodeRWBuildColumnHeadings()
    On Error Resume Next
    Application.Run "RWBuildColumnHeadings"
End Sub

Sub CodeRWTrimColumnHeadings()
    On Error Resume Next
    Application.Run "RWTrimColumnHeadings"
End Sub

Sub RWActivateSheet()
    On Error Resume Next
    Application.Run "GLDIActivateRWSheet"
End Sub

Sub RWDeactivateSheet()
    On Error Resume Next
    Application.Run "GLDIDeactivateRWSheet"
    DoEvents
    Application.Run "ActivateExcel"
End Sub

Sub Flag_Row()
Dim FirstRow As Integer
    On Error Resume Next
    FirstRow = 0
    FirstRow = ActiveWorkbook.Worksheets("Criteria" & Cells(1, 1).Value).Range("FirstDataRow" & Cells(1, 1).Value).Value
    If FirstRow = 0 Then Application.Run Macro:="GLDAFlag_Row"
    If ActiveCell.Row >= FirstRow Then Application.Run Macro:="GLDAFlag_Row"
End Sub

Sub Flag_Journal_Row()
    On Error Resume Next
    If ActiveCell.Row >= _
      ActiveWorkbook.Worksheets("Criteria" & Cells(1, 1).Value).Range("FirstDataRow" & Cells(1, 1).Value).Value _
      Then Application.Run Macro:="GLSAFlag_Journal_Row"
End Sub

Sub UpdateBudgetGraphView()
    On Error Resume Next
    Application.Run Macro:="GLDAUpdateBudgetGraphView"
End Sub

Sub JournalDoubleClick()
    On Error Resume Next
    Application.Run "GLDIJTLOV"
End Sub

Sub BudgetDoubleClick()
    On Error Resume Next
    Application.Run "GLDIBDLOV"
End Sub

Sub RWAddBuildTrimButtons(LayoutSheetName As Variant, MaskCellAddress As Variant)
Dim NewButton As Object
Dim ButtonName As Variant
Dim LayoutSheet As Object
Dim MaskCell As Object

    On Error Resume Next
    Set LayoutSheet = ActiveWorkbook.Worksheets(LayoutSheetName)
    Set MaskCell = LayoutSheet.Range(MaskCellAddress)
    
    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("RWBuildButton").Name
    
    If ButtonName = "" Then
        Set NewButton = LayoutSheet.Buttons.Add(MaskCell.Left, MaskCell.Top, 50, 14)
        NewButton.OnAction = "CodeRWBuildColumnHeadings"
        NewButton.Characters.Text = NLSGetString(2691, "Build", 3124)
        NewButton.Name = "RWBuildButton"
        CodeFormatButton NewButton, xlMove
        Set NewButton = Nothing
    End If
    
    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("RWTrimButton").Name
    
    If ButtonName = "" Then
        Set NewButton = LayoutSheet.Buttons.Add(MaskCell.Left + 50, MaskCell.Top, 50, 14)
        NewButton.OnAction = "CodeRWTrimColumnHeadings"
        NewButton.Characters.Text = NLSGetString(2692, "Trim", 3126)
        NewButton.Name = "RWTrimButton"
        CodeFormatButton NewButton, xlMove
        Set NewButton = Nothing
    End If

End Sub

Sub RWAddMaskRefreshButton(LayoutSheetName As Variant, MaskCellAddress As Variant)
Dim NewButton As Object
Dim ButtonName As Variant
Dim LayoutSheet As Object
Dim MaskCell As Object

    On Error Resume Next
    
    Set LayoutSheet = ActiveWorkbook.Worksheets(LayoutSheetName)
    Set MaskCell = LayoutSheet.Range(MaskCellAddress)
    
    ButtonName = ""
    On Error Resume Next
    ButtonName = LayoutSheet.DrawingObjects("RWRefreshButton").Name
    
    If ButtonName = "" Then
        Set NewButton = LayoutSheet.Buttons.Add(MaskCell.Left, MaskCell.Top, 58.5, 14)
        NewButton.OnAction = "CodeRWRefreshSampleValues"
        NewButton.Characters.Text = NLSGetString(2693, "Refresh", 3127)
        NewButton.Name = "RWRefreshButton"
        CodeFormatButton NewButton, xlMoveAndSize
        Set NewButton = Nothing
    End If
    
End Sub

Function ReadRegistryKeyString(strSubKeys As String, strValName As String, strDefault As String) As String
Dim lngResult As Long
Dim lngHandle As Long
Dim lngcbData As Long
Dim strRet As String
    
    On Error Resume Next

    If Not ERROR_SUCCESS = RegOpenKeyEx(HKEY_CURRENT_USER, strSubKeys, 0&, KEY_READ, lngHandle) Then
        ReadRegistryKeyString = strDefault
        Exit Function
    End If
       
    If Not ERROR_SUCCESS = RegQueryValueEx(lngHandle, strValName, 0&, REG_SZ, ByVal strRet, lngcbData) Then
        ReadRegistryKeyString = strDefault
        Exit Function
    End If
    
    strRet = Space(lngcbData)
    lngResult = RegQueryValueEx(lngHandle, strValName, 0&, REG_SZ, ByVal strRet, lngcbData)
    lngResult = RegCloseKey(lngHandle)
      
    If lngcbData > 0 Then
        lngcbData = lngcbData - 1
    End If
    ReadRegistryKeyString = Left$(strRet, lngcbData)

End Function

Function OpenFile(ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Variant) As Integer
Dim RetInt As Integer
Dim RetLng As Long
    On Error Resume Next
    RetLng = OpenFile32(lpFileName, lpReOpenBuff, CLng(wStyle))
    RetInt = CInt(RetLng)
    OpenFile = RetInt
End Function

Sub GLDIAddOutlineButtons(LayoutSheetName As Variant, SheetType As Variant, TemplateStyle As Integer)
Dim NewButton As Object
Dim ButtonName As Variant
Dim LayoutSheet As Object
Dim MaskCell As Object
Dim ContextButton As Object
Dim pDrawingObjects As Boolean
Dim pContents As Boolean
Dim pScenarios As Boolean

    On Error Resume Next

    Application.ScreenUpdating = False
    Set LayoutSheet = ActiveWorkbook.Worksheets(LayoutSheetName)
    Set MaskCell = LayoutSheet.Cells(1, 2)
    With LayoutSheet
        pDrawingObjects = .ProtectDrawingObjects
        pContents = .ProtectContents
        pScenarios = .ProtectScenarios
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Unprotect
        End If
    End With

    If SheetType = "JOURNAL" Then
        If TemplateStyle = 1 Then
            MaskCell.Offset(1, 0).EntireRow.RowHeight = 18
        Else
            MaskCell.Offset(1, 0).EntireRow.RowHeight = 20.25
        End If
    Else
        MaskCell.Offset(1, 0).EntireRow.RowHeight = 20.25
    End If
    
    If (SheetType = "BUDGET" Or SheetType = "ACTUAL") And TemplateStyle = 0 Then
        LayoutSheet.DrawingObjects("GLDIDescriptionsButton").Delete
        LayoutSheet.DrawingObjects("GLDIFilterButton").Delete
        LayoutSheet.DrawingObjects("GLDIResetFilterButton").Delete
        LayoutSheet.DrawingObjects("GLDIContextButton").Delete
    End If

    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("GLDIContextButton").Name
    
    If ButtonName = "" Then
        Set NewButton = LayoutSheet.Buttons.Add((MaskCell.Left / 2.5), MaskCell.Top + 3, 98, 14)
        Set ContextButton = NewButton
        NewButton.OnAction = "CodeGLDIContextButton"
        NewButton.Characters.Text = NLSGetString(2695, "View Context", 3129)
        NewButton.Name = "GLDIContextButton"
        CodeFormatButton NewButton, xlFreeFloating
        Set NewButton = Nothing
    Else
        Set ContextButton = LayoutSheet.DrawingObjects("GLDIContextButton")
    End If
    
    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("GLDIDescriptionsButton").Name
    
    If ButtonName = "" And (SheetType = "BUDGET" Or SheetType = "ACTUAL") And TemplateStyle > 0 Then
        Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (ContextButton.Width), MaskCell.Top + 3, 98, 14)
        NewButton.OnAction = "CodeGLDIDescriptionsButton"
        NewButton.Characters.Text = NLSGetString(2696, "View Desc", 3131)
        NewButton.Name = "GLDIDescriptionsButton"
        CodeFormatButton NewButton, xlFreeFloating
        Set NewButton = Nothing
        LayoutSheet.DrawingObjects("GLDIFilterButton").Delete
        LayoutSheet.DrawingObjects("GLDIResetFilterButton").Delete
    End If
    
    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("GLDIHeaderButton").Name
    
    If ButtonName = "" And SheetType = "JOURNAL" And TemplateStyle = 1 Then
        Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (ContextButton.Width), MaskCell.Top + 3, 98, 14)
        NewButton.OnAction = "CodeGLDIHeaderButton"
        NewButton.Characters.Text = NLSGetString(2697, "View Header", 3132)
        NewButton.Name = "GLDIHeaderButton"
        CodeFormatButton NewButton, xlFreeFloating
        Set NewButton = Nothing
    End If
    
    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("GLDIFilterButton").Name
    
    If ButtonName = "" And (SheetType = "BUDGET" Or SheetType = "ACTUAL") Then
        If TemplateStyle = 0 Then
            Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (ContextButton.Width) + 10, MaskCell.Top + 3, 98, 14)
        Else
            Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (2 * ContextButton.Width) + 10, MaskCell.Top + 3, 98, 14)
        End If
        NewButton.OnAction = "CodeGLDIFilterButton"
        NewButton.Characters.Text = NLSGetString(2698, "View Filter", 3133)
        NewButton.Name = "GLDIFilterButton"
        CodeFormatButton NewButton, xlFreeFloating
        Set NewButton = Nothing
    End If
    
    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("GLDIResetFilterButton").Name
    
    If ButtonName = "" And (SheetType = "BUDGET" Or SheetType = "ACTUAL") Then
        If TemplateStyle = 0 Then
            Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (2 * ContextButton.Width) + 10, MaskCell.Top + 3, 98, 14)
            'Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (2 * ContextButton.Width) + 10, MaskCell.Top + 3, 75, 14)
        Else
            Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (3 * ContextButton.Width) + 10, MaskCell.Top + 3, 98, 14)
            'Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (3 * ContextButton.Width) + 10, MaskCell.Top + 3, 75, 14)
        End If
        NewButton.OnAction = "CodeGLDIResetFilterButton"
        NewButton.Characters.Text = NLSGetString(2699, "Reset Filter", 3134)
        NewButton.Name = "GLDIResetFilterButton"
        CodeFormatButton NewButton, xlFreeFloating
        Set NewButton = Nothing
    End If

    With LayoutSheet
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Protect DrawingObjects:=pDrawingObjects, contents:=pContents, Scenarios:=pScenarios
        End If
    End With

End Sub

Sub CodeFADIContextButton()
    CodeGLDIShowRegions 2
End Sub

Sub CodeFADIHeaderButton()
    CodeGLDIShowRegions 5
End Sub

Sub CodeGLDIContextButton()
    CodeGLDIShowRegions 2
End Sub

Sub CodeGLDIHeaderButton()
    CodeGLDIShowRegions 7
End Sub

Sub CodeGLDIShowRegions(myRow As Long)
Dim pDrawingObjects As Boolean
Dim pContents As Boolean
Dim pScenarios As Boolean

    On Error Resume Next
    Application.ScreenUpdating = False
    With ActiveSheet
        pDrawingObjects = .ProtectDrawingObjects
        pContents = .ProtectContents
        pScenarios = .ProtectScenarios
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Unprotect
        End If
        .Cells(myRow, 1).EntireRow.ShowDetail = Not .Cells(myRow, 1).EntireRow.ShowDetail
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Protect DrawingObjects:=pDrawingObjects, contents:=pContents, Scenarios:=pScenarios
        End If
    End With

End Sub

Sub CodeGLDIDescriptionsButton()
Dim pDrawingObjects As Boolean
Dim pContents As Boolean
Dim pScenarios As Boolean
Dim NoOfFFSegments As Integer
Dim SheetNumber As Integer
Dim TheBook As Object
Dim CS As Object
    On Error Resume Next
    Application.ScreenUpdating = False
    With ActiveSheet
        pDrawingObjects = .ProtectDrawingObjects
        pContents = .ProtectContents
        pScenarios = .ProtectScenarios
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Unprotect
        End If
        SheetNumber = .Cells(1, 1).Value
        Set TheBook = ActiveWorkbook
        Set CS = TheBook.Sheets("Criteria" & SheetNumber)
        NoOfFFSegments = Application.CountA(CS.Range(CS.Range("FFSegment1_" & SheetNumber), CS.Range("FFSegment1_" & SheetNumber).Offset(30, 0)))
        .Cells(2, 1 + NoOfFFSegments).EntireColumn.ShowDetail = Not .Cells(2, 1 + NoOfFFSegments).EntireColumn.ShowDetail
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Protect DrawingObjects:=pDrawingObjects, contents:=pContents, Scenarios:=pScenarios
        End If
    End With
End Sub

Sub CodeGLDIFilterButton()
Dim pDrawingObjects As Boolean
Dim pContents As Boolean
Dim pScenarios As Boolean
Dim TopRow As Long
Dim TopColumn As Long
    On Error Resume Next
    Application.ScreenUpdating = False
    With ActiveSheet
        pDrawingObjects = .ProtectDrawingObjects
        pContents = .ProtectContents
        pScenarios = .ProtectScenarios
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Unprotect
        End If
        TopRow = ActiveWindow.VisibleRange.Row
        TopColumn = ActiveWindow.VisibleRange.Column
        .Cells(9, 2).AutoFilter
        ActiveWindow.ScrollRow = TopRow
        ActiveWindow.ScrollColumn = TopColumn
        If .AutoFilterMode = False Then
            .Protect DrawingObjects:=True, contents:=True, Scenarios:=True
        End If
    End With
End Sub

Sub CodeGLDIResetFilterButton()
Dim pDrawingObjects As Boolean
Dim pContents As Boolean
Dim pScenarios As Boolean
    On Error Resume Next
    Application.ScreenUpdating = False
    With ActiveSheet
        pDrawingObjects = .ProtectDrawingObjects
        pContents = .ProtectContents
        pScenarios = .ProtectScenarios
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Unprotect
        End If
        .ShowAllData
        If Not pDrawingObjects And Not pContents And Not pScenarios Then
        Else
            .Protect DrawingObjects:=pDrawingObjects, contents:=pContents, Scenarios:=pScenarios
        End If
    End With
End Sub

Sub RWOutputTemplateDoubleClick()
    On Error Resume Next
    Application.Run "GLDIRWOutputTemplateDoubleClick"
End Sub

Sub RWAddBackgroundPictureRefreshButton(LayoutSheetName As Variant, MaskCellAddress As Variant, MaskCellAddress2 As Variant, MaskCellAddress3 As Variant, MaskCellAddress4 As Variant)
Dim NewButton As Object
Dim ButtonName As Variant
Dim LayoutSheet As Object
Dim MaskCell As Object
Dim tmpUpdating As Boolean

    On Error Resume Next

    tmpUpdating = Application.ScreenUpdating
    If tmpUpdating = True Then
        Application.ScreenUpdating = False
    End If
    
    Set LayoutSheet = ActiveWorkbook.Worksheets(LayoutSheetName)
    Set MaskCell = LayoutSheet.Range(MaskCellAddress)
    
    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("RWFindBackgroundPictureButton").Name
    
    If ButtonName = "" Then
        Set NewButton = LayoutSheet.Buttons.Add(MaskCell.Left, MaskCell.Top, 58.5, 14)
        NewButton.Left = NewButton.Left + 3
        NewButton.Top = NewButton.Top - 1
        NewButton.OnAction = "CodeRWFindBackgroundPicture"
        NewButton.Characters.Text = NLSGetString(369, "Find", 3135)
        NewButton.Name = "RWFindBackgroundPictureButton"
        CodeFormatButton NewButton, xlMove
        Set NewButton = Nothing
    End If
    
    Set MaskCell = LayoutSheet.Range(MaskCellAddress2)
    
    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("RWFindHTMLLocation").Name
    
    If ButtonName = "" Then
        Set NewButton = LayoutSheet.Buttons.Add(MaskCell.Left, MaskCell.Top, 58.5, 14)
        NewButton.Left = NewButton.Left + 3
        NewButton.Top = NewButton.Top
        NewButton.OnAction = "CodeRWFindHTMLLocation"
        NewButton.Characters.Text = NLSGetString(369, "Find", 3135)
        NewButton.Name = "RWFindHTMLLocation"
        CodeFormatButton NewButton, xlMove
        Set NewButton = Nothing
    End If

    If MaskCellAddress3 <> "" Then
        Set MaskCell = LayoutSheet.Range(MaskCellAddress4)
        
        ButtonName = ""
        ButtonName = LayoutSheet.DrawingObjects("RWMoveLineItemLeft").Name
        
        If ButtonName = "" Then
            Set NewButton = LayoutSheet.Buttons.Add(MaskCell.Left + 2, MaskCell.Top, 39, 14)
            NewButton.Left = NewButton.Left + 3
            NewButton.Top = NewButton.Top
            NewButton.OnAction = "CodeRWMoveLineItemLeft"
            NewButton.Characters.Text = "<-"
            NewButton.Name = "RWMoveLineItemLeft"
            CodeFormatButton NewButton, xlFreeFloating
            Set NewButton = Nothing
        End If
    End If
    
    If MaskCellAddress4 <> "" Then
        
        ButtonName = ""
        ButtonName = LayoutSheet.DrawingObjects("RWMoveLineItemRight").Name
        
        If ButtonName = "" Then
            Set NewButton = LayoutSheet.Buttons.Add(MaskCell.Left + 43, MaskCell.Top, 39, 14)
            NewButton.Left = NewButton.Left + 3
            NewButton.Top = NewButton.Top
            NewButton.OnAction = "CodeRWMoveLineItemRight"
            NewButton.Characters.Text = "->"
            NewButton.Name = "RWMoveLineItemRight"
            CodeFormatButton NewButton, xlFreeFloating
            Set NewButton = Nothing
        End If

        ButtonName = ""
        ButtonName = LayoutSheet.DrawingObjects("Dummy").Name
        
        If ButtonName = "" Then
            Set NewButton = LayoutSheet.Buttons.Add(MaskCell.Left + 2, MaskCell.Top - 32, 80, 30)
            NewButton.Left = NewButton.Left + 3
            NewButton.Top = NewButton.Top
            NewButton.OnAction = "Dummy"
            NewButton.Characters.Text = NLSGetString(4864, "Move Line Items", 6661)
            NewButton.Name = "Dummy"
            CodeFormatButton NewButton, xlFreeFloating
            Set NewButton = Nothing
        End If
        
    End If
    
    LayoutSheet.Rows("9:9").RowHeight = 16

    If Application.ScreenUpdating <> tmpUpdating Then
        Application.ScreenUpdating = tmpUpdating
    End If

End Sub

Sub Dummy()
End Sub

Sub CodeRWFindBackgroundPicture()
Dim Image As String
Dim IsThere As Integer
Dim Buffer As OFSTRUCT
Dim ComDlg As Object
Dim Filename As String

    On Error Resume Next

    Image = Application.ActiveSheet.Range("BackgroundPicture").Value

    If InStr(Image, "ORACLE_HOME") Then
        Image = TextReplace(Image, "ORACLE_HOME", CStr(gOracleHome))
    End If

    If Image = "" Then
        Image = gApplicationPathName & "\" & gFSGImagePath
    End If
                
    Set ComDlg = CreateObject("oraGLDI32.ComDlg")
    Filename = ComDlg.FileOpen(InitDir:=Image, Filter:=NLSGetString(3062, "Picture Files (*.bmp, *.gif, *.jpg)|*.bmp;*.gif;*.jpg|All Files(*.*)|*.*", 3137))
    
    Set ComDlg = Nothing

    AppActivate Application.Caption
    
    If Filename <> "" Then
                    
        IsThere = OpenFile(Filename, Buffer, OF_EXIST)
        If IsThere = -1 Then
            MsgBox NLSGetString(2700, "File Not Found.", 3138)
        Else
            Application.DisplayAlerts = False
            ActiveSheet.Range("BackgroundPicture").Value = Filename
            ActiveSheet.SetBackgroundPicture Filename
            Application.DisplayAlerts = True
        End If

    End If

End Sub

Sub CodeRWFindHTMLLocation()
Dim HTMLFile As String
Dim IsThere As Integer
Dim Buffer As OFSTRUCT
Dim ComDlg As Object
Dim Filename As String

    On Error Resume Next

    HTMLFile = Application.ActiveSheet.Range("HTMLFileLocation").Value

    If HTMLFile = "" Then
        HTMLFile = gApplicationPathName
    End If
                
    Set ComDlg = CreateObject("oraGLDI32.ComDlg")
    Filename = ComDlg.FileSave(InitDir:=HTMLFile, Filter:=NLSGetString(3063, "HTML Files (*.htm, *.html)|*.htm;*.html|All Files(*.*)|*.*", 3139))
    Set ComDlg = Nothing

    AppActivate Application.Caption
        
    If Filename <> "" Then
        ActiveSheet.Range("HTMLFileLocation").Value = Filename
    End If

End Sub

Function TextReplace(ByVal FullString As String, SubStringToFind As String, SubStringReplace As String) As String
Dim TempString As String
Dim StartPosition As Integer
Dim EndPosition As Integer
Dim SearchStartPos As Integer
Dim OriginalString As String

    On Error Resume Next

    SearchStartPos = 1
    OriginalString = FullString

    While InStr(SearchStartPos, FullString, SubStringToFind) > 0

        StartPosition = InStr(FullString, SubStringToFind)
        EndPosition = StartPosition + Len(SubStringToFind)
        TempString = Left$(FullString, StartPosition - 1)
        TempString = TempString & SubStringReplace
        
        SearchStartPos = Len(TempString) + 1
        If SearchStartPos < 1 Then
            SearchStartPos = 1
        End If
        
        TempString = TempString & Mid$(FullString, EndPosition)
        FullString = TempString
    
    Wend
    
    TextReplace = FullString
            
End Function

Function NLSGetString(lngID As Long, NlSString As String, UsageID As Long) As String
    On Error Resume Next
    NLSGetString = Application.Run(AddinFileName() & "!NLSGetString", lngID, NlSString, UsageID)
End Function

Function NLSGetReservedString(lngID As Long, NlSString As String) As String
    On Error Resume Next
    NLSGetReservedString = Application.Run(AddinFileName() & "!NLSGetReservedString", lngID, NlSString)
End Function

Function AddinFileName() As String
    On Error Resume Next
    If CDbl(Left(Application.Version, 1)) < 8 Then
        AddinFileName = "GLDI90.XLA"
    Else
        AddinFileName = "GLDI9097.XLA"
    End If
End Function

Sub FADIAddOutlineButtons(LayoutSheetName As String, HeaderButton As Boolean)
Dim NewButton As Object
Dim ButtonName As Variant
Dim LayoutSheet As Object
Dim MaskCell As Object
Dim ContextButton As Object

    On Error Resume Next

    If Application.Version < 8 Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set LayoutSheet = ActiveWorkbook.Worksheets(LayoutSheetName)
    Set MaskCell = LayoutSheet.Cells(1, 2)

    MaskCell.Offset(1, 0).EntireRow.RowHeight = 18

    ButtonName = ""
    ButtonName = LayoutSheet.DrawingObjects("ButtonContext").Name

    If ButtonName = "" Then
        Set NewButton = LayoutSheet.Buttons.Add((MaskCell.Left / 2.5), MaskCell.Top + 3, 98, 14)
        Set ContextButton = NewButton
        NewButton.OnAction = "CodeFADIContextButton"
        NewButton.Characters.Text = NLSGetString(4611, "Show Context", 6248)
        NewButton.Name = "ButtonContext"
        CodeFormatButton NewButton, xlFreeFloating
        Set NewButton = Nothing
    Else
        Set ContextButton = LayoutSheet.DrawingObjects("ButtonContext")
    End If

    ButtonName = ""

    If HeaderButton Then
        ButtonName = LayoutSheet.DrawingObjects("ButtonHeader").Name
        Set NewButton = LayoutSheet.Buttons.Add((ContextButton.Left) + (ContextButton.Width), MaskCell.Top + 3, 98, 14)
        NewButton.OnAction = "CodeFADIHeaderButton"
        NewButton.Characters.Text = NLSGetString(4612, "Show Header", 6249)
        NewButton.Name = "ButtonHeader"
        CodeFormatButton NewButton, xlFreeFloating
        Set NewButton = Nothing
        ButtonName = ""
    End If

End Sub

Public Sub CodeFormatButton(myButton As Object, bPlacement As Variant)

    On Error Resume Next

    With myButton.Font
        .Name = NLSGetReservedString(61, "Arial")
        .FontStyle = "Bold"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlNone
        .ColorIndex = xlAutomatic
    End With
    With myButton
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = xlHorizontal
        .AutoSize = False
        .Placement = bPlacement
        .PrintObject = False
    End With

End Sub

Sub MoveLineItem(Direction As String)

Dim StartColumn As Integer
Dim EndColumn As Integer
Dim StartRow As Long
Dim EndRow As Long
Dim BlankColumn As Integer
Dim TempWidth As Variant
Dim CurrScrollCol As Integer
Dim CurrScrollRow As Long
Dim CurrCol As Integer
Dim CurrRow As Long

On Error Resume Next

    With Application
        .ScreenUpdating = False
        CurrScrollCol = .ActiveWindow.ScrollColumn
        CurrScrollRow = .ActiveWindow.ScrollRow
        CurrCol = .ActiveCell.Column
        CurrRow = .ActiveCell.Row
    End With
    
    StartColumn = Range("LINEITEMS").Column
    If Direction = mcRight Then
        EndColumn = StartColumn + 1
    Else
        EndColumn = StartColumn - 1
    End If
    
    StartRow = Range("COLUMNHEADINGS").Row - 1
    EndRow = Range("BACKGROUNDPICTURE").Row - 2
    
    Range("COLUMNHEADINGS").CurrentRegion.Select
    BlankColumn = Selection.Columns(Selection.Columns.Count).Column + 1
    
    If ((Direction = mcRight) And (EndColumn < BlankColumn)) Or ((Direction = mcLeft) And (EndColumn >= 1)) Then
        
        With ActiveSheet
            Range(Cells(StartRow, StartColumn), Cells(EndRow, StartColumn)).Cut
        
            Cells(StartRow, BlankColumn).Select
            .Paste
        
            Range(Cells(StartRow, EndColumn), Cells(EndRow, EndColumn)).Cut
        
            Cells(StartRow, StartColumn).Select
            .Paste
        
            Range(Cells(StartRow, BlankColumn), Cells(EndRow, BlankColumn)).Cut
        
            Cells(StartRow, EndColumn).Select
            .Paste
        
            TempWidth = Columns(StartColumn).ColumnWidth
            Columns(StartColumn).ColumnWidth = Columns(EndColumn).ColumnWidth
            Columns(EndColumn).ColumnWidth = TempWidth
        
            Cells(6, 2).Select
        End With
    End If

    With Application
        .ActiveWindow.ScrollColumn = CurrScrollCol
        .ActiveWindow.ScrollRow = CurrScrollRow
        .Cells(CurrRow, CurrCol).Select
        .ScreenUpdating = True
    End With
    
End Sub

Sub CodeRWMoveLineItemLeft()

    MoveLineItem mcLeft

End Sub

Sub CodeRWMoveLineItemRight()

    MoveLineItem mcRight

End Sub


