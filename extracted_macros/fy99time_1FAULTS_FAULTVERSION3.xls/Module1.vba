Attribute VB_Name = "Module1"









Const strToolbar = "Standard"
Const szDataShtName = "TimelineWizardData"
Const szCommandLine = "/excel"
Const szRunWizrdErr = "Unable to locate Create Timeline Wizard.  Please run wizard from Visio."
Const strButtonName = "Visio Timeline Wizard"    ' The button name is also used as the ToolTip.
Const iButtonID = 231                       ' ID of blank button

Dim lRetVal As Long
'I stole this from the Org chart but will not now change sub names

Sub Auto_Open()
Attribute Auto_Open.VB_ProcData.VB_Invoke_Func = " \n14"
    'Add Org Wizard Button

    AddOrgWizardButton
    
End Sub

Sub Auto_Close()
Attribute Auto_Close.VB_ProcData.VB_Invoke_Func = " \n14"
    'Remove Org Wizard Button
    
    Call RemoveOrgWizardButton

End Sub

Private Sub AddOrgWizardButton()
' Add the "InsertVisioDrawing" button to Excel's standard toolbar
' if the button does not already exist.

    Set btns = Toolbars(strToolbar).ToolbarButtons
    Set btn = ButtonsIndex(btns, strButtonName)
    
    ' Check if toolbar button already exists
    If Not (btn Is Nothing) Then
        btn.Delete
    End If
    
    ' Add a blank button to the Standard toolbar,
    ' after the "Drawing" toolbar button.
    iLoc = ButtonsLoc(btns, "Drawing")
    If iLoc = 0 Then
        Set btn = btns.Add(iButtonID)
    Else
        Set btn = btns.Add(iButtonID, iLoc + 1)
    End If
    btn.Name = strButtonName
    
    ' Copy the button bitmap to the clipboard.
    ' Paste it onto the button.
    Set objWorkbook = Application.ThisWorkbook
    objWorkbook.Sheets(szDataShtName).DrawingObjects(2).CopyPicture
    btn.PasteFace
                
    ' Set the macro the toolbar button will run.
    btn.OnAction = "RunOrgChartWizard"
End Sub

Private Sub RemoveOrgWizardButton()

    Set btns = Toolbars(strToolbar).ToolbarButtons
    Set btn = ButtonsIndex(btns, strButtonName)
    
    If Not (btn Is Nothing) Then
        btn.Delete
    End If
        
End Sub

Sub RunOrgChartWizard()
Attribute RunOrgChartWizard.VB_Description = "Creates a Visio Timeline from entered data."
Attribute RunOrgChartWizard.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim szOrgWizExe As String
    
    On Error GoTo ErrRunWizard
    
    szOrgWizExe = Application.ThisWorkbook.Worksheets(szDataShtName).Cells(1).Formula
    
    lRetVal = Shell(szOrgWizExe & " " & szCommandLine, 5)
        
Exit Sub
        
ErrRunWizard:
    MsgBox szRunWizrdErr, 48
End Sub


Private Function ButtonsIndex(ByVal Buttons As Object, ByVal bname As String) As Object
' Index any collection by name.
' Returns the object with a given name.
' Returns Nothing if not found.

    For Each btn In Buttons
        If btn.Name = bname Then
            Set ButtonsIndex = btn
            Exit For
        End If
    Next
End Function


Private Function ButtonsLoc(ByVal Buttons As Object, ByVal bname As String) As Integer
' Returns the location of a button with a given name
' or zero if not found.

    n = Buttons.Count
    For i = 1 To n
        If Buttons(i).Name = bname Then
            ButtonsLoc = i
            Exit For
        End If
    Next
End Function
