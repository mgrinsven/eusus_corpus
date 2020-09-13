Attribute VB_Name = "Sheet9"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "ScrollBar1, 1, 0, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar2, 2, 1, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar3, 3, 2, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar4, 4, 3, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar5, 5, 4, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar6, 6, 5, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar7, 7, 6, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar8, 8, 7, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar9, 9, 8, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar10, 10, 9, MSForms, ScrollBar"
Attribute VB_Control = "ScrollBar11, 11, 10, MSForms, ScrollBar"
Attribute VB_Control = "ButtonReset, 12, 11, MSForms, CommandButton"
Attribute VB_Control = "ButtonRun, 13, 12, MSForms, CommandButton"
Attribute VB_Control = "ScrollBar12, 14, 13, MSForms, ScrollBar"
Private Sub ButtonReset_Click()
    SimInit
    'force display of input page
    Sheets("Game").Select
    Range("RLV_Repair_System_Constraints").Select
End Sub

Private Sub ButtonRun_Click()
    SimAnimating = False
    SimContinue
    'force display of input page
    Sheets("Game").Select
    Range("RLV_Repair_System_Constraints").Select
End Sub
