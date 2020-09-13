Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "ButtonInitialize, 1, 0, MSForms, CommandButton"
Attribute VB_Control = "ButtonSimulate, 2, 1, MSForms, CommandButton"
Private Sub ButtonInitialize_Click()
    SimInit
End Sub

Private Sub ButtonSimulate_Click()
    SimAnimating = True
    SimContinue
End Sub

