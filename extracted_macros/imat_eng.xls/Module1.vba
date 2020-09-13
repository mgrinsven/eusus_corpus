Attribute VB_Name = "Module1"
Private Sub UserForm_Initialize()
    MultiPage1.Instructions.ControlTipText = "Here in page 1"
    MultiPage1.Collection& Calculation.ControlTipText = "Now in page 2"
    
    CommandButton1.ControlTipText = "And now here's"
    CommandButton2.ControlTipText = "a tip from"
    CommandButton3.ControlTipText = "your controls!"
End Sub
