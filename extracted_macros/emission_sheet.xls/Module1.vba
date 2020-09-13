Attribute VB_Name = "Module1"
Function noerror(X, Optional ERRMSG = "")
'returns nothing if x is an error calculation, like divsion by zero.
'useful for ignoring errors
'errmsg is an optional error message to show when there is an error
If IsError(X) Then
    noerror = ERRMSG
     Else
    noerror = X
End If
End Function

Function NoNull(X, Optional xMSG = "")
If X = 0 Then
    NoNull = xMSG
     Else
    NoNull = X
End If
End Function
