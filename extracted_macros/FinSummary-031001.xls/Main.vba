Attribute VB_Name = "Main"
Option Explicit

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Type LU
    strCell As String
    strQuery As String
End Type

'This array holds cell/query combinations to fill.
'Note: The first element in this array is the sentinel. If the cell of its name
'contains data, we'll exit.
Public aLUs(1) As LU

Sub Auto_Open()
'
' AutoOpen Macro

' This is where the data to be written is specified:
aLUs(0).strCell = "B22"
aLUs(0).strQuery = "Profiler_MEGATASK600111_A1"

' Uncomment this to fill the document from the profile database
    FillDoc True
' Or uncomment this to fill the document from the current wizard database
'    FillDoc False
End Sub

Sub FillDoc(bProfile As Boolean)
' This does the actual work. If bProfile is true we get the info out of the profile database,
' otherwise we get it from the current wizard db

    On Error Resume Next

    Dim dbs As Database
    Dim dbPath As String
    Dim rs As Recordset
    
    If Worksheets(1).Range(aLUs(0).strCell).Value <> "" Then Exit Sub
    
    Dim i As Integer
    
    If bProfile = True Then
        dbPath = GetProfileDB()
    Else
        dbPath = GetCurrentDB()
    End If
    
    If dbPath = "" Then Exit Sub
    Set dbs = DBEngine.Workspaces(0).OpenDatabase(dbPath)
   
    For i = 0 To UBound(aLUs) - 1
        
        Err.Clear
        Set rs = dbs.OpenRecordset(aLUs(i).strQuery, dbOpenForwardOnly)
        If Err = 0 Then
            If Not (rs.EOF And rs.BOF) Then
                Worksheets(1).Range(aLUs(i).strCell).Value = CStr("" & rs(0))
            End If
        End If
        rs.Close
        
    Next i

    dbs.Close
End Sub

Function GetCurrentDB() As String
    GetCurrentDB = GetOptionsKey("CurrentWizardDB")
End Function

Function GetProfileDB() As String
Dim strDir

    strDir = GetOptionsKey("CurrentProfileDir")
    If strDir <> "" Then
        If Right$(strDir, 1) <> "\" Then strDir = strDir & "\"
        strDir = strDir & "msbp_plz.mdb"
    End If
    GetProfileDB = strDir
End Function

Function GetOptionsKey(strValue As String) As String
Dim strKey, strData As String
Dim hKey As Long
Dim dwAction, dwDataSize, dwType As Long

Const HKEY_CURRENT_USER = &H80000001
Const REG_SZ = 1

    strKey = "Software\Microsoft\Microsoft Reference\SBB\9.0Z\Options"

    'This is a template moniker First, get the content root
    If RegCreateKey(HKEY_CURRENT_USER, strKey, hKey) = 0 Then

        strData = String(1024, 32)

        If RegQueryValueEx(hKey, strValue, 0, REG_SZ, ByVal strData, 1024) = 0 Then
            If InStr(strData, Chr$(0)) > 0 Then
                strData = Left$(strData, InStr(strData, Chr$(0)) - 1)
            End If
            strData = Trim(strData)
            GetOptionsKey = strData
        End If

        RegCloseKey (hKey)
    End If
End Function

