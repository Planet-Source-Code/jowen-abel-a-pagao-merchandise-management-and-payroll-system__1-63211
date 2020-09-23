Attribute VB_Name = "ModDB"
Public ac As New ADODB.Connection
Public ar As New ADODB.Recordset


Public CurrentForm As Form
Public strConek, pword, CurrentUser As String
Public rc, ctr, passFlag, liCtr, dbFlag, menuFlag, saveFlag As Integer


Public Function dbconek()
    Set ac = New ADODB.Connection
    Set ar = New ADODB.Recordset
    strConek = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbase.mdb;Persist Security Info=False"
    
End Function
Public Function dbconek1()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    strConek = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbase.mdb;Persist Security Info=False"
    
End Function

Public Function dbclose()
On Error Resume Next
    ac.Close
    Set ac = Nothing
    ar.Close
    Set ar = Nothing
End Function
Public Function offDefine(Key_Ascii As Integer, ByVal ControlName As Object, sFilter As String)

If InStr(sFilter, Chr(Key_Ascii)) = 0 Then
    Key_Ascii = 0
End If

End Function

