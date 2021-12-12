Attribute VB_Name = "Module1"
Option Explicit
Public Con As New Connection
Public rs As New Recordset

Public Sub Connect()

 If Con.State = 1 Then
  Con.Close
 End If
 Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COSMETA.mdb;Persist Security Info=False"
 Con.Open
 'MsgBox "Connection Successfull"
End Sub

Public Sub retData(Q As String)
If rs.State = 1 Then
  rs.Close
  rs.CursorLocation = adUseClient
 End If
 
 rs.CursorLocation = adUseClient
 rs.Open Q, Con, adOpenStatic, adLockOptimistic
End Sub
