Attribute VB_Name = "Module1"
Public ConN As New ADODB.Connection
Public RsAbsen As New ADODB.Recordset
Sub koneksi()
Set ConN = New ADODB.Connection
Set RsAbsen = New ADODB.Recordset
ConN.Open "Provider=microsoft.jet.oledb.4.0;data source = " & App.Path & "\latihan.mdb"
End Sub
