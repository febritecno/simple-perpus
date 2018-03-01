Attribute VB_Name = "Module1"
Public Conn As New ADODB.Connection
Public RS As New ADODB.Recordset
Public RSpinjam As ADODB.Recordset
Public RSbuku As ADODB.Recordset
Public RSpengunjung As ADODB.Recordset
Public RSlogin As ADODB.Recordset

Public Sub sipp()
Set Conn = New ADODB.Connection
Set RSbuku = New ADODB.Recordset
Set RSpengunjung = New ADODB.Recordset
Set RSpinjam = New ADODB.Recordset
Set RSlogin = New ADODB.Recordset
Set RS = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\perpustakaan.mdb"
End Sub

