Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public Function connect()
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=liba;Data Source=(local)"
End Function

