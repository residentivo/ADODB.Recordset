Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' from https://www.w3schools.com/asp/ado_recordset.asp
''' </summary>
<TestClass()> Public Class BasicUnitTest

    <TestMethod()> Public Sub CreateInstanceRecordSet()
        'set objRecordset=Server.CreateObject("ADODB.recordset")

        Dim objRecordset As New ADODB.Recordset()

        Assert.IsNotNull(objRecordset)

    End Sub

    <TestMethod()> Public Sub CreateInstanceConnection()
        'set conn=Server.CreateObject("ADODB.Connection")
        'conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        'conn.Open "c:/webdata/northwind.mdb"

        Dim conn As New ADODB.Connection
        conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        'Required change for .NET
        conn.Open("c:/webdata/northwind.mdb")

        Assert.IsNotNull(conn)

    End Sub

    <TestMethod()> Public Sub CreateConnectionAndRecordset()
        'set conn=Server.CreateObject("ADODB.Connection")
        'conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        'conn.Open "c:/webdata/northwind.mdb"
        'set rs=Server.CreateObject("ADODB.recordset")
        'rs.Open "Select * from Customers", conn

    End Sub
    <TestMethod()> Public Sub CreateReadContents()
        '    Set conn=Server.CreateObject("ADODB.Connection")
        'conn.Provider="Microsoft.Jet.OLEDB.4.0"
        'conn.Open "c:/webdata/northwind.mdb"

        'Set rs=Server.CreateObject("ADODB.recordset")
        'rs.Open "Select * from Customers", conn

        'For Each x In rs.fields
        '  response.write(x.name)
        '  response.write(" = ")
        '  response.write(x.value)
        'Next

    End Sub

End Class