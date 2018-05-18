Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class BasicInstanceTest


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
        'This would be ignored
        'conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        'Required change for .NET and using of local database for now
        conn.Open(Commun.StringConnection)

        Assert.IsNotNull(conn.innerConnection)
        'Connection is not open for default
        Assert.AreEqual(conn.innerConnection.State, ConnectionState.Closed)

    End Sub

    <TestMethod()> Public Sub CreateConnectionAndRecordset()
        'set conn=Server.CreateObject("ADODB.Connection")
        'conn.Provider = "Microsoft.Jet.OLEDB.4.0"
        'conn.Open "c:/webdata/northwind.mdb"
        'set rs=Server.CreateObject("ADODB.recordset")
        'rs.Open "Select * from Customers", conn

        Dim conn As New ADODB.Connection
        conn.Open(Commun.StringConnection)

        Dim rs As New ADODB.Recordset
        rs.Open("Select * from Movies", conn)

        Assert.IsNotNull(rs.innerCommand)

    End Sub



End Class