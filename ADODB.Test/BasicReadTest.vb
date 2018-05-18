Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' from https://www.w3schools.com/asp/ado_recordset.asp
''' </summary>
<TestClass()> Public Class BasicReadTest


    <TestMethod()> Public Sub CreateReadContents()
        '    Set conn=Server.CreateObject("ADODB.Connection")
        'conn.Provider="Microsoft.Jet.OLEDB.4.0"
        'conn.Open "c:/webdata/northwind.mdb"
        'Set rs=Server.CreateObject("ADODB.recordset")
        'rs.Open "Select * from Customers", conn


        Dim conn As New ADODB.Connection
        conn.Open(Commun.StringConnection)

        Dim rs As New ADODB.Recordset
        rs.Open("Select id,Title,ReleaseDate from Movies", conn)

        Dim index As Byte = 0

        For Each x In rs.fields
            Assert.IsNotNull(x)

            Select Case index
                Case 0
                    Assert.AreEqual(x.Name, "id")
                    Assert.IsInstanceOfType(x.Value, GetType(Integer))
                Case 1
                    Assert.AreEqual(x.Name, "Title")
                    Assert.IsInstanceOfType(x.Value, GetType(String))
                Case 2
                    Assert.AreEqual(x.Name, "ReleaseDate")
                    Assert.IsInstanceOfType(x.Value, GetType(DateTime))
                Case Else
                    Exit Select
            End Select

            index += 1
        Next

    End Sub

    <TestMethod()> Public Sub CreateReadAllContents()
        '    Set conn=Server.CreateObject("ADODB.Connection")
        'conn.Provider="Microsoft.Jet.OLEDB.4.0"
        'conn.Open "c:/webdata/northwind.mdb"
        'Set rs=Server.CreateObject("ADODB.recordset")
        'rs.Open "Select * from Customers", conn

        Dim conn As New ADODB.Connection
        conn.Open(Commun.StringConnection)

        Dim rs As New ADODB.Recordset
        rs.Open("Select id,Title from Movies", conn)

        Dim count As Byte = 0

        Do Until rs.EOF = True

            Assert.AreEqual(rs.fields(0).Name, "id")
            Assert.IsInstanceOfType(rs.fields(0).Value, GetType(Integer))

            Assert.AreEqual(rs.fields(1).Name, "Title")
            Assert.IsInstanceOfType(rs.fields(1).Value, GetType(String))

            rs.MoveNext()
            count += 1

        Loop

        Assert.AreEqual(Of Byte)(count, 3)
        Assert.IsTrue(rs.EOF)

    End Sub
End Class