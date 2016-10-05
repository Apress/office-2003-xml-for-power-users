<%@ WebService language="VB" class="ReportService" %>

Imports System
Imports System.Web.Services
Imports System.Xml.Serialization
Imports System.Data
Imports System.Data.OleDb

<WebService([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport")> Public Class ReportService

   <WebMethod> Public Function SubmitReport(Root as Root) As String

        Dim ConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\ExpenseReport\Expenses.mdb"
        Dim Connection As New OleDbConnection(ConnectionString)
        Connection.Open()

        ' Create the ExpenseReport record.
        Dim SQL As String = "INSERT INTO ExpenseReport (Name, Email, IDNumber, Purpose) " & _
          "VALUES ('" & Root.Meta.Name & "','" & Root.Meta.Email & "','" & Root.Meta.IDNumber & "','" & Root.Meta.Purpose & "')"
        Dim Command As New OleDbCommand(SQL, Connection)
        Command.ExecuteNonQuery()

        ' Create the linked ExpenseItem records.
        Dim Item As RootExpenseItem
        For Each Item In Root.ExpenseItem
            SQL = "INSERT INTO ExpenseItem " & _
              "([Date], Description, Miles, Rate, AirFare, Other, Meals, " & _
              "Conference, Misc, MiscCode, Total, ExpenseReportID)" & _
              "VALUES ('" & Item.Date & "','" & Item.Description & "'," & _
              Item.Miles & "," & Item.Rate & "," & Item.AirFare & "," & _
              Item.Other & "," & Item.Meals & "," & Item.Conference & "," & _
              Item.Misc & ",'" & Item.MiscCode & "'," & Item.Total & ",'" & _
              Root.Meta.IDNumber & "')"

            Command.CommandText = SQL
            Command.ExecuteNonQuery()
        Next

        Connection.Close()

        Return Root.Meta.Name & " was successfully added to the database."
   End Function

End Class



'<remarks/>
<System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport"),  _
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport", IsNullable:=false)>  _
Public Class Root

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Meta As RootMeta

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute("ExpenseItem")>  _
    Public ExpenseItem() As RootExpenseItem
End Class

'<remarks/>
<System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport")> _
Public Class RootMeta

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Name As String

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Email As String

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public IDNumber As String

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Purpose As String
End Class

'<remarks/>
<System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport")>  _
Public Class RootExpenseItem

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()> _
    Public [Date] As String 'Date

'DataType:="date")>  _
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Description As String

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(DataType:="integer")>  _
    Public Miles As String

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Rate As Decimal

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public AirFare As Decimal

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Other As Decimal

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Meals As Decimal

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Conference As Decimal

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Misc As Decimal

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public MiscCode As String

    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute()>  _
    Public Total As Decimal
End Class
