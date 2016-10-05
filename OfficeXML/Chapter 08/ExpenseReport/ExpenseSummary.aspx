<%@ Page Language="VB" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.OleDb" %>
<script runat="server">

    ' Insert page code here
    '
    
    Sub Page_Load(sender As Object, e As EventArgs)
        Dim ConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\ExpenseReport\Expenses.mdb"
        Dim Conection As New OleDbConnection(ConnectionString)
        Dim SQL As String = "SELECT * FROM ExpenseReport"
        Dim Adapter As New OleDbDataAdapter(SQL, Connection)
        Dim DS As New DataSet("Expenses")
        Adapter.Fill(DS, "ExpenseReport")
        Adapter.SelectCommand.CommandText = "SELECT * FROM ExpenseItem"
        Adapter.Fill(DS, "ExpenseItem")
    
        Dim Relation As New DataRelation("Records_Items", _
                DS.Tables("ExpenseReport").Columns("IDNumber"), _
                DS.Tables("ExpenseItem").Columns("ExpenseReportID"))
    
        Relation.Nested = True
        DS.Relations.Add(Relation)
    
        Xml1.DocumentContent = DS.GetXml()
        Xml1.TransformSource = "ExpenseSummary.xslt"
    End Sub

</script>
<html>
<head>
</head>
<body>
    <form runat="server">
        <asp:Xml id="Xml1" runat="server"></asp:Xml>
        <!-- Insert content here -->
    </form>
</body>
</html>
