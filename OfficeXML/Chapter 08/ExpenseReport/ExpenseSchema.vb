﻿'------------------------------------------------------------------------------
' <autogenerated>
'     This code was generated by a tool.
'     Runtime Version: 1.1.4322.573
'
'     Changes to this file may cause incorrect behavior and will be lost if 
'     the code is regenerated.
' </autogenerated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System.Xml.Serialization

'
'This source code was auto-generated by xsd, Version=1.1.4322.573.
'

'<remarks/>
<System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport"),  _
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport", IsNullable:=false)>  _
Public Class Root
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Meta As RootMeta
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute("ExpenseItem", Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public ExpenseItem() As RootExpenseItem
End Class

'<remarks/>
<System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport")>  _
Public Class RootMeta
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Name As String
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Email As String
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public IDNumber As String
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Purpose As String
End Class

'<remarks/>
<System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.prosetech.com/Schemas/ExpenseReport")>  _
Public Class RootExpenseItem
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType:="date")>  _
    Public [Date] As Date
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Description As String
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType:="integer")>  _
    Public Miles As String
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Rate As Decimal
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public AirFare As Decimal
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Other As Decimal
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Meals As Decimal
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Conference As Decimal
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Misc As Decimal
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public MiscCode As String
    
    '<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)>  _
    Public Total As Decimal
End Class
