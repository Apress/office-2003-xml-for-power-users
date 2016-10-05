-----------------
About the Samples
-----------------

This archive includes the sample content for the book
"Office 2003 XML for Power Users."

For other related examples, and to download the most recent version of this
file (which may include corrections or additional examples in the future),
visit http://www.prosetech.com.

NOTE: It's recommended that you unzip this archive to C:\.
This will create a C:\OfficeXML directory, with subdirectories
for each chapter. You can use a different directory structure if you want,
but using C:\OfficeXML will make it easier to use the InfoPath examples,
and to follow the steps for creating the virtual directory in Chapter 8.


-----------
Sample List
-----------

The following list identifies the examples in each directory.
Detailed setup instructions for the .NET Framework and the development
tools used in Chapter 8 are provided in the book.

---------------------------------------------------------------------------

\Chapter01\
* Contains sample XML documents (.xml files) with and without namespaces.
* Contains the test application GetXMLNamespaces which identifies all the
namespaces in an XML document. This application requires the .NET Framework.

\Chapter01\Code\GetXMLNamespaces\
* Contains the VB .NET source code for the GetXMLNamespaces application.
To edit this code, you'll need Visual Studio .NET.

--------------

\Chapter02\
* Contains sample XML documents (.xml files) and matching schemas (.xsd files).
* Contains the test application ValidateXML which checks if an XML document
adheres to the rules of the schema. This application requires the .NET Framework.

\Chapter02\Namespaces\
* Contains examples of XML documents that use namespaces, with the matching
schemas.

\Chapter02\LinkedSchemas\
* Contains examples of XML documents that link directly to schema documents
(using the schemaLocation or noNamespaceSchemaLocation attributes).

\Chapter02\Code\ValidateXML\
* Contains the VB .NET source code for the ValidateXML application.
To edit this code, you'll need Visual Studio .NET.

--------------

\Chapter03\Lists\
* Contains sample Excel spreadsheets that show several different ways of
mapping XML documents with product list information.

\Chapter03\AnalyzingXML\
* Contains a mapped product expense spreadsheet with basic analysis
(calculated columns and a chart). You can use this example with three
different XML files.

\Chapter03\TemplateRetrofit\
* Contains a mapped expense report spreadsheet.
* Contains two sample applications for reading the exported XML.
VB6_Reader uses Visual Basic 6 code, while VB.NET_Reader performs the same
task using VB .NET code.

\Chapter03\TemplateRetrofit\Code\VB 6\
* Contains the source code for the VB6_Reader application.
To edit this code, you'll need Visual Basic 6.

\Chapter03\TemplateRetrofit\Code\VB .NET\
* Contains the source code for the VB.NET_Reader application.
To edit this code, you'll need Visual Studio .NET.

--------------

\Chapter04\
* Contains sample Word documents that show several different ways of
mapping XML documents.

\Chapter04\Memo\
* Contains the memo sample, which uses a mapped Word document as the basis
of a template.

\Chapter04\TwoNamespaces\
* Contains an example of a mapped Word document that combines two
namespaces, from different schemas.

--------------

\Chapter05\
* COntains the sample Northwind Access database.

--------------

\Chapter06\
* Contains sample SpreadsheetML and WordML documents.
* Contains the CreateOfficeDoc test application, which generates a new WordML
document based on a template.
* Contains the GetOfficeProperties test application, which reads and displays
information from SpreadsheetML and WordML documents.

\Chapter06\Code\VB 6\
* Contains the source code for the CreateOfficeDoc and GetOfficeProperties
applications. To edit this code, you'll need Visual Basic 6.

\Chapter06\Code\VB .NET\
* Contains the source code for an alternate version of the CreateOfficeDoc
and GetOfficeProperties applications written in VB .NET. To edit this code,
you'll need Visual Studio .NET.

--------------

\Chapter07\
* Contains XML documents (.xml files) and XSLT stylesheets (.xslt files) for
a range of different conversions.
* Contains the TransformXML application that allows you to perform an XSLT
transformation on an XML document. This application requires the .NET Framework.

\Chapter07\Memo\
* Contains the Word memo example, with XSLT stylesheets that convert data-only
XML to formatted WordML.

\Chapter07\OfficeXSLT\
* Contains sample XLST stylesheets that extract information from WordML or
SpreadsheetML files, and use it to create rich web pages. Also includes sample
WordML documents that are configured to open up in a browser and apply these
transformations automatically.

\Chapter07\Code\TransformXML\
* Contains the VB .NET source code for the TransformXML application.
To edit this code, you'll need Visual Studio .NET.

--------------

\Chapter08\ExpenseReport\
* Contains the files required for the ExpenseReport workflow solution. These
include a web service and web page written in VB .NET, an Access database, and
a expense spreadsheet with macro code. These files are described in detail in the
book.

--------------

\Chapter09\
* Contains templates (.xsn files) that you can use with InfoPath.
If you change the directory path, you will not be able to open these forms
for editing. However, you can always open them in design mode, as described
in the book.

---------------------------------------------------------------------------