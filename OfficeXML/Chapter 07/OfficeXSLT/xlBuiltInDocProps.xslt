<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
	<xsl:template match="/">
		<html>
			<head>
				<STYLE type="text/css">
		          .tblTitle { background-color: #CCCCCC; font-weight: bold; border=0 }
		          .tblData { border-style:solid}
		        </STYLE>
			</head>
			<body>
				<xsl:apply-templates select="//ss:Workbook/o:DocumentProperties"/>
			</body>
		</html>
	</xsl:template>
	<xsl:template match="//ss:Workbook/o:DocumentProperties">
		<table border="1" width="100%">
			<xsl:if test="position()=1">
				<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
			</xsl:if>
			<tr>
				<td class="tblTitle" width="10%">Title</td>
				<td class="tblData" width="23%">
					<xsl:value-of select="o:Title"/>
				</td>
				<td class="tblTitle">Author</td>
				<td class="tblData">
					<xsl:value-of select="o:Author"/>
				</td>
				<td class="tblTitle" width="10%">Category</td>
				<td class="tblData" width="23%">
					<xsl:value-of select="o:Category"/>
				</td>
			</tr>
			<tr>
				<td class="tblTitle">Subject</td>
				<td class="tblData">
					<xsl:value-of select="o:Subject"/>
				</td>
				<td class="tblTitle">Manager</td>
				<td class="tblData">
					<xsl:value-of select="o:Manager"/>
				</td>
				<td class="tblTitle">Keywords</td>
				<td class="tblData">
					<xsl:value-of select="o:Keywords"/>
				</td>
			</tr>
			<tr>
				<td/>
				<td/>
				<td class="tblTitle" width="10%">Company</td>
				<td class="tblData" width="23%">
					<xsl:value-of select="o:Company"/>
				</td>
				<td class="tblTitle">Comments</td>
				<td class="tblData">
					<xsl:value-of select="o:Description"/>
				</td>
			</tr>
			<xsl:if test="position()=last()">
				<xsl:text disable-output-escaping="yes">&lt;/tbody&gt;</xsl:text>
			</xsl:if>
		</table>
	</xsl:template>
</xsl:stylesheet>
