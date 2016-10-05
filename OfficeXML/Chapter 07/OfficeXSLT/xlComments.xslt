<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
	<xsl:template match="/">
		<html>
			<head>
				<STYLE type="text/css">
		          .tblHead { background-color: DarkGray; font-weight: bold; border=0}
		          .tblTitle { background-color:LightGrey; font-weight: bold; border=0 }
		          .tblData { border-style:solid}
		        </STYLE>
			</head>
			<body>
				<xsl:if test="count(//ss:Comment)=0">
					<xsl:text>This workbook does not contain any comments.</xsl:text>
				</xsl:if>
				<xsl:for-each select="//ss:Comment">
					<xsl:if test="position()=1">
						<xsl:text disable-output-escaping="yes">&lt;table border="1"width="100%"&gt;</xsl:text>
					</xsl:if>
					<xsl:if test="position()=1">
						<tr>
							<td class="tblTitle">Comment</td>
							<td class="tblTitle">Author</td>
						</tr>
					</xsl:if>
					<xsl:if test="position()=1">
						<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
					</xsl:if>
					<tr>
						<td class="tblData" width="70%">
							<xsl:for-each select="ss:Data">
								<xsl:for-each select="html:Font">
									<xsl:value-of select="."/>
								</xsl:for-each>
							</xsl:for-each>
						</td>
						<td class="tblData">
							<xsl:value-of select="@ss:Author"/>
						</td>
					</tr>
					<xsl:if test="position()=last()">
						<xsl:text disable-output-escaping="yes">&lt;/tbody&gt;</xsl:text>
					</xsl:if>
					<xsl:if test="position()=last()">
						<xsl:text disable-output-escaping="yes">&lt;/table&gt;</xsl:text>
					</xsl:if>
				</xsl:for-each>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>
