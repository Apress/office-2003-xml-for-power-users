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
					<xsl:if test ="count(//o:CustomDocumentProperties)=0">
						<xsl:text>This document does not contain any custom document properties.</xsl:text>
					</xsl:if>
					<xsl:for-each select="//o:CustomDocumentProperties">
						<xsl:for-each select="*">
						<xsl:if test="position()=1">
							<xsl:text disable-output-escaping="yes">&lt;table border="1"&gt;</xsl:text>
						</xsl:if>
						<xsl:if test="position()=1">
							<tr class="tblHead">
								<td>Property Name</td>
								<td>Value</td>
								<td>Data Type</td>
								</tr>
						</xsl:if>
						<xsl:if test="position()=1">
							<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
						</xsl:if>
							<tr>
								<td class="tblTitle" width="30%">
									<xsl:value-of select="name()"/>
								</td>
								<td class="tblData" width="60%">
									<xsl:value-of select="."/>
								</td>
								<td class="tblData" >
									<xsl:value-of select="@dt:dt"/>
								</td>
							</tr>
						<xsl:if test="position()=last()">
							<xsl:text disable-output-escaping="yes">&lt;/tbody&gt;</xsl:text>
						</xsl:if>
						<xsl:if test="position()=last()">
							<xsl:text disable-output-escaping="yes">&lt;/table&gt;</xsl:text>
						</xsl:if>
					</xsl:for-each>
				</xsl:for-each>
			</body>
		</html>
	</xsl:template>
	<!--Remove's stated Namespace prefix-->
	<xsl:template name="TruncNS">
		<xsl:param name="Element"/>
			<xsl:if test ="starts-with('o:', $Element)">
				<xsl:value-of select="substring-after('o:',$Element)"/>
			</xsl:if>
	</xsl:template>
</xsl:stylesheet>
