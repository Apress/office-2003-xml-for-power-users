<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
	<xsl:template match="/">
		<xsl:param name="sLink"/>
		<html>
			<head>
				<STYLE type="text/css">
		          .tblHead { background-color: DarkGray; font-weight: bold; border=0}
		          .tblTitle { background-color:LightGrey; font-weight: bold; border=0 }
		          .tblData { border-style:solid}
		        </STYLE>
			</head>
			<body>
				<xsl:if test="count(/ss:Workbook/ss:Names/ss:NamedRange)=0">
					<xsl:text>This workbook does not contain any defined names.</xsl:text>
				</xsl:if>
					<xsl:for-each select="/ss:Workbook/ss:Names/ss:NamedRange">
					<xsl:if test="position()=1">
						<xsl:text disable-output-escaping="yes">&lt;table border="1"width="100%"&gt;</xsl:text>
					</xsl:if>
					<xsl:if test="position()=1">
						<tr>
							<td class="tblTitle">Name</td>
							<td class="tblTitle">Refers To</td>
						</tr>
					</xsl:if>
					<xsl:if test="position()=1">
						<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
					</xsl:if>
					<tr>
						<td class="tblData" width="30%">
							 	<a target="_blank">
									<xsl:attribute name="href">
										<xsl:value-of select="$sLink"/>
										<xsl:text>#</xsl:text>
										<xsl:value-of select="@ss:Name"/>
								    	</xsl:attribute>
									<xsl:value-of select="@ss:Name"/>
								 </a>
						</td>
						<td class="tblData" width="30%">
							<xsl:value-of select="@ss:RefersTo"/>
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
