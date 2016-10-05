<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:SL="http://schemas.microsoft.com/schemaLibrary/2003/core" xmlns:aml="http://schemas.microsoft.com/aml/2001/core" xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882">
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
			<table width="100%">				
				<xsl:for-each select="w:wordDocument">
					<xsl:for-each select="w:styles">
						<xsl:for-each select="w:style">
							<xsl:if test="position()=1">
								<xsl:text disable-output-escaping="yes">&lt;table border="1"&gt;</xsl:text>
							</xsl:if>
							<xsl:if test="position()=1">
									<tr>
										<td class ="tblHead ">Name</td>
										<td class ="tblHead ">Alias</td>
										<td class ="tblHead ">Type</td>
										<td class ="tblHead ">Based On</td>
										<td class ="tblHead ">Next Style</td>
										<td class ="tblHead ">Paragraph Style Type</td>
										<td class ="tblHead ">Font</td>
										<td class ="tblHead ">Hidden</td>
									</tr>
							</xsl:if>
							<xsl:if test="position()=1">
								<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
							</xsl:if>
							<tbody class="tblData" width="100%">
							<tr>
								<td class ="tblData">
									<xsl:for-each select="w:name">
										<xsl:for-each select="@w:val">
											<xsl:value-of select="."/>
										</xsl:for-each>
									</xsl:for-each>
								</td>
								<td class ="tblData">
									<xsl:for-each select="wx:uiName">
										<xsl:for-each select="@wx:val">
											<xsl:value-of select="."/>
										</xsl:for-each>
									</xsl:for-each>
								</td>
								<td class ="tblData">
									<xsl:for-each select="@w:type">
										<xsl:value-of select="."/>
									</xsl:for-each>
								</td>
								<td class ="tblData">
									<xsl:for-each select="w:basedOn">
										<xsl:for-each select="@w:val">
											<xsl:value-of select="."/>
										</xsl:for-each>
									</xsl:for-each>
								</td>
								<td class ="tblData">
									<xsl:for-each select="w:next">
										<xsl:for-each select="@w:val">
											<xsl:value-of select="."/>
										</xsl:for-each>
									</xsl:for-each>
								</td>
								<td class ="tblData">
									<xsl:for-each select="w:pPr">
										<xsl:for-each select="w:pStyle">
											<xsl:for-each select="@w:val">
												<xsl:value-of select="."/>
											</xsl:for-each>
										</xsl:for-each>
									</xsl:for-each>
								</td>
								<td class ="tblData">
									<xsl:for-each select="w:rPr">
										<xsl:for-each select="wx:font">
											<xsl:for-each select="@wx:val">
												<xsl:value-of select="."/>
											</xsl:for-each>
										</xsl:for-each>
									</xsl:for-each>&#160;
									<xsl:for-each select="w:rPr">
										<xsl:for-each select="w:sz">
											<xsl:for-each select="@w:val">
												<xsl:value-of select="."/>
											</xsl:for-each>
										</xsl:for-each>
									</xsl:for-each>
								</td>
								<td class ="tblData">
									<xsl:for-each select="w:hidden">
										<xsl:for-each select="@w:val">
											<xsl:value-of select="."/>
										</xsl:for-each>
									</xsl:for-each>
								</td>
							</tr>
							<xsl:if test="position()=last()">
								<xsl:text disable-output-escaping="yes">&lt;/tbody&gt;</xsl:text>
							</xsl:if>
							<xsl:if test="position()=last()">
								<xsl:text disable-output-escaping="yes">&lt;/table&gt;</xsl:text>
							</xsl:if>
							</tbody>
						</xsl:for-each>
					</xsl:for-each>
				</xsl:for-each>
				</table>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>
