<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:w="http://schemas.microsoft.com/office/word/2003/2/wordml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:SL="http://schemas.microsoft.com/schemaLibrary/2003/2/core" xmlns:aml="http://schemas.microsoft.com/aml/2001/core" xmlns:wx="http://schemas.microsoft.com/office/word/2003/2/auxHint" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882">
	<xsl:template match="/">
		<xsl:apply-templates select="//o:DocumentProperties"/>
	</xsl:template>
	<xsl:template match="//o:DocumentProperties">
		<html>
			<head>
				<STYLE type="text/css">
		          .tblHead { background-color: DarkGray; font-weight: bold; border=0}
		          .tblTitle { background-color:LightGrey; font-weight: bold; border=0 }
		          .tblData { border-style:solid}
		        </STYLE>
			</head>
			<body>
				<table border="1" width="100%">
					<xsl:if test="position()=1">
						<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
					</xsl:if>
					<tr>
						<td class="tblTItle" width="15%">Last Saved By</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:LastAuthor"/>
						</td>
						<td class="tblTItle" width="15%">Template</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:Template"/>
<xsl:call-template name="Template"/>
						</td>
					</tr>
					<tr>
						<td class="tblTItle" width="15%">Revision Number</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:Revision"/>
						</td>
						<td class="tblTItle" width="15%">Application</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:AppName"/>
							<xsl:text>Microsoft Word</xsl:text>							
						</td>
					</tr>
					<tr>
						<td class="tblTItle" width="15%">Total Editing Time</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:TotalTime"/>
							<xsl:if test ="count(o:TotalTime)=0">
								<xsl:text>0</xsl:text>
							</xsl:if>
							<xsl:text> Minutes</xsl:text>
						</td>
						<td class="tblTItle" width="15%">Version</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:Version"/>
						</td>
					</tr>
					<xsl:if test="position()=last()">
						<xsl:text disable-output-escaping="yes">&lt;/tbody&gt;</xsl:text>
					</xsl:if>
				</table>
				<br/>
				<table border="1" width="100%">
					<xsl:if test="position()=1">
						<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
					</xsl:if>
					<tr>
						<td class="tblHead" colspan="2">Dates:</td>
						<td class="tblHead" colspan="2">Counts:</td>
					</tr>
					<tr>
						<td class="tblTItle" width="15%">Created</td>
						<td class="tblData" width="35%">
							<xsl:for-each select="o:Created">
									<xsl:call-template name="formatADODate">
										<xsl:with-param name="ado-date" select="."/>
									</xsl:call-template>
							</xsl:for-each>
						</td>
						<td class="tblTItle" width="15%">Pages</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:Pages"/>
						</td>
					</tr>
					<tr>
						<td class="tblTItle" width="15%">Modified</td>
						<td class="tblData" width="35%">
							<xsl:for-each select="o:LastSaved">
									<xsl:call-template name="formatADODate">
										<xsl:with-param name="ado-date" select="."/>
									</xsl:call-template>
							</xsl:for-each>
						</td>
						<td class="tblTItle" width="15%">Paragraphs</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:Paragraphs"/>
						</td>
					</tr>
					<tr>
						<td class="tblTItle" width="15%">Printed</td>
						<td class="tblData" width="35%">
							<xsl:for-each select="o:LastPrint">
									<xsl:call-template name="formatADODate">
										<xsl:with-param name="ado-date" select="."/>
									</xsl:call-template>
							</xsl:for-each>
							<xsl:if test ="count(o:LastPrint)=0">
								<xsl:text>Source document never printed</xsl:text>
							</xsl:if>
						</td>
						<td class="tblTItle" width="15%">Lines</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:Lines"/>
						</td>
					</tr>
					<tr>
						<td/>
						<td/>
						<td class="tblTItle" width="15%">Words</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:Words"/>
						</td>
					</tr>
					<tr>
						<td/>
						<td/>
						<td class="tblTItle" width="15%">Characters</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:Characters"/>
						</td>
					</tr>
					<tr>
						<td/>
						<td/>
						<td class="tblTItle" width="15%">Characters (with spaces)</td>
						<td class="tblData" width="35%">
							<xsl:value-of select="o:CharactersWithSpaces"/>
						</td>
					</tr>
				</table>
			</body>
		</html>
	</xsl:template>
	<!--Format date to mm/dd/yyyy hh:mm:ss-->
	<xsl:template name="formatADODate">
		<xsl:param name="ado-date"/>
		<xsl:value-of select="substring($ado-date,6,2)"/>/<xsl:value-of select="substring($ado-date,9,2)"/>/<xsl:value-of select="substring($ado-date,1,4)"/>
		<xsl:text> </xsl:text>
		<xsl:value-of select="substring($ado-date,12,3)"/>
		<xsl:value-of select="substring($ado-date,12,8)"/>
	</xsl:template>
	<xsl:template name="Template">
							<xsl:for-each select="/w:wordDocument/w:docPr/w:attachedTemplate">
								<xsl:value-of select="@w:val"/>
								<xsl:if test ="string-length(@w:val)=0">
									<xsl:text>No template specified (Normal.dot)</xsl:text>
								</xsl:if>
							</xsl:for-each>
	</xsl:template>
</xsl:stylesheet>
