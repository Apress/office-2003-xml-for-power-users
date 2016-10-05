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
				<xsl:if test="count(.//aml:annotation[@w:type='Word.Comment'])=0">
					<xsl:text>This document does not contain any comments.</xsl:text>
				</xsl:if>
				<xsl:apply-templates select=".//aml:annotation[@w:type='Word.Comment']"/>
			</body>
		</html>
	</xsl:template>
	<xsl:template match="aml:annotation[@w:type='Word.Comment']">
		<xsl:param name="sLink"/>
		<xsl:param name="sType"/>
		<xsl:if test="position()=1">
			<xsl:text disable-output-escaping="yes">&lt;table border="1" width="100%"&gt;</xsl:text>
		</xsl:if>
		<xsl:if test="position()=1">
			<tr class="tblHead">
				<td align="center">Delete</td>
				<td align="center">Edit</td>
				<td>Comment</td>
				<td>Created By</td>
				<td>Date Created</td>
			</tr>
		</xsl:if>
		<xsl:if test="position()=1">
			<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
		</xsl:if>
		<tr>
			<td class="tblTitle" align="center" width="50px">
				<img src="images/bcrossm.bmp" border="0" style="cursor: hand;">
					<xsl:attribute name="onclick"><xsl:text>NoteChange ('</xsl:text><xsl:value-of select="$sLink"/><xsl:text>','Delete','</xsl:text><xsl:value-of select="$sType"/><xsl:text>','</xsl:text><xsl:value-of select="@aml:id"/><xsl:text>')</xsl:text></xsl:attribute>
				</img>
			</td>
			<td class="tblTitle" align="center" width="50px">
				<img src="images/comment.bmp" border="0" style="cursor: hand;">
					<xsl:attribute name="onclick"><xsl:text>NoteChange ('</xsl:text><xsl:value-of select="$sLink"/><xsl:text>','Edit','</xsl:text><xsl:value-of select="$sType"/><xsl:text>','</xsl:text><xsl:value-of select="@aml:id"/><xsl:text>')</xsl:text></xsl:attribute>
				</img>
			</td>
			<td class="tblData" width="60%">
				<xsl:apply-templates select=".//w:p"/>
			</td>
			<td class="tblData">
				<xsl:apply-templates select="@aml:author"/>
			</td>
			<td class="tblData">
				<xsl:call-template name="formatADODate">
					<xsl:with-param name="ado-date">
						<xsl:value-of select="@aml:createdate"/>
					</xsl:with-param>
				</xsl:call-template>
			</td>
		</tr>
		<xsl:if test="position()=last()">
			<xsl:text disable-output-escaping="yes">&lt;/tbody&gt;</xsl:text>
		</xsl:if>
		<xsl:if test="position()=last()">
			<xsl:text disable-output-escaping="yes">&lt;/table&gt;</xsl:text>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:p">
		<div>
			<xsl:value-of select="."/>
		</div>
	</xsl:template>
	<!--Format date to mm/dd/yyyy hh:mm:ss-->
	<xsl:template name="formatADODate">
		<xsl:param name="ado-date"/>
		<xsl:value-of select="substring($ado-date,6,2)"/>/<xsl:value-of select="substring($ado-date,9,2)"/>/<xsl:value-of select="substring($ado-date,1,4)"/>
		<xsl:text> </xsl:text>
		<xsl:value-of select="substring($ado-date,12,3)"/>
		<xsl:value-of select="substring($ado-date,12,8)"/>
	</xsl:template>
</xsl:stylesheet>
