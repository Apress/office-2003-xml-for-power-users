<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
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
				<xsl:if test ="count(//bookmark)=0">
					<xsl:text>This document does not contain any bookmarks.</xsl:text>
				</xsl:if>
				<xsl:for-each select="//bookmark">
					<xsl:if test="position()=1">
						<xsl:text disable-output-escaping="yes">&lt;table border="1"width="100%"&gt;</xsl:text>
					</xsl:if>
					<xsl:if test="position()=1">
						<tr>
							<td class="tblHead ">Bookmark Name</td>
							<td class="tblHead ">Text</td>
						</tr>
					</xsl:if>
					<xsl:if test="position()=1">
						<xsl:text disable-output-escaping="yes">&lt;tbody&gt;</xsl:text>
					</xsl:if>
					<tr>
						<td class="tblTitle" width="30%">
							 	<a target="_blank">
									<xsl:attribute name="href">
										<xsl:value-of select="//@link"/>
										<xsl:text>#</xsl:text>
										<xsl:value-of select="name"/>
								    	</xsl:attribute>
									<xsl:value-of select="name"/>
								 </a>
						</td>
						<td class="tblData">
							<xsl:value-of select="text"/>
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
