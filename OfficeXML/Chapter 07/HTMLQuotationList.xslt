<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <xsl:output method="html"/>
    <xsl:template match="quotationList">
        <html>
        <head><title>Quotation List</title></head>
        <body>
            <ul> 
                <xsl:apply-templates select="quotation"/>
            </ul>
        </body>
        </html>
    </xsl:template>

   <xsl:template match="quotation">
     <li>"<xsl:value-of select="text"/>" <i>(<xsl:value-of select="source"/>)</i></li>
   </xsl:template>
</xsl:stylesheet>
