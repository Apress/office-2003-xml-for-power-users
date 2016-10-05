<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:output method="text"/>
   <xsl:template match="quotationList">
       <xsl:apply-templates select="quotation[category='Ancient Wisdom']"/>
   </xsl:template>
   <xsl:template match="quotation">
     "<xsl:value-of select="text"/>" (<xsl:value-of select="source"/>)
   </xsl:template>
</xsl:stylesheet>

