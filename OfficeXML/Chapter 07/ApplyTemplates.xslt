<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:output method="text"/>

   <xsl:template match="quotationList/quotation">
       <xsl:apply-templates select="source"/>
       <xsl:apply-templates select="category"/>
       <xsl:apply-templates select="text"/>
   </xsl:template>

   <xsl:template match="text">
     TEXT: <xsl:value-of select="."/>
   </xsl:template>
   <xsl:template match="source">
     SOURCE: <xsl:value-of select="."/>
   </xsl:template>
   <xsl:template match="category">
     CATEGORY: <xsl:value-of select="."/>
   </xsl:template>
</xsl:stylesheet>
