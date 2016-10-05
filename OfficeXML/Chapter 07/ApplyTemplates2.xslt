<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:output method="text"/>

   <xsl:template match="quotationList">
       There are <xsl:value-of select="count(quotation)"/> quotations.

       Sources include: <xsl:apply-templates select="quotation/source"/>

       Categories include: <xsl:apply-templates select="quotation/category"/>
   </xsl:template>

   <xsl:template match="source">
         * <xsl:value-of select="."/></xsl:template>
   <xsl:template match="category">
         * <xsl:value-of select="."/></xsl:template>
</xsl:stylesheet>
