<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:output method="text"/>

   <xsl:template match="quotation">
     Matched 1 <xsl:value-of select="category"/> quotation from category <xsl:value-of select="source"/>
   </xsl:template>

</xsl:stylesheet>
