<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <xsl:output method="text"/>
    <xsl:template match="text">
     [Text]</xsl:template>

    <xsl:template match="source">
     [Source]</xsl:template>

    <xsl:template match="category">
     [Category]</xsl:template>

</xsl:stylesheet>
