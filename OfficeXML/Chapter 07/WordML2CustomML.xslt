<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
xmlns:w="http://schemas.microsoft.com/office/word/2003/2/wordml"
xmlns:ns0="http://www.prosetech.com/Schemas/QuotationList">

    <xsl:output method="xml" indent="yes"/>

    <!-- This catches all nodes, and ignores them, unless
         one of the following two templates is matched. -->
    <xsl:template match="@* | node()">   
        <xsl:apply-templates/>    
    </xsl:template>  

    <!-- Every time an element is matched in the target namespace,
         output the element tag, and then process all children. -->
    <xsl:template match="ns0:*" >
        <xsl:element name="{name()}">
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>

    <!-- Copy the value of any text elements. -->
    <xsl:template match="w:t">
        <xsl:value-of select="."/>
    </xsl:template>

</xsl:stylesheet>
