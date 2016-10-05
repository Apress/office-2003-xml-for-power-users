<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <xsl:output method="html"/>
    <xsl:template match="productCatalog">
        <html>
        <head><title>ProductCatalog</title></head>
        <body>
            <xsl:apply-templates select="catalogName"/>
            <table border="1" cellpadding="2" width="100%">
                <xsl:apply-templates select="products/product"/>
            </table>
            <xsl:apply-templates select="expiryDate"/>
        </body>
        </html>
    </xsl:template>

   <xsl:template match="catalogName">
     <h2><xsl:value-of select="."/></h2>
   </xsl:template>

   <xsl:template match="product">
       <tr>
           <td><xsl:value-of select="@id"/></td>
           <td><xsl:value-of select="productName"/></td>
           <td>$<xsl:value-of select="productPrice"/></td>
           <td>&#160;<xsl:if test="inStock='false'">
               <font color="#FF0000">Sold out!</font>
           </xsl:if></td>
       </tr>
   </xsl:template>

   <xsl:template match="expiryDate">
     <br/><h5><i>expires <xsl:value-of select="."/></i></h5>
   </xsl:template>
</xsl:stylesheet>
