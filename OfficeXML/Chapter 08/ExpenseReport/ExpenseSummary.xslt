<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <xsl:output method="html"/>

    <xsl:template match="/Expenses">
        <font face="Verdana">
        <h2>Expense Reports</h2>
        <xsl:apply-templates select="ExpenseReport"/>
        </font>
    </xsl:template>

    <xsl:template match="ExpenseReport">
        <hr/>
        <font size="4">Expenses - <xsl:value-of select="Name"/></font>
        <font size="1">
        <i><xsl:value-of select="IDNumber"/>: <xsl:value-of select="Purpose"/></i>
        </font>
        <table border="1" cellpadding="1" width="100%">
            <xsl:apply-templates select="ExpenseItem"/>
        </table>
        <br/><br/>
    </xsl:template>

   <xsl:template match="ExpenseItem">
       <tr>
           <td><font size="1"><xsl:value-of select="Date"/></font></td>
           <td><font size="1"><xsl:value-of select="Description"/></font></td>
           <td bgcolor="#FFCC33"><font size="1">$<xsl:value-of select="Total"/></font></td>
       </tr>
   </xsl:template>


</xsl:stylesheet>
