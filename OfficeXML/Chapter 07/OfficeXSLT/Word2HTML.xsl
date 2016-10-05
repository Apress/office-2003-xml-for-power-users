<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" 
	xmlns:o="urn:schemas-microsoft-com:office:office" 
	xmlns:v="urn:schemas-microsoft-com:vml" 
	xmlns:WX="http://schemas.microsoft.com/office/word/2003/auxHint"
	xmlns:aml="http://schemas.microsoft.com/aml/2001/core"
	xmlns:w10="urn:schemas-microsoft-com:office:word"
	version="1.0">

<!--
This document is provided for informational purposes only. MICROSOFT MAKES NO WARRANTIES, EXPRESS OR IMPLIED, AS TO THE INFORMATION IN THIS DOCUMENT. Complying with all applicable copyright laws is the responsibility of the user.  Microsoft may have patents, patent applications, trademarks, copyrights, or other intellectual property rights covering subject matter in this document. Except as expressly provided in any written license agreement from Microsoft, the furnishing of this document does not give you any license to these patents, trademarks, copyrights, or other intellectual property. © 2003 Microsoft Corporation. All rights reserved.
-->

<!--
Please note that this is just a beta version of the Word2Html transform.
This transform is intended to work with Microsoft Office Word 2003 Beta 2 XML documents. A quick way to verify that the XML document you have is from the Beta 2 version of Microsoft Office Word 2003 is to look at the namespace. If the Word namespace is:
“http://schemas.microsoft.com/office/word/2003/wordml”, then it should be the correct version.
Also, please note that images and other embedded objects will most likely not appear. This is a known limitation of this XSLT, but it will be fixed for the final version. 
-->


<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes"/>

<!--
	VARIABLE NAMING CONVENTION
	t.	 = type and identifiers
	i.	 = index value
	b.	 = boolean ($on, $off, depending on situation, NA can be $na or just left empty)
	p.	 = path, one node
	ns.	 = node set
	pr.	 = property (a string to used to represent a property value, how it is being formated depends on the type of property)
	prs. = a string that is used to represent more than one property (usually concat together with a $sepa)
	
	Conv    = convertone format to another
	Get     = usually return a property after some searching
	Apply   = the resulting string is usually something that can be applied as CSS style
	Wrap    = wrap an element around some stuff inside
	Display = the resulting string should represent the element in html format
	
	.once	= something to be called once, directly at the element
	.class	= called once at each CSS class level
	.many   = called many times, once at each word style (group of properties) level, for accumulation
	
-->

<!-- CONSTANTS -->

<!-- default styleId -->
<xsl:variable name="pStyleId.default">Normal</xsl:variable>
<xsl:variable name="tStyleId.default">TableNormal</xsl:variable>

<!-- class name suffix -->
<xsl:variable name="styleSuffix.table">-T</xsl:variable>
<xsl:variable name="styleSuffix.row">-R</xsl:variable>
<xsl:variable name="styleSuffix.cell">-C</xsl:variable>
<xsl:variable name="styleSuffix.para">-P</xsl:variable>
<xsl:variable name="styleSuffix.char">-H</xsl:variable>

<!-- default paragraph margins -->
<xsl:variable name="pMargin.default.top">0pt</xsl:variable>
<xsl:variable name="pMargin.default.right">0pt</xsl:variable>
<xsl:variable name="pMargin.default.bottom">.0001pt</xsl:variable>
<xsl:variable name="pMargin.default.left">0pt</xsl:variable>

<!-- contextual spacing identifiers -->
<xsl:variable name="t.cSpacing.all"></xsl:variable>
<xsl:variable name="t.cSpacing.top">t</xsl:variable>
<xsl:variable name="t.cSpacing.bottom">b</xsl:variable>
<xsl:variable name="t.cSpacing.none">
	<xsl:value-of select="$t.cSpacing.top"/><xsl:value-of select="$t.cSpacing.bottom"/>
</xsl:variable>

<!-- border side identifiers and suffix -->
<xsl:variable name="bdrSide.top">-top</xsl:variable>
<xsl:variable name="bdrSide.right">-right</xsl:variable>
<xsl:variable name="bdrSide.bottom">-bottom</xsl:variable>
<xsl:variable name="bdrSide.left">-left</xsl:variable>
<xsl:variable name="bdrSide.char"></xsl:variable>

<!-- identifiers for property retrievers -->
<xsl:variable name="t.frame">1</xsl:variable>
<xsl:variable name="t.defaultCellpadding">2</xsl:variable>
<xsl:variable name="t.cellspacing">3</xsl:variable>
<xsl:variable name="t.bdrPr.top">4</xsl:variable>
<xsl:variable name="t.bdrPr.right">5</xsl:variable>
<xsl:variable name="t.bdrPr.bottom">6</xsl:variable>
<xsl:variable name="t.bdrPr.left">7</xsl:variable>
<xsl:variable name="t.bdrPr.between">8</xsl:variable>
<xsl:variable name="t.bdrPr.bar">9</xsl:variable>
<xsl:variable name="t.bdrPr.insideH">A</xsl:variable>
<xsl:variable name="t.bdrPr.insideV">B</xsl:variable>
<xsl:variable name="t.listSuff">C</xsl:variable>
<xsl:variable name="t.listInd">D</xsl:variable>
<xsl:variable name="t.applyRPr">E</xsl:variable>
<xsl:variable name="t.updateRPr">F</xsl:variable>
<xsl:variable name="t.applyTcPr">G</xsl:variable>
<xsl:variable name="t.customCellpadding">H</xsl:variable>
<xsl:variable name="t.trCantSplit">I</xsl:variable>
<xsl:variable name="t.tblInd">J</xsl:variable>

<!-- type names for cnf styles -->
<xsl:variable name="cnfType.firstRow">firstRow</xsl:variable>
<xsl:variable name="cnfType.lastRow">lastRow</xsl:variable>
<xsl:variable name="cnfType.firstCol">firstCol</xsl:variable>
<xsl:variable name="cnfType.lastCol">lastCol</xsl:variable>
<xsl:variable name="cnfType.band1Vert">band1Vert</xsl:variable>
<xsl:variable name="cnfType.band2Vert">band2Vert</xsl:variable>
<xsl:variable name="cnfType.band1Horz">band1Horz</xsl:variable>
<xsl:variable name="cnfType.band2Horz">band2Horz</xsl:variable>
<xsl:variable name="cnfType.neCell">neCell</xsl:variable>
<xsl:variable name="cnfType.nwCell">nwCell</xsl:variable>
<xsl:variable name="cnfType.seCell">seCell</xsl:variable>
<xsl:variable name="cnfType.swCell">swCell</xsl:variable>

<!-- position of the cnf type in the cnfStyle binary -->
<xsl:variable name="i.cnfType.firstRow">1</xsl:variable>
<xsl:variable name="i.cnfType.lastRow">2</xsl:variable>
<xsl:variable name="i.cnfType.firstCol">3</xsl:variable>
<xsl:variable name="i.cnfType.lastCol">4</xsl:variable>
<xsl:variable name="i.cnfType.band1Vert">5</xsl:variable>
<xsl:variable name="i.cnfType.band2Vert">6</xsl:variable>
<xsl:variable name="i.cnfType.band1Horz">7</xsl:variable>
<xsl:variable name="i.cnfType.band2Horz">8</xsl:variable>
<xsl:variable name="i.cnfType.neCell">9</xsl:variable>
<xsl:variable name="i.cnfType.nwCell">10</xsl:variable>
<xsl:variable name="i.cnfType.seCell">11</xsl:variable>
<xsl:variable name="i.cnfType.swCell">12</xsl:variable>

<!-- boolean constants, also used within an encoded property -->
<xsl:variable name="off">0</xsl:variable>
<xsl:variable name="on">1</xsl:variable>
<xsl:variable name="na">2</xsl:variable>

<!-- separators used within an encoded property -->
<xsl:variable name="sepa">/</xsl:variable>
<xsl:variable name="sepa1">|</xsl:variable>
<xsl:variable name="sepa2">,</xsl:variable>

<!-- automatic color constants -->
<xsl:variable name="autoColor.hex">auto</xsl:variable>
<xsl:variable name="autoColor.text">windowtext</xsl:variable>
<xsl:variable name="autoColor.bg">transparent</xsl:variable>

<!-- automatic color constants -->
<xsl:variable name="transparentColor.hex">transparent</xsl:variable>
<xsl:variable name="transparentColor.text">transparent</xsl:variable>
<xsl:variable name="transparentColor.bg">transparent</xsl:variable>

<!-- values of list w:suff property -->
<xsl:variable name="pr.listSuff.space">Space</xsl:variable>
<xsl:variable name="pr.listSuff.nothing">Nothing</xsl:variable>

<!-- common node-sets and paths -->
<xsl:variable name="ns.styles" select="/w:wordDocument[1]/w:styles[1]/w:style"/>
<xsl:variable name="p.lists" select="/w:wordDocument[1]/w:lists[1]"/>
<xsl:variable name="p.docPr" select="/w:wordDocument[1]/w:docPr[1]"/>
<xsl:variable name="p.docInfo" select="/w:wordDocument[1]/w:docInfo[1]"/>

<!-- web options -->
<xsl:variable name="pixelsPerInch">
	<xsl:choose>
		<xsl:when test="$p.docPr/w:pixelsPerInch/@w:val">
			<xsl:value-of select="$p.docPr/w:pixelsPerInch/@w:val"/>
		</xsl:when>
		<xsl:otherwise>96</xsl:otherwise>
	</xsl:choose>
</xsl:variable>

<!-- property positions for an encoded rPr -->
<xsl:variable name="i.emboss-imprint">1</xsl:variable>
<xsl:variable name="i.u-em">2</xsl:variable>
<xsl:variable name="i.strike-dstrike">3</xsl:variable>
<xsl:variable name="i.sup">4</xsl:variable>
<xsl:variable name="i.sub">5</xsl:variable>
<xsl:variable name="i.vanish-webhidden">6</xsl:variable>
<xsl:variable name="i.bcs">7</xsl:variable>
<xsl:variable name="i.ics">8</xsl:variable>
<xsl:variable name="i.szcs">9</xsl:variable>

<!-- property positions for an encoded pPr -->
<xsl:variable name="i.textAutospace.o">1</xsl:variable>
<xsl:variable name="i.textAutospace.n">2</xsl:variable>
<xsl:variable name="i.ind">3</xsl:variable>

<!-- default encoded properties -->
<xsl:variable name="prs.r.default">
	<xsl:value-of select="$na"/><xsl:value-of select="$na"/><xsl:value-of select="$na"/><xsl:value-of select="$na"/><xsl:value-of select="$na"/>
	<xsl:value-of select="$na"/><xsl:value-of select="$na"/><xsl:value-of select="$na"/><xsl:value-of select="$na"/>
</xsl:variable>
<xsl:variable name="prs.p.default">
	<xsl:value-of select="$na"/><xsl:value-of select="$na"/>
</xsl:variable>


<!-- GENERAL TEMPLATES -->

<!-- convert a hexBinary string to individual decimal numbers separated by space -->
<xsl:template name="ConvHex2Dec">
	<xsl:param name="value"/>
	<xsl:param name="i" select="1"/>
	<xsl:param name="s" select="1"/>
	<xsl:variable name="hexDigit" select="substring($value,$i,1)"/>
	<xsl:if test="not($hexDigit = '')">
		<xsl:text> </xsl:text>
		<xsl:choose>
			<xsl:when test="$hexDigit = 'A'">10</xsl:when>
			<xsl:when test="$hexDigit = 'B'">11</xsl:when>
			<xsl:when test="$hexDigit = 'C'">12</xsl:when>
			<xsl:when test="$hexDigit = 'D'">13</xsl:when>
			<xsl:when test="$hexDigit = 'E'">14</xsl:when>
			<xsl:when test="$hexDigit = 'F'">15</xsl:when>
			<xsl:otherwise><xsl:value-of select="$hexDigit"/></xsl:otherwise>
		</xsl:choose>
		<xsl:call-template name="ConvHex2Dec">
			<xsl:with-param name="value" select="$value"/>
			<xsl:with-param name="i" select="$i+$s"/>
			<xsl:with-param name="s" select="$s"/>
		</xsl:call-template>
	</xsl:if>
</xsl:template>

<!-- convert border style type from Word to CSS -->
<xsl:template name="ConvBorderStyle">
	<xsl:param name="value"/>
	<xsl:choose>
		<xsl:when test="$value='none' or $value='nil'">none</xsl:when>
		<xsl:when test="$value='single'">solid</xsl:when>
		<xsl:when test="contains($value,'stroke')">solid</xsl:when>
		<xsl:when test="$value='dashed'">dashed</xsl:when>
		<xsl:when test="contains($value,'dash')">dashed</xsl:when>
		<xsl:when test="$value='double'">double</xsl:when>
		<xsl:when test="$value='triple'">double</xsl:when>
		<xsl:when test="contains($value,'double')">double</xsl:when>
		<xsl:when test="contains($value,'gap')">double</xsl:when>
		<xsl:when test="$value='dotted'">dotted</xsl:when>
		<xsl:when test="$value='three-d-emboss'">ridge</xsl:when>
		<xsl:when test="$value='three-d-engrave'">groove</xsl:when>
		<xsl:when test="$value='outset'">outset</xsl:when>
		<xsl:when test="$value='inset'">inset</xsl:when>
		<xsl:otherwise>solid</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- evaluate TableWidthProperty to either % or pt value -->
<xsl:template name="EvalTableWidth">
	<xsl:choose>
		<xsl:when test="@w:type = 'pct'"><xsl:value-of select="@w:w div 50"/>%</xsl:when>
		<xsl:otherwise><xsl:value-of select="@w:w div 20"/>pt</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- convert Word color to HTML color -->
<xsl:template name="ConvColor">
	<xsl:param name="value"/>
	<xsl:choose>
		<xsl:when test="$value='black'">black</xsl:when>
		<xsl:when test="$value='blue'">blue</xsl:when>
		<xsl:when test="$value='cyan'">aqua</xsl:when>
		<xsl:when test="$value='green'">lime</xsl:when>
		<xsl:when test="$value='magenta'">fuchsia</xsl:when>
		<xsl:when test="$value='red'">red</xsl:when>
		<xsl:when test="$value='yellow'">yellow</xsl:when>
		<xsl:when test="$value='white'">white</xsl:when>
		<xsl:when test="$value='dark-blue'">navy</xsl:when>
		<xsl:when test="$value='dark-cyan'">teal</xsl:when>
		<xsl:when test="$value='dark-green'">green</xsl:when>
		<xsl:when test="$value='dark-magenta'">purple</xsl:when>
		<xsl:when test="$value='dark-red'">maroon</xsl:when>
		<xsl:when test="$value='dark-yellow'">olive</xsl:when>
		<xsl:when test="$value='dark-gray'">gray</xsl:when>
		<xsl:when test="$value='light-gray'">silver</xsl:when>
		<xsl:when test="$value='none'">transparent</xsl:when>
	</xsl:choose>
</xsl:template>

<!-- convert Word color in hex to HTML color in hex 
		auto translates to param autoColor		-->
<xsl:template name="ConvHexColor">
	<xsl:param name="value"/>
	<xsl:param name="autoColor" select="$autoColor.text"/>
	<xsl:param name="transparentColor">transparent</xsl:param>
	<xsl:choose>
		<xsl:when test="$value = $autoColor.hex or $value = ''">
			<xsl:value-of select="$autoColor"/>
		</xsl:when>
		<xsl:when test="$value = $transparentColor.hex">
			<xsl:value-of select="$transparentColor"/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="concat('#',$value)"/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- evaluate a BooleanType property to $on or $off -->
<xsl:template name="EvalBooleanType">
	<xsl:choose>
		<xsl:when test="@w:val = 'off' or @w:val = 'none'"><xsl:value-of select="$off"/></xsl:when>
		<xsl:otherwise><xsl:value-of select="$on"/></xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- retrieve values of BorderProperty separated by $sepa2 -->
<xsl:template name="GetBorderPr">
		<xsl:value-of select="@w:val"/><xsl:value-of select="$sepa2"/>
		<xsl:value-of select="@w:color"/><xsl:value-of select="$sepa2"/>
		<xsl:choose>
			<xsl:when test="@WX:bdrwidth">
				<xsl:value-of select="@WX:bdrwidth"/><xsl:value-of select="$sepa2"/>
			</xsl:when>
			<xsl:otherwise>0<xsl:value-of select="$sepa2"/></xsl:otherwise>
		</xsl:choose>
		<xsl:value-of select="@w:space"/><xsl:value-of select="$sepa2"/>
		<xsl:value-of select="@w:shadow"/>
</xsl:template>

<!-- apply the border property -->
<xsl:template name="ApplyBorderPr">
	<xsl:param name="pr.bdr"/>
	<xsl:param name="bdrSide" select="$bdrSide.char"/>
	<xsl:if test="not($pr.bdr='')">
		<xsl:text>border</xsl:text><xsl:value-of select="$bdrSide"/><xsl:text>:</xsl:text>
		<xsl:call-template name="ConvBorderStyle">
			<xsl:with-param name="value" select="substring-before($pr.bdr,$sepa2)"/>
		</xsl:call-template>
		<xsl:variable name="temp" select="substring-after($pr.bdr,$sepa2)"/>
		<xsl:text> </xsl:text>
		<xsl:call-template name="ConvHexColor">
			<xsl:with-param name="value" select="substring-before($temp,$sepa2)"/>
		</xsl:call-template>
		<xsl:text> </xsl:text>
		<xsl:value-of select="substring-before(substring-after($temp,$sepa2),$sepa2) div 20"/><xsl:text>pt;</xsl:text>
		<xsl:if test="$bdrSide = $bdrSide.char">padding:0;</xsl:if>
	</xsl:if>
</xsl:template>

<!-- apply CSS background-color from ShdProperty -->
<xsl:template name="ApplyShd">
	<xsl:text>background-color:</xsl:text>
	<xsl:choose>
		<!-- use shd fill color when shd type is "clear" -->
		<xsl:when test="@w:val = 'clear' or not(@w:val)">
			<xsl:call-template name="ConvHexColor">
				<xsl:with-param name="value" select="@w:fill"/>
				<xsl:with-param name="autoColor" select="$autoColor.bg"/>
			</xsl:call-template>
		</xsl:when>
		<!-- otherwise, use bgcolor -->
		<xsl:otherwise>
			<xsl:call-template name="ConvHexColor"><xsl:with-param name="value" select="@WX:bgcolor"/><xsl:with-param name="autoColor" select="$autoColor.bg"/></xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
	<xsl:text>;</xsl:text>
</xsl:template>

<!-- apply CSS background-color from ShdProperty -->
<xsl:template name="ApplyShdHint">
	<xsl:text>background-color:</xsl:text>
	<xsl:call-template name="ConvHexColor">
		<xsl:with-param name="value" select="@WX:bgcolor"/>
		<xsl:with-param name="autoColor" select="$autoColor.bg"/>
		<xsl:with-param name="transparentColor">transparent</xsl:with-param>
	</xsl:call-template>
	<xsl:text>;</xsl:text>
</xsl:template>
<!-- apply CSS layout-flow from TextDirectionProperty -->
<xsl:template name="ApplyTextDirection">
	<xsl:text>layout-flow:</xsl:text>
	<xsl:choose>
		<xsl:when test="@w:val = 'tb-rl-v'">vertical-ideographic</xsl:when>
		<xsl:when test="@w:val = 'lr-tb-v'">horizontal-ideographic</xsl:when>
		<xsl:otherwise>normal</xsl:otherwise>
	</xsl:choose>
	<xsl:text>;</xsl:text>
</xsl:template>

<!-- apply CSS padding for table cells, either tcMarElt or tblCellMarElt -->
<xsl:template name="ApplyCellMar">
	<xsl:choose>
		<xsl:when test="@w:val='none'">none</xsl:when>
		<xsl:otherwise>
			<xsl:text>padding:</xsl:text>
			<xsl:choose><xsl:when test="w:top"><xsl:for-each select="w:top[1]"><xsl:call-template name="EvalTableWidth"/></xsl:for-each></xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
			<xsl:choose><xsl:when test="w:right"><xsl:for-each select="w:right[1]"><xsl:call-template name="EvalTableWidth"/></xsl:for-each></xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
			<xsl:choose><xsl:when test="w:bottom"><xsl:for-each select="w:bottom[1]"><xsl:call-template name="EvalTableWidth"/></xsl:for-each></xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
			<xsl:choose><xsl:when test="w:left"><xsl:for-each select="w:left[1]"><xsl:call-template name="EvalTableWidth"/></xsl:for-each></xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose><xsl:text>;</xsl:text>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>


<!-- PROPERTY RETRIEVERS -->

<!-- updates the encoded paragraph properties with the context style 
		needs to be called at every style level in the following order
		tstyle, conditional formatings, pstyle, and p direct			-->
<xsl:template name="UpdatePPr">
	<xsl:param name="prs.p" select="$prs.p.default"/>
	<xsl:param name="p.style" select="."/>
	<xsl:variable name="prs.p.temp">
		<xsl:for-each select="$p.style">
			<xsl:call-template name="UpdatePPr.a">
				<xsl:with-param name="prs.p" select="$prs.p"/>
			</xsl:call-template>
		</xsl:for-each>
	</xsl:variable>
	<xsl:choose>
		<xsl:when test="$prs.p.temp=''">
			<xsl:value-of select="$prs.p"/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="$prs.p.temp"/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<xsl:template name="UpdatePPr.a">
	<xsl:param name="prs.p" select="$prs.p.default"/>
	<xsl:for-each select="w:pPr[1]">
		<!-- 1 textAutospace.o -->
		<xsl:variable name="b.textAutospace.o">
			<xsl:for-each select="w:autoSpaceDE[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.textAutospace.o=''"><xsl:value-of select="substring($prs.p,$i.textAutospace.o,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.textAutospace.o"/></xsl:otherwise>
		</xsl:choose>
		<!-- 2 textAutospace.n -->
		<xsl:variable name="b.textAutospace.n">
			<xsl:for-each select="w:autoSpaceDN[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.textAutospace.n=''"><xsl:value-of select="substring($prs.p,$i.textAutospace.n,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.textAutospace.n"/></xsl:otherwise>
		</xsl:choose>
		<!-- 3 ind -->
		<xsl:variable name="pr.ind">
			<xsl:for-each select="w:ind[1]">
				<xsl:value-of select="@w:left"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:left-chars"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:right"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:right-chars"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:hanging"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:hanging-chars"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:first-line"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:first-line-chars"/>				
			</xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$pr.ind=''"><xsl:value-of select="substring($prs.p,$i.ind)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$pr.ind"/></xsl:otherwise>
		</xsl:choose>
	</xsl:for-each>
</xsl:template>

<!-- updates the encoded character properties with the context style 
		needs to be called at every style level in the following order
		tstyle, conditional formatings, pstyle, rstyle, r direct (and listPr for list index) 
		UpdateRPr calls UpdateRPr.a-->
<xsl:template name="UpdateRPr">
	<xsl:param name="prs.r" select="$prs.r.default"/>
	<xsl:param name="p.style" select="."/>
	<xsl:variable name="prs.r.temp">
		<xsl:for-each select="$p.style">
			<xsl:call-template name="UpdateRPr.a">
				<xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:for-each>
	</xsl:variable>
	<xsl:choose>
		<xsl:when test="$prs.r.temp=''">
			<xsl:value-of select="$prs.r"/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="$prs.r.temp"/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template name="UpdateRPr.a">
	<xsl:param name="prs.r" select="$prs.r.default"/>
	<xsl:param name="type" select="non-list"/>
	<xsl:for-each select="w:rPr[1]">
		<!-- 1 emboss-imprint -->
		<xsl:variable name="b.emboss-imprint">
			<xsl:variable name="condition1"><xsl:for-each select="w:emboss[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each></xsl:variable>
			<xsl:variable name="condition2"><xsl:for-each select="w:imprint[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each></xsl:variable>
			<xsl:choose>
				<xsl:when test="$condition1 = $on or $condition2 = $on">
					<xsl:value-of select="$on"/>
				</xsl:when>
				<xsl:when test="$condition1 = $off or $condition2 = $off">
					<xsl:value-of select="$off"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.emboss-imprint = ''"><xsl:value-of select="substring($prs.r,$i.emboss-imprint,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.emboss-imprint"/></xsl:otherwise>
		</xsl:choose>
		<!-- 2 u-em -->
		<xsl:variable name="b.u-em">
			<xsl:variable name="condition1"><xsl:for-each select="w:u[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each></xsl:variable>
			<xsl:variable name="condition2"><xsl:for-each select="w:em[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each></xsl:variable>
			<xsl:choose><xsl:when test="$condition1 = $on or $condition2 = $on"><xsl:value-of select="$on"/></xsl:when><xsl:when test="$condition1 = $off or $condition2 = $off"><xsl:value-of select="$off"/></xsl:when></xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.u-em = ''">
				<xsl:choose>
					<xsl:when test="$type='list'">
						<xsl:value-of select="$off"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="substring($prs.r,$i.u-em,1)"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.u-em"/></xsl:otherwise>
		</xsl:choose>
		<!-- 3 strike-dstrike -->
		<xsl:variable name="b.strike-dstrike">
			<xsl:variable name="condition1"><xsl:for-each select="w:strike[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each></xsl:variable>
			<xsl:variable name="condition2"><xsl:for-each select="w:dstrike[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each></xsl:variable>
			<xsl:choose><xsl:when test="$condition1 = $on or $condition2 = $on"><xsl:value-of select="$on"/></xsl:when><xsl:when test="$condition1 = $off or $condition2 = $off"><xsl:value-of select="$off"/></xsl:when></xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.strike-dstrike = ''"><xsl:value-of select="substring($prs.r,$i.strike-dstrike,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.strike-dstrike"/></xsl:otherwise>
		</xsl:choose>
		<!-- 4 sup -->
		<xsl:variable name="b.sup">
			<xsl:for-each select="w:sup[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.sup = ''"><xsl:value-of select="substring($prs.r,$i.sup,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.sup"/></xsl:otherwise>
		</xsl:choose>
		<!-- 5 sub -->
		<xsl:variable name="b.sub">
			<xsl:for-each select="w:sub[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.sub = ''"><xsl:value-of select="substring($prs.r,$i.sub,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.sub"/></xsl:otherwise>
		</xsl:choose>
		<!-- 6 vanish-webhidden -->
		<xsl:variable name="b.vanish-webhidden">
			<xsl:variable name="condition1"><xsl:for-each select="w:vanish[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each></xsl:variable>
			<xsl:variable name="condition2"><xsl:for-each select="w:webHidden[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each></xsl:variable>
			<xsl:choose><xsl:when test="$condition1 = $on or $condition2 = $on"><xsl:value-of select="$on"/></xsl:when><xsl:when test="$condition1 = $off or $condition2 = $off"><xsl:value-of select="$off"/></xsl:when></xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.vanish-webhidden = ''"><xsl:value-of select="substring($prs.r,$i.vanish-webhidden,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.vanish-webhidden"/></xsl:otherwise>
		</xsl:choose>
		<!-- 7 bcs -->
		<xsl:variable name="b.bcs">
			<xsl:for-each select="w:b-cs[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.bcs = ''"><xsl:value-of select="substring($prs.r,$i.bcs,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.bcs"/></xsl:otherwise>
		</xsl:choose>
		<!-- 8 ics -->
		<xsl:variable name="b.ics">
			<xsl:for-each select="w:i-cs[1]"><xsl:call-template name="EvalBooleanType"/></xsl:for-each>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$b.ics = ''"><xsl:value-of select="substring($prs.r,$i.ics,1)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$b.ics"/></xsl:otherwise>
		</xsl:choose>
		<!-- 9 szcs -->
		<xsl:variable name="pr.szcs" select="string(w:sz-cs[1]/@w:val)"/>
		<xsl:choose>
			<xsl:when test="$pr.szcs = ''"><xsl:value-of select="substring($prs.r,$i.szcs)"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$pr.szcs"/></xsl:otherwise>
		</xsl:choose>
	</xsl:for-each>
</xsl:template>

<!-- retrieve specific list property(s) or perform a specific action at list property
		needs to be called at w:pPr
		param prs.r is optional unless param type = $updateRPr
		GetListPr branches into GetListPr.a
		GetListPr.a branches into GetListPr.b	-->
<xsl:template name="GetListPr">
	<xsl:param name="type"/>
	<xsl:param name="prs.r"/>
	<xsl:for-each select="w:listPr">
		<xsl:choose>
			<!-- direct listPr overrides paragraph style's listPr -->
			<xsl:when test="w:ilfo and w:ilvl">
				<xsl:call-template name="GetListPr.a"><xsl:with-param name="type" select="$type"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:when>
			<!-- find ilfo and ilvl from paragraph style's listPr -->
			<xsl:otherwise>
				<xsl:variable name="pstyleId">
					<xsl:for-each select="ancestor::w:p[1]">
						<xsl:call-template name="GetPStyleId"/>	
					</xsl:for-each>
				</xsl:variable>
				<xsl:for-each select="($ns.styles[@w:styleId=$pstyleId])[1]/w:pPr[1]/w:listPr[1]">
					<xsl:call-template name="GetListPr.a"><xsl:with-param name="type" select="$type"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
				</xsl:for-each>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:for-each>
</xsl:template>
<xsl:template name="GetListPr.a">
	<xsl:param name="type"/>
	<xsl:param name="prs.r"/>
	<!-- retrieve ilfo and ilvl -->
	<xsl:variable name="ilfo" select="w:ilfo/@w:val"/>
	<xsl:variable name="ilvl" select="w:ilvl/@w:val"/>
	<xsl:for-each select="$p.lists">
		<!-- select the correct list style -->
		<xsl:variable name="list" select="w:list[@w:ilfo=$ilfo][1]"/>
		<xsl:choose>
			<!-- go into lvlOverride if it exists for that level -->
			<xsl:when test="$list/w:lvlOverride[@w:ilvl=$ilvl]">
				<xsl:for-each select="$list/w:lvlOverride[@w:ilvl=$ilvl]">
					<xsl:call-template name="GetListPr.b"><xsl:with-param name="type" select="$type"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
			<!-- normal list style definition if lvlOverride doesn't exist -->
			<xsl:otherwise>
				<xsl:for-each select="w:listDef[@w:listDefId=$list/w:ilst/@w:val][1]/w:lvl[@w:ilvl=$ilvl][1]">
					<xsl:call-template name="GetListPr.b"><xsl:with-param name="type" select="$type"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
				</xsl:for-each>	
			</xsl:otherwise>
		</xsl:choose>
	</xsl:for-each>
</xsl:template>

<xsl:template name="GetListPr.b">
	<xsl:param name="type"/>
	<xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- retrieve list suff property -->
		<xsl:when test="$type = $t.listSuff">
			<xsl:variable name="suff" select="w:suff[1]/@w:val"/>
			<xsl:choose>
				<xsl:when test="$suff = $pr.listSuff.space or $suff = $pr.listSuff.nothing"><xsl:value-of select="$suff"/></xsl:when>
				<!-- KL: handle list-tab with one space, tab hints may be able to fix this -->
				<xsl:otherwise><xsl:value-of select="$pr.listSuff.space"/></xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!-- retrieve list indentation property -->
		<xsl:when test="$type = $t.listInd">
			<xsl:for-each select="w:pPr[1]/w:ind[1]">
				<xsl:value-of select="@w:left"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:left-chars"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:hanging"/><xsl:value-of select="$sepa2"/>
				<xsl:value-of select="@w:hanging-chars"/>
			</xsl:for-each>
		</xsl:when>
		<!-- call ApplyRPr on list index's rPr -->
		<xsl:when test="$type = $t.applyRPr">
			<xsl:call-template name="ApplyRPr.class"/>
		</xsl:when>
		<!-- call UpdateRPr  on list index's rPr -->
		<xsl:when test="$type = $t.updateRPr">
			<xsl:call-template name="UpdateRPr.a">
				<xsl:with-param name="type">list</xsl:with-param>
				<xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:when>
	</xsl:choose>
</xsl:template>

<!-- retrieve specific paragraph property(s) at the paragraph level (direct, style)
		first direct property, then property from style
		needs to be called at w:p
		GetPPr branches into GetPPr.a	-->
<xsl:template name="GetPPr">
	<xsl:param name="type"/>
	<xsl:param name="p.pStyle"/>
	<xsl:variable name="result">
		<xsl:call-template name="GetPPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
	</xsl:variable>
	<xsl:if test="$result=''">
		<xsl:for-each select="$p.pStyle">
			<xsl:call-template name="GetPPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
		</xsl:for-each>	
	</xsl:if>
	<xsl:value-of select="$result"/>
</xsl:template>
<xsl:template name="GetPPr.a">
	<xsl:param name="type"/>
	<xsl:for-each select="w:pPr[1]">
		<xsl:choose>
			<xsl:when test="$type = $t.bdrPr.top">
				<xsl:for-each select="w:bdr[1]/w:top[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.right">
				<xsl:for-each select="w:bdr[1]/w:right[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.bottom">
				<xsl:for-each select="w:bdr[1]/w:bottom[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.left">
				<xsl:for-each select="w:bdr[1]/w:left[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.between">
				<xsl:for-each select="w:bdr[1]/w:between[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.bar">
				<xsl:for-each select="w:bdr[1]/w:bar[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.frame">
				<xsl:for-each select="w:framePr[1]">
					<xsl:value-of select="@w:w"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:h"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:h-rule"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:x-align"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:vspace"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:hspace"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:wrap"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:drop-cap"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:lines"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:x"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:y-align"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:y"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:hanchor"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:vanchor"/><xsl:value-of select="$sepa2"/>
					<xsl:value-of select="@w:anchor-lock"/>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
	</xsl:for-each>
</xsl:template>

<!-- retrieve specific table property(s) at the table level (direct, style)
		first direct property, then property from style
		needs to be called at w:tbl
		GetTblPr branches into GetTblPr.a	-->
<xsl:template name="GetTblPr">
	<xsl:param name="type"/>
	<xsl:param name="p.tStyle"/>
	<xsl:variable name="result">
		<xsl:call-template name="GetTblPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
	</xsl:variable>
	<xsl:if test="$result=''">
		<xsl:for-each select="$p.tStyle">
			<xsl:call-template name="GetTblPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
		</xsl:for-each>	
	</xsl:if>
	<xsl:value-of select="$result"/>
</xsl:template>
<xsl:template name="GetTblPr.a">
	<xsl:param name="type"/>
	<xsl:for-each select="w:tblPr[1]">
		<xsl:choose>
			<xsl:when test="$type = $t.bdrPr.top">
				<xsl:for-each select="w:tblBorders[1]/w:top[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.left">
				<xsl:for-each select="w:tblBorders[1]/w:left[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.bottom">
				<xsl:for-each select="w:tblBorders[1]/w:bottom[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.right">
				<xsl:for-each select="w:tblBorders[1]/w:right[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.insideH">
				<xsl:for-each select="w:tblBorders[1]/w:insideH[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.bdrPr.insideV">
				<xsl:for-each select="w:tblBorders[1]/w:insideV[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>		
			</xsl:when>
			<xsl:when test="$type = $t.defaultCellpadding">
				<xsl:for-each select="w:tblCellMar[1]"><xsl:call-template name="ApplyCellMar"/></xsl:for-each>
			</xsl:when>
			<xsl:when test="$type = $t.cellspacing">
				<xsl:value-of select="w:tblCellSpacing[1]/@w:w"/>
			</xsl:when>	
			<xsl:when test="$type = $t.tblInd">
				<xsl:for-each select="w:tblInd[1]">
					<xsl:call-template name="EvalTableWidth"/>
				</xsl:for-each>
			</xsl:when>	
		</xsl:choose>
	</xsl:for-each>
</xsl:template>

<!-- for each conditional format, if it exists, wrap DIV around the content of TD and apply the CSS class of the paragraph and character styles -->
<!-- there are 5 levels of wrapping to be considered,
		first (lowest priority) are the  horizontal bands (WrapCnf)
		second are the vertical bands (WrapCnf.a)
		third are the first/last columns (WrapCnf.b)
		forth are the first/last rows (WrapCnf.c)
		last (highest priority) are the corner cells (WrapCnf.d) -->
<!-- at each level, the prs.p and prs.r are being updated, and prs.pMany will accumunate -->
<xsl:template name="WrapCnf">
	<xsl:param name="p.tStyle"/><xsl:param name="cnfCol"/><xsl:param name="cnfRow"/>
	<xsl:param name="prs.pMany"/><xsl:param name="prs.p"/><xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- horizontal band 1 -->
		<xsl:when test="substring($cnfRow,$i.cnfType.band1Horz,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.band1Horz][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and go into second level -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.band1Horz)}">
			<xsl:call-template name="WrapCnf.a">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/>
			</xsl:call-template>
			</div>
		</xsl:when>
		<!-- horizontal band 2 -->
		<xsl:when test="substring($cnfRow,$i.cnfType.band2Horz,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.band2Horz][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and go into second level -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.band2Horz)}">
			<xsl:call-template name="WrapCnf.a">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/>
			</xsl:call-template>
			</div>
		</xsl:when>
		<!-- no horizontal band -->
		<xsl:otherwise>
			<!-- go into second level -->
			<xsl:call-template name="WrapCnf.a">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/><xsl:with-param name="prs.p" select="$prs.p"/><xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<xsl:template name="WrapCnf.a">
	<xsl:param name="p.tStyle"/><xsl:param name="cnfCol"/><xsl:param name="cnfRow"/>
	<xsl:param name="prs.pMany"/><xsl:param name="prs.p"/><xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- vertical band 1 -->
		<xsl:when test="substring($cnfCol,$i.cnfType.band1Vert,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.band1Vert][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and go into third level -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.band1Vert)}">
			<xsl:call-template name="WrapCnf.b">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/>
			</xsl:call-template>
			</div>
		</xsl:when>
		<!-- vertical band 2 -->
		<xsl:when test="substring($cnfCol,$i.cnfType.band2Vert,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.band2Vert][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and go into third level -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.band2Vert)}">
			<xsl:call-template name="WrapCnf.b">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/>
			</xsl:call-template>
			</div>
		</xsl:when>
		<!-- no vertical band -->
		<xsl:otherwise>
			<!-- no wrap and go into third level -->
			<xsl:call-template name="WrapCnf.b">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/><xsl:with-param name="prs.p" select="$prs.p"/><xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<xsl:template name="WrapCnf.b">
	<xsl:param name="p.tStyle"/><xsl:param name="cnfCol"/><xsl:param name="cnfRow"/>
	<xsl:param name="prs.pMany"/><xsl:param name="prs.p"/><xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- first column -->
		<xsl:when test="substring($cnfCol,$i.cnfType.firstCol,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.firstCol][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and go into forth level -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.firstCol)}">
			<xsl:call-template name="WrapCnf.c">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/>
			</xsl:call-template>
			</div>
		</xsl:when>
		<!-- last column -->
		<xsl:when test="substring($cnfCol,$i.cnfType.lastCol,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.lastCol][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and go into third level -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.lastCol)}">
			<xsl:call-template name="WrapCnf.c">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/>
			</xsl:call-template>
			</div>
		</xsl:when>
		<!-- neither first nor last column -->
		<xsl:otherwise>
			<!-- no wrap and go into third level -->
			<xsl:call-template name="WrapCnf.c">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/><xsl:with-param name="prs.p" select="$prs.p"/><xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<xsl:template name="WrapCnf.c">
	<xsl:param name="p.tStyle"/><xsl:param name="cnfCol"/><xsl:param name="cnfRow"/>
	<xsl:param name="prs.pMany"/><xsl:param name="prs.p"/><xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- first row -->
		<xsl:when test="substring($cnfRow,$i.cnfType.firstRow,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.firstRow][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and go into last level -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.firstRow)}">
			<xsl:call-template name="WrapCnf.d">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/>
			</xsl:call-template>
			</div>
		</xsl:when>
		<!-- last row -->
		<xsl:when test="substring($cnfRow,$i.cnfType.lastRow,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.lastRow][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and go into last level -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.lastRow)}">
			<xsl:call-template name="WrapCnf.d">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/>
			</xsl:call-template>
			</div>
		</xsl:when>
		<!-- neither first nor last row -->
		<xsl:otherwise>
			<!-- no wrap and go into last level -->
			<xsl:call-template name="WrapCnf.d">
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/><xsl:with-param name="prs.p" select="$prs.p"/><xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<xsl:template name="WrapCnf.d">
	<xsl:param name="p.tStyle"/><xsl:param name="cnfCol"/><xsl:param name="cnfRow"/>
	<xsl:param name="prs.pMany"/><xsl:param name="prs.p"/><xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- ne corner -->
		<xsl:when test="substring($cnfCol,$i.cnfType.neCell,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.neCell][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and display content -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.neCell)}">
			<xsl:call-template name="DisplayBodyContent"><xsl:with-param name="ns.content" select="*"/><xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/></xsl:call-template>
			</div>
		</xsl:when>
		<!-- nw corner -->
		<xsl:when test="substring($cnfCol,$i.cnfType.nwCell,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.nwCell][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and display content -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.nwCell)}">
			<xsl:call-template name="DisplayBodyContent"><xsl:with-param name="ns.content" select="*"/><xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/></xsl:call-template>
			</div>
		</xsl:when>
		<!-- se corner -->
		<xsl:when test="substring($cnfCol,$i.cnfType.seCell,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.seCell][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and display content -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.seCell)}">
			<xsl:call-template name="DisplayBodyContent"><xsl:with-param name="ns.content" select="*"/><xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/></xsl:call-template>
			</div>
		</xsl:when>
		<!-- sw corner -->
		<xsl:when test="substring($cnfCol,$i.cnfType.swCell,1)=$on">
			<xsl:variable name="p.cnfType" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.swCell][1]"/>
			<!-- updates -->
			<xsl:variable name="prs.p.updated">
				<xsl:call-template name="UpdatePPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.p" select="$prs.p"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.r.updated">
				<xsl:call-template name="UpdateRPr"><xsl:with-param name="p.style" select="$p.cnfType"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:variable>
			<xsl:variable name="prs.pMany.updated">
				<xsl:value-of select="$prs.pMany"/><xsl:for-each select="$p.cnfType"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
			</xsl:variable>
			<!-- wrap and display content -->
			<div class="{concat($p.tStyle/@w:styleId,'-',$cnfType.swCell)}">
			<xsl:call-template name="DisplayBodyContent"><xsl:with-param name="ns.content" select="*"/><xsl:with-param name="prs.pMany" select="$prs.pMany.updated"/><xsl:with-param name="prs.p" select="$prs.p.updated"/><xsl:with-param name="prs.r" select="$prs.r.updated"/></xsl:call-template>
			</div>
		</xsl:when>
		<!-- not a corner -->
		<xsl:otherwise>
			<!-- display content -->
			<xsl:call-template name="DisplayBodyContent"><xsl:with-param name="ns.content" select="*"/><xsl:with-param name="prs.pMany" select="$prs.pMany"/><xsl:with-param name="prs.p" select="$prs.p"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- get the specific property at all levels 
		used to apply all conditional formating for cell properties 
		many properties will be returned (one at each level)-->
<xsl:template name="GetCnfPr.all">
	<xsl:param name="type"/><xsl:param name="cnfCol"/><xsl:param name="cnfRow"/>
	<xsl:choose>
		<xsl:when test="substring($cnfRow,$i.cnfType.band1Horz,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band1Horz][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
		<xsl:when test="substring($cnfRow,$i.cnfType.band2Horz,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band2Horz][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="substring($cnfCol,$i.cnfType.band1Vert,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band1Vert][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.band2Vert,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band2Vert][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="substring($cnfCol,$i.cnfType.firstCol,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.firstCol][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.lastCol,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.lastCol][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="substring($cnfRow,$i.cnfType.firstRow,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.firstRow][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
		<xsl:when test="substring($cnfRow,$i.cnfType.lastRow,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.lastRow][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="substring($cnfCol,$i.cnfType.neCell,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.neCell][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.nwCell,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.nwCell][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.seCell,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.seCell][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.swCell,1)=$on">
			<xsl:for-each select="w:tStylePr[@w:type=$cnfType.swCell][1]">
				<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
			</xsl:for-each>
		</xsl:when>
	</xsl:choose>
</xsl:template>

<!-- grab the property in the highest existing conditional formatting level -->
<!-- it is used at the cell level, with both cnfRow and cnfCol -->
<!-- only one property will be returned -->
<xsl:template name="GetCnfPr.cell">
	<xsl:param name="type"/><xsl:param name="cnfCol"/><xsl:param name="cnfRow"/>
	<xsl:variable name="result1">
		<xsl:choose>
			<xsl:when test="substring($cnfCol,$i.cnfType.neCell,1)=$on">
				<xsl:for-each select="w:tStylePr[@w:type=$cnfType.neCell][1]">
					<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test="substring($cnfCol,$i.cnfType.nwCell,1)=$on">
				<xsl:for-each select="w:tStylePr[@w:type=$cnfType.nwCell][1]">
					<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test="substring($cnfCol,$i.cnfType.seCell,1)=$on">
				<xsl:for-each select="w:tStylePr[@w:type=$cnfType.seCell][1]">
					<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test="substring($cnfCol,$i.cnfType.swCell,1)=$on">
				<xsl:for-each select="w:tStylePr[@w:type=$cnfType.swCell][1]">
					<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
	</xsl:variable>
	<xsl:value-of select="$result1"/>
	<xsl:if test="$result1=''">
		<xsl:variable name="result2">
			<xsl:choose>
				<xsl:when test="substring($cnfRow,$i.cnfType.firstRow,1)=$on">
					<xsl:for-each select="w:tStylePr[@w:type=$cnfType.firstRow][1]">
						<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
					</xsl:for-each>
				</xsl:when>
				<xsl:when test="substring($cnfRow,$i.cnfType.lastRow,1)=$on">
					<xsl:for-each select="w:tStylePr[@w:type=$cnfType.lastRow][1]">
						<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
					</xsl:for-each>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:value-of select="$result2"/>
		<xsl:if test="$result2=''">
			<xsl:variable name="result3">
				<xsl:choose>
					<xsl:when test="substring($cnfCol,$i.cnfType.firstCol,1)=$on">
						<xsl:for-each select="w:tStylePr[@w:type=$cnfType.firstCol][1]">
							<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="substring($cnfCol,$i.cnfType.lastCol,1)=$on">
						<xsl:for-each select="w:tStylePr[@w:type=$cnfType.lastCol][1]">
							<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
						</xsl:for-each>
					</xsl:when>
				</xsl:choose>
			</xsl:variable>
			<xsl:value-of select="$result3"/>
			<xsl:if test="$result3=''">
				<xsl:variable name="result4">
					<xsl:choose>
						<xsl:when test="substring($cnfCol,$i.cnfType.band1Vert,1)=$on">
							<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band1Vert][1]">
								<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
							</xsl:for-each>
						</xsl:when>
						<xsl:when test="substring($cnfCol,$i.cnfType.band2Vert,1)=$on">
							<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band2Vert][1]">
								<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
							</xsl:for-each>
						</xsl:when>
					</xsl:choose>
				</xsl:variable>
				<xsl:value-of select="$result4"/>
				<xsl:if test="$result4=''">
					<xsl:choose>
						<xsl:when test="substring($cnfRow,$i.cnfType.band1Horz,1)=$on">
							<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band1Horz][1]">
								<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
							</xsl:for-each>
						</xsl:when>
						<xsl:when test="substring($cnfRow,$i.cnfType.band2Horz,1)=$on">
							<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band2Horz][1]">
								<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
							</xsl:for-each>
						</xsl:when>
					</xsl:choose>
				</xsl:if>
			</xsl:if>
		</xsl:if>
	</xsl:if>
</xsl:template>

<!-- grab the property in the highest existing conditional formatting level -->
<!-- it is used at the row level, with only cnfRow -->
<!-- only one property will be returned -->
<xsl:template name="GetCnfPr.row">
	<xsl:param name="type"/><xsl:param name="cnfRow"/>
	<xsl:variable name="result1">
		<xsl:choose>
			<xsl:when test="substring($cnfRow,$i.cnfType.firstRow,1)=$on">
				<xsl:for-each select="w:tStylePr[@w:type=$cnfType.firstRow][1]">
					<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test="substring($cnfRow,$i.cnfType.lastRow,1)=$on">
				<xsl:for-each select="w:tStylePr[@w:type=$cnfType.lastRow][1]">
					<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
	</xsl:variable>
	<xsl:value-of select="$result1"/>
	<xsl:if test="$result1=''">
		<xsl:choose>
			<xsl:when test="substring($cnfRow,$i.cnfType.band1Horz,1)=$on">
				<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band1Horz][1]">
					<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test="substring($cnfRow,$i.cnfType.band2Horz,1)=$on">
				<xsl:for-each select="w:tStylePr[@w:type=$cnfType.band2Horz][1]">
					<xsl:call-template name="GetCnfPr.a"><xsl:with-param name="type" select="$type"/></xsl:call-template>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
	</xsl:if>
</xsl:template>

<!-- the inner branch of GetCnfPr,
	return/perform the specified (by parameter type) property/action -->
<xsl:template name="GetCnfPr.a">
	<xsl:param name="type"/>
	<xsl:choose>
		<xsl:when test="$type = $t.applyTcPr">
			<xsl:call-template name="ApplyTcPr.class"/>
		</xsl:when>
		<xsl:when test="$type = $t.customCellpadding">
			<xsl:for-each select="w:tcPr[1]/w:tcMar[1]"><xsl:call-template name="ApplyCellMar"/></xsl:for-each>
		</xsl:when>
		<xsl:when test="$type = $t.defaultCellpadding">
			<xsl:for-each select="w:tblPr[1]/w:tblCellMar[1]"><xsl:call-template name="ApplyCellMar"/></xsl:for-each>
		</xsl:when>
		<xsl:when test="$type = $t.trCantSplit">
			<xsl:for-each select="w:trPr[1]/w:cantSplit[1]">
				<xsl:choose>
					<xsl:when test="@w:val = 'off'">page-break-inside:auto;</xsl:when>
					<xsl:otherwise>page-break-inside:avoid;</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</xsl:when>
	</xsl:choose>
</xsl:template>

<!-- return the highest priority conditional formatting type for the current cell -->
<xsl:template name="GetCnfType">
	<xsl:param name="cnfCol"/><xsl:param name="cnfRow"/>
	<xsl:choose>
		<xsl:when test="substring($cnfCol,$i.cnfType.neCell,1)=$on">
			<xsl:value-of select="$cnfType.neCell"/>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.nwCell,1)=$on">
			<xsl:value-of select="$cnfType.nwCell"/>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.seCell,1)=$on">
			<xsl:value-of select="$cnfType.seCell"/>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.swCell,1)=$on">
			<xsl:value-of select="$cnfType.swCell"/>
		</xsl:when>
		<xsl:when test="substring($cnfRow,$i.cnfType.firstRow,1)=$on">
			<xsl:value-of select="$cnfType.firstRow"/>
		</xsl:when>
		<xsl:when test="substring($cnfRow,$i.cnfType.lastRow,1)=$on">
			<xsl:value-of select="$cnfType.lastRow"/>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.firstCol,1)=$on">
			<xsl:value-of select="$cnfType.firstCol"/>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.lastCol,1)=$on">
			<xsl:value-of select="$cnfType.lastCol"/>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.band1Vert,1)=$on">
			<xsl:value-of select="$cnfType.band1Vert"/>
		</xsl:when>
		<xsl:when test="substring($cnfCol,$i.cnfType.band2Vert,1)=$on">
			<xsl:value-of select="$cnfType.band2Vert"/>
		</xsl:when>
		<xsl:when test="substring($cnfRow,$i.cnfType.band1Horz,1)=$on">
			<xsl:value-of select="$cnfType.band1Horz"/>
		</xsl:when>
		<xsl:when test="substring($cnfRow,$i.cnfType.band2Horz,1)=$on">
			<xsl:value-of select="$cnfType.band2Horz"/>
		</xsl:when>
	</xsl:choose>
</xsl:template>


<!-- BORDER HANDLING -->

<!-- go through the node-set and display groups of adj r that has the same borders -->
<xsl:template name="DisplayRBorder">
	<xsl:param name="ns.content" select="*"/>
	<xsl:param name="i.range.start" select="1"/>
	<xsl:param name="i.this" select="number($i.range.start)"/>
	<xsl:param name="pr.bdr.prev" select="''"/>
	<xsl:param name="b.bidi"/>
	<xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- if it is not the last element on the node-list -->
		<xsl:when test="($ns.content)[$i.this]">
			<xsl:for-each select="($ns.content)[$i.this]">
				<xsl:choose>
					<!-- ignore nodes that are not content by incrementing the position and call recursivly -->
					<xsl:when test="name() = 'w:proofErr' or (name() = 'aml:annotation' and not(@w:type = 'Word.Insertion'))">
						<xsl:call-template name="DisplayRBorder">
							<xsl:with-param name="ns.content" select="$ns.content"/>
							<xsl:with-param name="i.range.start" select="$i.range.start"/>
							<xsl:with-param name="i.this" select="$i.this+1"/>
							<xsl:with-param name="pr.bdr.prev" select="$pr.bdr.prev"/>
							<xsl:with-param name="b.bidi" select="$b.bidi"/>
							<xsl:with-param name="prs.r" select="$prs.r"/>
						</xsl:call-template>
					</xsl:when>	
					<xsl:otherwise>
						<!-- retrieve the border properties for current element -->
						<xsl:variable name="pr.bdr.this">
							<xsl:choose>
								<!-- annotation-insertion breaks off border -->
								<xsl:when test="name()='aml:annotation'"/>
								<xsl:otherwise>
									<!-- note that list index can have borders too -->
									<xsl:for-each select="descendant-or-self::*[name()='w:pPr' or name()='w:r'][1]">
										<xsl:for-each select="w:rPr[1]/w:bdr[1]">
											<xsl:call-template name="GetBorderPr"/>
										</xsl:for-each>
									</xsl:for-each>				
								</xsl:otherwise>
							</xsl:choose>
						</xsl:variable>
						<xsl:choose>
							<!-- if border of the previous element is the same as this element -->
							<xsl:when test="$pr.bdr.prev = $pr.bdr.this">
								<!-- continue recurisvely -->
								<xsl:call-template name="DisplayRBorder">
									<xsl:with-param name="ns.content" select="$ns.content"/>
									<xsl:with-param name="i.range.start" select="$i.range.start"/>
									<xsl:with-param name="i.this" select="$i.this+1"/>
									<xsl:with-param name="pr.bdr.prev" select="$pr.bdr.prev"/>
									<xsl:with-param name="b.bidi" select="$b.bidi"/>
									<xsl:with-param name="prs.r" select="$prs.r"/>
								</xsl:call-template>
							</xsl:when>
							<!-- if they are different -->
							<xsl:otherwise>
								<!-- wrap the previous group of elements under the same border and display the internal elements -->
								<xsl:call-template name="WrapRBorder">
									<xsl:with-param name="ns.content" select="$ns.content"/>
									<xsl:with-param name="i.bdrRange.start" select="$i.range.start"/>
									<xsl:with-param name="i.bdrRange.end" select="$i.this"/>
									<xsl:with-param name="pr.bdr" select="$pr.bdr.prev"/>
									<xsl:with-param name="b.bidi" select="$b.bidi"/>
									<xsl:with-param name="prs.r" select="$prs.r"/>
								</xsl:call-template>
								<!-- continue recurisvely, starting a new group -->
								<xsl:call-template name="DisplayRBorder">
									<xsl:with-param name="ns.content" select="$ns.content"/>
									<xsl:with-param name="i.range.start" select="$i.this"/>
									<xsl:with-param name="i.this" select="$i.this+1"/>
									<xsl:with-param name="pr.bdr.prev" select="$pr.bdr.this"/>
									<xsl:with-param name="b.bidi" select="$b.bidi"/>
									<xsl:with-param name="prs.r" select="$prs.r"/>
								</xsl:call-template>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>									<!-- set up the resulting record at the end -->
			</xsl:for-each>		
		</xsl:when>
		<!-- if it is the last element on the node-list -->
		<xsl:otherwise>
			<!-- display the last group of elements under the same border -->
			<xsl:call-template name="WrapRBorder">
				<xsl:with-param name="ns.content" select="$ns.content"/>
				<xsl:with-param name="i.bdrRange.start" select="$i.range.start"/>
				<xsl:with-param name="i.bdrRange.end" select="$i.this"/>
				<xsl:with-param name="pr.bdr" select="$pr.bdr.prev"/>
				<xsl:with-param name="b.bidi" select="$b.bidi"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- display/wrap the border around text, adj runs are grouped together under one border -->
<xsl:template name="WrapRBorder">
	<xsl:param name="ns.content"/>
	<xsl:param name="i.bdrRange.start"/>
	<xsl:param name="i.bdrRange.end"/>
	<xsl:param name="pr.bdr"/>
	<xsl:param name="b.bidi"/>
	<xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- no border around this group of r -->
		<xsl:when test="$pr.bdr = ''">
			<xsl:apply-templates select="($ns.content)[position() &gt;= $i.bdrRange.start and position() &lt; $i.bdrRange.end]">
				<xsl:with-param name="b.bidi" select="$b.bidi"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:apply-templates>
		</xsl:when>
		<!-- wrap border around this group -->
		<xsl:otherwise>
			<span>
			<xsl:attribute name="style">
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$pr.bdr"/></xsl:call-template>
			</xsl:attribute>
			<xsl:apply-templates select="($ns.content)[position() &gt;= $i.bdrRange.start and position() &lt; $i.bdrRange.end]">
				<xsl:with-param name="b.bidi" select="$b.bidi"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:apply-templates>
			</span>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- display body content, wrap borders around paragraphs if needed -->
<xsl:template name="DisplayPBorderOld">
	<xsl:param name="pr.frame.prev"/>
	<xsl:param name="pr.bdrTop.prev"/>
	<xsl:param name="pr.bdrLeft.prev"/>
	<xsl:param name="pr.bdrBottom.prev"/>
	<xsl:param name="pr.bdrRight.prev"/>
	<xsl:param name="pr.bdrBetween.prev"/>
	<xsl:param name="pr.bdrBar.prev"/>
	<xsl:param name="ns.content"/>
	<xsl:param name="i.range.start" select="1"/>
	<xsl:param name="i.this" select="number($i.range.start)"/>	<!-- keep track of the position -->
	<xsl:param name="prs.pMany"/>
	<xsl:param name="prs.p"/>
	<xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- if the current element exists -->
		<xsl:when test="($ns.content)[$i.this]">   
			<xsl:for-each select="($ns.content)[$i.this]">
				<xsl:variable name="pstyle">
					<xsl:call-template name="GetPStyleId"/>
				</xsl:variable>
				<xsl:variable name="p.pStyle" select="($ns.styles[@w:styleId=$pstyle])[1]"/>
				<!-- retrieve the frame and border properties for current element -->
				<xsl:variable name="pr.frame.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.frame"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>		
				</xsl:variable>
				<xsl:variable name="pr.bdrTop.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.top"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrLeft.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.left"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrBottom.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.bottom"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrRight.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.right"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrBetween.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.between"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrBar.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.bar"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:choose>
					<!-- if the frame and border properties are the same, continue recursively -->
					<xsl:when test="0 = 1 and $pr.frame.prev = $pr.frame.this and $pr.bdrTop.prev = $pr.bdrTop.this and $pr.bdrLeft.prev = $pr.bdrLeft.this and $pr.bdrBottom.prev = $pr.bdrBottom.this and $pr.bdrRight.prev = $pr.bdrRight.this and $pr.bdrBetween.prev = $pr.bdrBetween.this and $pr.bdrBar.prev = $pr.bdrBar.this">
						<xsl:call-template name="DisplayPBorder">
							<xsl:with-param name="ns.content" select="$ns.content"/>
							<xsl:with-param name="i.range.start" select="$i.range.start"/>
							<xsl:with-param name="i.this" select="$i.this+1"/>
							<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
							<xsl:with-param name="prs.p" select="$prs.p"/>
							<xsl:with-param name="prs.r" select="$prs.r"/>
							<xsl:with-param name="pr.frame.prev" select="$pr.frame.prev"/>
							<xsl:with-param name="pr.bdrTop.prev" select="$pr.bdrTop.prev"/>
							<xsl:with-param name="pr.bdrLeft.prev" select="$pr.bdrLeft.prev"/>
							<xsl:with-param name="pr.bdrBottom.prev" select="$pr.bdrBottom.prev"/>
							<xsl:with-param name="pr.bdrRight.prev" select="$pr.bdrRight.prev"/>
							<xsl:with-param name="pr.bdrBetween.prev" select="$pr.bdrBetween.prev"/>
							<xsl:with-param name="pr.bdrBar.prev" select="$pr.bdrBar.prev"/>	
						</xsl:call-template>
					</xsl:when>
					<!-- if they are different, wrap frame/border around and display the group -->
					<xsl:otherwise>
						<!-- wrap frame/border around the group and display the group of elements -->
						<xsl:call-template name="wrapFrame">
							<xsl:with-param name="ns.content" select="$ns.content"/>
							<xsl:with-param name="i.bdrRange.start" select="$i.range.start"/>
							<xsl:with-param name="i.bdrRange.end" select="$i.this"/>
							<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
							<xsl:with-param name="prs.p" select="$prs.p"/>
							<xsl:with-param name="prs.r" select="$prs.r"/>
							<xsl:with-param name="framePr" select="$pr.frame.prev"/>
							<xsl:with-param name="pr.bdrTop" select="$pr.bdrTop.prev"/>
							<xsl:with-param name="pr.bdrLeft" select="$pr.bdrLeft.prev"/>
							<xsl:with-param name="pr.bdrBottom" select="$pr.bdrBottom.prev"/>
							<xsl:with-param name="pr.bdrRight" select="$pr.bdrRight.prev"/>
							<xsl:with-param name="pr.bdrBetween" select="$pr.bdrBetween.prev"/>
							<xsl:with-param name="pr.bdrBar" select="$pr.bdrBar.prev"/>	
						</xsl:call-template>
						<!-- continue on with the rest of the elements recursively -->
						<xsl:call-template name="DisplayPBorder">
							<xsl:with-param name="ns.content" select="$ns.content"/>
							<xsl:with-param name="i.range.start" select="$i.this"/>
							<xsl:with-param name="i.this" select="$i.this+1"/>
							<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
							<xsl:with-param name="prs.p" select="$prs.p"/>
							<xsl:with-param name="prs.r" select="$prs.r"/>
							<xsl:with-param name="pr.frame.prev" select="$pr.frame.this"/>
							<xsl:with-param name="pr.bdrTop.prev" select="$pr.bdrTop.this"/>
							<xsl:with-param name="pr.bdrLeft.prev" select="$pr.bdrLeft.this"/>
							<xsl:with-param name="pr.bdrBottom.prev" select="$pr.bdrBottom.this"/>
							<xsl:with-param name="pr.bdrRight.prev" select="$pr.bdrRight.this"/>
							<xsl:with-param name="pr.bdrBetween.prev" select="$pr.bdrBetween.this"/>
							<xsl:with-param name="pr.bdrBar.prev" select="$pr.bdrBar.this"/>	
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</xsl:when>
		<!-- if the end of the node-list is reached, display the last group -->
		<xsl:otherwise>
			<xsl:call-template name="wrapFrame">
				<xsl:with-param name="ns.content" select="$ns.content"/>
				<xsl:with-param name="i.bdrRange.start" select="$i.range.start"/>
				<xsl:with-param name="i.bdrRange.end" select="$i.this"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
				<xsl:with-param name="prs.p" select="$prs.p"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
				<xsl:with-param name="framePr" select="$pr.frame.prev"/>
				<xsl:with-param name="pr.bdrTop" select="$pr.bdrTop.prev"/>
				<xsl:with-param name="pr.bdrLeft" select="$pr.bdrLeft.prev"/>
				<xsl:with-param name="pr.bdrBottom" select="$pr.bdrBottom.prev"/>
				<xsl:with-param name="pr.bdrRight" select="$pr.bdrRight.prev"/>
				<xsl:with-param name="pr.bdrBetween" select="$pr.bdrBetween.prev"/>
				<xsl:with-param name="pr.bdrBar" select="$pr.bdrBar.prev"/>	
			</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>				

<xsl:template name="DisplayPBorder">
	<xsl:param name="pr.frame.prev"/>
	<xsl:param name="pr.bdrTop.prev"/>
	<xsl:param name="pr.bdrLeft.prev"/>
	<xsl:param name="pr.bdrBottom.prev"/>
	<xsl:param name="pr.bdrRight.prev"/>
	<xsl:param name="pr.bdrBetween.prev"/>
	<xsl:param name="pr.bdrBar.prev"/>
	<xsl:param name="ns.content"/>
	<xsl:param name="i.range.start" select="1"/>
	<xsl:param name="i.this" select="number($i.range.start)"/>	<!-- keep track of the position -->
	<xsl:param name="prs.pMany"/>
	<xsl:param name="prs.p"/>
	<xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- if the current element exists -->
		<xsl:when test="($ns.content)[$i.this]">   
			<xsl:for-each select="($ns.content)"> <!-- [$i.this]">-->

				<xsl:variable name="pstyle">
					<xsl:call-template name="GetPStyleId"/>
				</xsl:variable>
				<xsl:variable name="p.pStyle" select="($ns.styles[@w:styleId=$pstyle])[1]"/>
				<!-- retrieve the frame and border properties for current element -->
				<xsl:variable name="pr.frame.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.frame"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>		
				</xsl:variable>
				<xsl:variable name="pr.bdrTop.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.top"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrLeft.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.left"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrBottom.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.bottom"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrRight.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.right"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrBetween.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.between"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
				<xsl:variable name="pr.bdrBar.this">
					<xsl:call-template name="GetPPr"><xsl:with-param name="type" select="$t.bdrPr.bar"/><xsl:with-param name="p.pStyle" select="$p.pStyle"/></xsl:call-template>
				</xsl:variable>
						<!-- wrap frame/border around the group and display the group of elements -->
						<xsl:call-template name="wrapFrame">
							<xsl:with-param name="ns.content" select="."/>
							<xsl:with-param name="i.bdrRange.start" select="1"/>
							<xsl:with-param name="i.bdrRange.end" select="2"/>
							<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
							<xsl:with-param name="prs.p" select="$prs.p"/>
							<xsl:with-param name="prs.r" select="$prs.r"/>
							<xsl:with-param name="framePr" select="$pr.frame.prev"/>
							<xsl:with-param name="pr.bdrTop" select="$pr.bdrTop.prev"/>
							<xsl:with-param name="pr.bdrLeft" select="$pr.bdrLeft.prev"/>
							<xsl:with-param name="pr.bdrBottom" select="$pr.bdrBottom.prev"/>
							<xsl:with-param name="pr.bdrRight" select="$pr.bdrRight.prev"/>
							<xsl:with-param name="pr.bdrBetween" select="$pr.bdrBetween.prev"/>
							<xsl:with-param name="pr.bdrBar" select="$pr.bdrBar.prev"/>	
						</xsl:call-template>
			</xsl:for-each>

		</xsl:when>
		<!-- if the end of the node-list is reached, display the last group -->
		<xsl:otherwise>
			<xsl:call-template name="wrapFrame">
				<xsl:with-param name="ns.content" select="$ns.content"/>
				<xsl:with-param name="i.bdrRange.start" select="$i.range.start"/>
				<xsl:with-param name="i.bdrRange.end" select="$i.this"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
				<xsl:with-param name="prs.p" select="$prs.p"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
				<xsl:with-param name="framePr" select="$pr.frame.prev"/>
				<xsl:with-param name="pr.bdrTop" select="$pr.bdrTop.prev"/>
				<xsl:with-param name="pr.bdrLeft" select="$pr.bdrLeft.prev"/>
				<xsl:with-param name="pr.bdrBottom" select="$pr.bdrBottom.prev"/>
				<xsl:with-param name="pr.bdrRight" select="$pr.bdrRight.prev"/>
				<xsl:with-param name="pr.bdrBetween" select="$pr.bdrBetween.prev"/>
				<xsl:with-param name="pr.bdrBar" select="$pr.bdrBar.prev"/>	
			</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>				

<!-- wrap frame around a range of paragraphs -->
<xsl:template name="wrapFrame">
	<xsl:param name="framePr"/>
	<xsl:param name="pr.bdrTop"/><xsl:param name="pr.bdrLeft"/><xsl:param name="pr.bdrBottom"/><xsl:param name="pr.bdrRight"/><xsl:param name="pr.bdrBetween"/><xsl:param name="pr.bdrBar"/>
	<xsl:param name="ns.content"/>
	<xsl:param name="i.bdrRange.start"/>
	<xsl:param name="i.bdrRange.end"/>
	<xsl:param name="prs.pMany"/>
	<xsl:param name="prs.p"/>
	<xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- try border if there is no frame -->
		<xsl:when test="$framePr = ''">
			<xsl:call-template name="wrapPBdr">
				<xsl:with-param name="ns.content" select="$ns.content"/>
				<xsl:with-param name="i.bdrRange.start" select="$i.bdrRange.start"/><xsl:with-param name="i.bdrRange.end" select="$i.bdrRange.end"/>
				<xsl:with-param name="pr.bdrTop" select="$pr.bdrTop"/><xsl:with-param name="pr.bdrLeft" select="$pr.bdrLeft"/><xsl:with-param name="pr.bdrBottom" select="$pr.bdrBottom"/><xsl:with-param name="pr.bdrRight" select="$pr.bdrRight"/><xsl:with-param name="pr.bdrBetween" select="$pr.bdrBetween"/><xsl:with-param name="pr.bdrBar" select="$pr.bdrBar"/>
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
				<xsl:with-param name="prs.p" select="$prs.p"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:when>
		<!-- wrap frame around the group -->
		<xsl:otherwise>
			<xsl:variable name="width" select="substring-before($framePr,$sepa2)"/><xsl:variable name="framePr1" select="substring-after($framePr,$sepa2)"/>
			<xsl:variable name="height" select="substring-before($framePr1,$sepa2)"/><xsl:variable name="framePr2" select="substring-after($framePr1,$sepa2)"/>
			<xsl:variable name="hrule" select="substring-before($framePr2,$sepa2)"/><xsl:variable name="framePr3" select="substring-after($framePr2,$sepa2)"/>
			<xsl:variable name="xalign" select="substring-before($framePr3,$sepa2)"/><xsl:variable name="framePr4" select="substring-after($framePr3,$sepa2)"/>
			<xsl:variable name="vspace" select="substring-before($framePr4,$sepa2)"/><xsl:variable name="framePr5" select="substring-after($framePr4,$sepa2)"/>
			<xsl:variable name="hspace" select="substring-before($framePr5,$sepa2)"/><xsl:variable name="framePr6" select="substring-after($framePr5,$sepa2)"/>	
			<xsl:variable name="wrap" select="substring-before($framePr6,$sepa2)"/>
			<!-- apply frame and a bunch of properties -->
			<table cellspacing="0" cellpadding="0" hspace="0" vspace="0">
			<xsl:if test="not($width = '' and $height='')">
				<xsl:attribute name="style">
					<xsl:if test="not($width = '')">width:<xsl:value-of select="$width div 20"/>pt;</xsl:if>
					<xsl:if test="not($height = '')">height:<xsl:value-of select="$height div 20"/>pt;</xsl:if>
				</xsl:attribute>
			</xsl:if>
			<xsl:attribute name="align">
				<xsl:choose>
					<xsl:when test="$xalign = 'right' or $xalign = 'outside'">right</xsl:when>
					<xsl:otherwise>left</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
			<tr><td valign="top" align="left">
			<xsl:attribute name="style">
				<xsl:text>padding:</xsl:text>
				<xsl:choose><xsl:when test="$vspace = ''">0</xsl:when><xsl:otherwise><xsl:value-of select="$vspace div 20"/>pt</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
				<xsl:choose><xsl:when test="$hspace = ''">0</xsl:when><xsl:otherwise><xsl:value-of select="$hspace div 20"/>pt</xsl:otherwise></xsl:choose><xsl:text>;</xsl:text>
			</xsl:attribute>
			<!-- try border within the frame -->
			<xsl:call-template name="wrapPBdr">
				<xsl:with-param name="ns.content" select="$ns.content"/>
				<xsl:with-param name="i.bdrRange.start" select="$i.bdrRange.start"/><xsl:with-param name="i.bdrRange.end" select="$i.bdrRange.end"/>
				<xsl:with-param name="pr.bdrTop" select="$pr.bdrTop"/><xsl:with-param name="pr.bdrLeft" select="$pr.bdrLeft"/><xsl:with-param name="pr.bdrBottom" select="$pr.bdrBottom"/><xsl:with-param name="pr.bdrRight" select="$pr.bdrRight"/><xsl:with-param name="pr.bdrBetween" select="$pr.bdrBetween"/><xsl:with-param name="pr.bdrBar" select="$pr.bdrBar"/>	
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
				<xsl:with-param name="prs.p" select="$prs.p"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
			</td></tr></table>
			<xsl:if test="$wrap = '' or $wrap = 'none' or $wrap = 'not-beside'"><br clear="all"/></xsl:if>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- display/wrap the border around paragraph(s), adj paragraphs are grouped together under one border -->
<xsl:template name="wrapPBdr">
	<xsl:param name="pr.bdrTop"/><xsl:param name="pr.bdrLeft"/><xsl:param name="pr.bdrBottom"/><xsl:param name="pr.bdrRight"/><xsl:param name="pr.bdrBetween"/><xsl:param name="pr.bdrBar"/>	
	<xsl:param name="ns.content"/>
	<xsl:param name="i.bdrRange.start"/>
	<xsl:param name="i.bdrRange.end"/>
	<xsl:param name="prs.pMany"/>
	<xsl:param name="prs.p"/>
	<xsl:param name="prs.r"/>
	<xsl:choose>
		<!-- if no border is defined, just display the range of elements -->
		<xsl:when test="$pr.bdrTop = '' and $pr.bdrLeft = '' and $pr.bdrBottom = '' and $pr.bdrRight = '' and $pr.bdrBar = ''">
			<xsl:apply-templates select="($ns.content)[position() &gt;= $i.bdrRange.start and position() &lt; $i.bdrRange.end]">
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
				<xsl:with-param name="prs.p" select="$prs.p"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
				<xsl:with-param name="pr.bdrBetween" select="$pr.bdrBetween"/>
			</xsl:apply-templates>
		</xsl:when>
		<!-- if any border is defined, wrap the group of elements with the borders -->
		<xsl:otherwise>
			<div>
			<!-- apply the borders -->
			<xsl:attribute name="style">			<!-- apply the borders -->
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$pr.bdrBar"/><xsl:with-param name="bdrSide" select="$bdrSide.left"/></xsl:call-template>
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$pr.bdrTop"/><xsl:with-param name="bdrSide" select="$bdrSide.top"/></xsl:call-template>
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$pr.bdrLeft"/><xsl:with-param name="bdrSide" select="$bdrSide.left"/></xsl:call-template>
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$pr.bdrBottom"/><xsl:with-param name="bdrSide" select="$bdrSide.bottom"/></xsl:call-template>
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$pr.bdrRight"/><xsl:with-param name="bdrSide" select="$bdrSide.right"/></xsl:call-template>
				<xsl:text>padding:</xsl:text>		<!-- apply the paddings -->
				<xsl:variable name="topPad" select="substring-before(substring-after(substring-after(substring-after($pr.bdrTop,$sepa2),$sepa2),$sepa2),$sepa2)"/>
				<xsl:variable name="rightPad" select="substring-before(substring-after(substring-after(substring-after($pr.bdrRight,$sepa2),$sepa2),$sepa2),$sepa2)"/>
				<xsl:variable name="bottomPad" select="substring-before(substring-after(substring-after(substring-after($pr.bdrBottom,$sepa2),$sepa2),$sepa2),$sepa2)"/>
				<xsl:variable name="leftPad" select="substring-before(substring-after(substring-after(substring-after($pr.bdrLeft,$sepa2),$sepa2),$sepa2),$sepa2)"/>
				<xsl:choose><xsl:when test="$topPad = ''">0</xsl:when><xsl:otherwise><xsl:value-of select="$topPad"/>pt</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
				<xsl:choose><xsl:when test="$rightPad = ''">0</xsl:when><xsl:otherwise><xsl:value-of select="$rightPad"/>pt</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
				<xsl:choose><xsl:when test="$bottomPad = ''">0</xsl:when><xsl:otherwise><xsl:value-of select="$bottomPad"/>pt</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
				<xsl:choose><xsl:when test="$leftPad = ''">0</xsl:when><xsl:otherwise><xsl:value-of select="$leftPad"/>pt</xsl:otherwise></xsl:choose><xsl:text>;</xsl:text>
			</xsl:attribute>
			<!--  display the range of elements -->
			<xsl:apply-templates select="($ns.content)[position() &gt;= $i.bdrRange.start and position() &lt; $i.bdrRange.end]">
				<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
				<xsl:with-param name="prs.p" select="$prs.p"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
				<xsl:with-param name="pr.bdrBetween" select="$pr.bdrBetween"/>
			</xsl:apply-templates>
			</div>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>


<!-- TRANSFORMATIONS -->

<!-- apply argument attributes to <script> or <applet> -->
<xsl:template name="ApplyArgs">
	<xsl:param name="value"/>
	<xsl:variable name="attributeName" select="normalize-space(substring-before($value,'='))"/>
	<xsl:variable name="afterName" select="concat(substring-after($value,'='),' ')"/>
	<xsl:if test="not($attributeName = '')">
		<xsl:attribute name="{$attributeName}"><xsl:value-of select="normalize-space(translate(substring-before($afterName,' '),'&quot;',' '))"/></xsl:attribute>
		<xsl:call-template name="ApplyArgs"><xsl:with-param name="value" select="normalize-space(substring-after($afterName,' '))"/></xsl:call-template>
	</xsl:if>
</xsl:template>

<!-- script -->
<xsl:template match="w:scriptAnchor">
	<script>
	<xsl:apply-templates select="*" mode="scriptAnchor"/>	
	</script>
</xsl:template>
<xsl:template match="w:args" mode="scriptAnchor">
	<xsl:call-template name="ApplyArgs"><xsl:with-param name="value" select="."/></xsl:call-template>
</xsl:template>
<xsl:template match="w:language" mode="scriptAnchor">
	<xsl:attribute name="language"><xsl:value-of select="."/></xsl:attribute>
</xsl:template>
<xsl:template match="w:scriptId" mode="scriptAnchor">
	<xsl:attribute name="id"><xsl:value-of select="."/></xsl:attribute>
</xsl:template>
<xsl:template match="w:scriptText" mode="scriptAnchor">
	<xsl:value-of disable-output-escaping="yes" select="."/>
</xsl:template>
<xsl:template match="*" mode="scriptAnchor"/>

<!-- applet -->
<xsl:template match="w:applet">
	<applet>
	<xsl:apply-templates select="*" mode="applet"/>
	</applet>
</xsl:template>
<xsl:template match="w:appletText" mode="applet">
	<xsl:value-of disable-output-escaping="yes" select="."/>
</xsl:template>
<xsl:template match="w:args" mode="applet">
	<xsl:call-template name="ApplyArgs"><xsl:with-param name="value" select="."/></xsl:call-template>
</xsl:template>
<xsl:template match="*" mode="applet"/>

<!-- text-box -->
<xsl:template match="w:txbxContent">
	<xsl:call-template name="DisplayBodyContent">
		<xsl:with-param name="ns.content" select="*"/>
	</xsl:call-template>
</xsl:template>

<!-- border hints -->

<xsl:template match="WX:pBdrGroup">
	<xsl:variable name="dxaLeft" select="WX:margin-left/@WX:val"/>
	<xsl:variable name="dxaRight" select="WX:margin-right/@WX:val"/>
	<xsl:variable name="ns.borders" select="WX:borders"/>

	<xsl:variable name="bdrStyles">
		<xsl:if test="$ns.borders/WX:top">
			<xsl:text>border-top:</xsl:text>
				<xsl:value-of select="$ns.borders/WX:top/@WX:val"/>
				<xsl:text> </xsl:text>
				<xsl:value-of select="$ns.borders/WX:top/@WX:bdrwidth div 20"/>
				<xsl:text>pt </xsl:text>
				<xsl:call-template name="ConvHexColor">
					<xsl:with-param name="value" select="$ns.borders/WX:top/@WX:color"/>
				</xsl:call-template>
				<xsl:text>;padding-top:</xsl:text>
				<xsl:value-of select="$ns.borders/WX:top/@WX:space"/>
				<xsl:text>pt</xsl:text>
		</xsl:if>
		<xsl:if test="$ns.borders/WX:bottom">
			<xsl:text>;border-bottom:</xsl:text>
				<xsl:value-of select="$ns.borders/WX:bottom/@WX:val"/>
				<xsl:text> </xsl:text>
				<xsl:value-of select="$ns.borders/WX:bottom/@WX:bdrwidth div 20"/>
				<xsl:text>pt </xsl:text>
				<xsl:call-template name="ConvHexColor">
					<xsl:with-param name="value" select="$ns.borders/WX:bottom/@WX:color"/>
				</xsl:call-template>
				<xsl:text>;padding-bottom:</xsl:text>
				<xsl:value-of select="$ns.borders/WX:bottom/@WX:space"/>
				<xsl:text>pt</xsl:text>
		</xsl:if>
		<xsl:if test="$ns.borders/WX:right">
			<xsl:text>;border-right:</xsl:text>
				<xsl:value-of select="$ns.borders/WX:right/@WX:val"/>
				<xsl:text> </xsl:text>
				<xsl:value-of select="$ns.borders/WX:right/@WX:bdrwidth div 20"/>
				<xsl:text>pt </xsl:text>
				<xsl:call-template name="ConvHexColor">
					<xsl:with-param name="value" select="$ns.borders/WX:right/@WX:color"/>
				</xsl:call-template>
				<xsl:text>;padding-right:</xsl:text>
				<xsl:value-of select="$ns.borders/WX:right/@WX:space"/>
				<xsl:text>pt</xsl:text>
		</xsl:if>
		<xsl:if test="$ns.borders/WX:left">
			<xsl:text>;border-left:</xsl:text>
				<xsl:value-of select="$ns.borders/WX:left/@WX:val"/>
				<xsl:text> </xsl:text>
				<xsl:value-of select="$ns.borders/WX:left/@WX:bdrwidth div 20"/>
				<xsl:text>pt </xsl:text>
				<xsl:call-template name="ConvHexColor">
					<xsl:with-param name="value" select="$ns.borders/WX:left/@WX:color"/>
				</xsl:call-template>
				<xsl:text>;padding-left:</xsl:text>
				<xsl:value-of select="$ns.borders/WX:left/@WX:space"/>
				<xsl:text>pt</xsl:text>
		</xsl:if>
		<xsl:if test="$dxaLeft">
			<xsl:text>;margin-left:</xsl:text>
			<xsl:value-of select="$dxaLeft div 20"/>
			<xsl:text>pt</xsl:text>
		</xsl:if>
		<xsl:if test="$dxaRight">
			<xsl:text>;margin-right:</xsl:text>
			<xsl:value-of select="$dxaRight div 20"/>
			<xsl:text>pt</xsl:text>
		</xsl:if>
		<xsl:if test="WX:shd">
			<xsl:text>;background-color:</xsl:text>
			<xsl:call-template name="ConvHexColor">
				<xsl:with-param name="value" select="WX:shd/@WX:bgcolor"/>
				<xsl:with-param name="autoColor" select="$autoColor.bg"/>
				<xsl:with-param name="transparentColor">transparent</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:variable>
	
	<xsl:choose>
		<xsl:when test="WX:apo">
			<table cellspacing="0" cellpadding="0" hspace="0" vspace="0">
				<xsl:choose>
					<xsl:when test="WX:apo/WX:jc/@WX:val">
						<xsl:attribute name="align"><xsl:value-of select="WX:apo/WX:jc/@WX:val"/></xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="align"><xsl:text>left</xsl:text></xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
				<xsl:attribute name="style">
					<xsl:if test="WX:apo/WX:width/@WX:val">
						<xsl:text>;width:</xsl:text>
						<xsl:value-of select="WX:apo/WX:width/@WX:val div 20"/>
						<xsl:text>pt</xsl:text>
					</xsl:if>
					<xsl:if test="WX:apo/WX:height/@WX:val">
						<xsl:text>;height:</xsl:text>
						<xsl:value-of select="WX:apo/WX:height/@WX:val div 20"/>
						<xsl:text>pt</xsl:text>
					</xsl:if>
				</xsl:attribute>
				<tr>
					<td valign="top" align="left">
						<xsl:attribute name="style">
							<xsl:if test="WX:apo/WX:vertFromText/@WX:val">
								<xsl:text>;padding-top:</xsl:text>
								<xsl:value-of select="WX:apo/WX:vertFromText/@WX:val div 20"/>
								<xsl:text>pt</xsl:text>
								<xsl:text>;padding-bottom:</xsl:text>
								<xsl:value-of select="WX:apo/WX:vertFromText/@WX:val div 20"/>
								<xsl:text>pt</xsl:text>
							</xsl:if>
							<xsl:if test="WX:apo/WX:horizFromText/@WX:val">
								<xsl:text>;padding-right:</xsl:text>
								<xsl:value-of select="WX:apo/WX:horizFromText/@WX:val div 20"/>
								<xsl:text>pt</xsl:text>
								<xsl:text>;padding-left:</xsl:text>
								<xsl:value-of select="WX:apo/WX:horizFromText/@WX:val div 20"/>
								<xsl:text>pt</xsl:text>
							</xsl:if>
						</xsl:attribute>						
						<div>
							<xsl:attribute name="style">
								<xsl:value-of select="$bdrStyles"/>
							</xsl:attribute>
							
	<!-- Now we're going to have another div to offset any margin changes made above
		 (note the negative signs below.  this allows us to have a properly indented border,
		 but still let's the paragraphs individually control their indents)  -->
							<div>
								<xsl:attribute name="style">
									<xsl:if test="$dxaLeft">
										<xsl:text>;margin-left:-</xsl:text>
										<xsl:value-of select="$dxaLeft div 20"/>
										<xsl:text>pt</xsl:text>
									</xsl:if>
									<xsl:if test="$dxaRight">
										<xsl:text>;margin-right:-</xsl:text>
										<xsl:value-of select="$dxaRight div 20"/>
										<xsl:text>pt</xsl:text>
									</xsl:if>
								</xsl:attribute>
								<xsl:call-template name="DisplayBodyContent">
									<xsl:with-param name="ns.content" select="*"/>
								</xsl:call-template>
							</div>
						</div>
					</td>
				</tr>
			</table>
		</xsl:when>
		<xsl:otherwise>
			<div>
				<xsl:attribute name="style">
					<xsl:value-of select="$bdrStyles"/>
				</xsl:attribute>
			
			<!-- Now we're going to have another div to offset any margin changes made above
			(note the negative signs below.  this allows us to have a properly indented border,
			but still let's the paragraphs individually control their indents)  -->
			<div>
				<xsl:attribute name="style">
					<xsl:if test="$dxaLeft">
						<xsl:text>;margin-left:-</xsl:text>
						<xsl:value-of select="$dxaLeft div 20"/>
						<xsl:text>pt</xsl:text>
					</xsl:if>
					<xsl:if test="$dxaRight">
						<xsl:text>;margin-right:-</xsl:text>
						<xsl:value-of select="$dxaRight div 20"/>
						<xsl:text>pt</xsl:text>
					</xsl:if>
				</xsl:attribute>
				
				<xsl:call-template name="DisplayBodyContent">
					<xsl:with-param name="ns.content" select="*"/>
				</xsl:call-template>
			</div>
		</div>
	</xsl:otherwise>
	</xsl:choose>
</xsl:template>
	
<!-- KL: need to handle pictures later on -->
<xsl:template match="w:pict">
	<xsl:apply-templates select="*"/>
</xsl:template>
	
	<!-- br -->
<xsl:template match="w:br">
	<br>
	<xsl:attribute name="clear">		<!-- @clear -->
		<xsl:choose>
			<xsl:when test="@w:clear"><xsl:value-of select="@w:clear"/></xsl:when>
			<xsl:otherwise>all</xsl:otherwise>
		</xsl:choose>
	</xsl:attribute>
	<xsl:if test="@w:type = 'page'">	<!-- page-break -->
		<xsl:attribute name="style">page-break-before:always</xsl:attribute>
	</xsl:if>
	</br>
</xsl:template>

<!-- InstrText -->
<xsl:template match="w:instrText">
</xsl:template>

<!-- del -->
<xsl:template match="w:delText">
	<del>
	<xsl:value-of select="."/>
	</del>
</xsl:template>

<xsl:template match="w:r//w:t[../w:rPr/WX:sym]">
	<xsl:variable name="p.SymHint" select="../w:rPr/WX:sym"/>

	<span><xsl:attribute name="style">font-family:<xsl:value-of select="$p.SymHint/@WX:font"/></xsl:attribute>
		<xsl:choose>
			<xsl:when test="starts-with($p.SymHint/@WX:char, 'F0')">
				<xsl:text disable-output-escaping="yes">&amp;</xsl:text>#x<xsl:value-of select="substring-after($p.SymHint/@WX:char, 'F0')"/><xsl:text>;</xsl:text>
			</xsl:when>
			<xsl:when test="starts-with($p.SymHint/@WX:char, 'f0')">
				<xsl:text disable-output-escaping="yes">&amp;</xsl:text>#x<xsl:value-of select="substring-after($p.SymHint/@WX:char, 'f0')"/><xsl:text>;</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text disable-output-escaping="yes">&amp;</xsl:text>#x<xsl:value-of select="$p.SymHint/@WX:char"/><xsl:text>;</xsl:text>
			</xsl:otherwise>
		</xsl:choose></span>

<!--	<span style="font-family:Arial;"><xsl:value-of select="local-name()"/></span>-->
</xsl:template>

<!-- text content -->
<xsl:template match="w:t">
	<xsl:value-of select="."/>
</xsl:template>

<!-- symbol -->
<xsl:template match="w:sym">
	<span><xsl:attribute name="style">font-family:<xsl:value-of select="@w:font"/></xsl:attribute>
		<xsl:choose>
			<xsl:when test="starts-with(@w:char, 'F0')">
				<xsl:text disable-output-escaping="yes">&amp;</xsl:text>#x<xsl:value-of select="substring-after(@w:char, 'F0')"/><xsl:text>;</xsl:text>
			</xsl:when>
			<xsl:when test="starts-with(@w:char, 'f0')">
				<xsl:text disable-output-escaping="yes">&amp;</xsl:text>#x<xsl:value-of select="substring-after(@w:char, 'f0')"/><xsl:text>;</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text disable-output-escaping="yes">&amp;</xsl:text>#x<xsl:value-of select="@w:char"/><xsl:text>;</xsl:text>
			</xsl:otherwise>
		</xsl:choose></span>
</xsl:template>

<xsl:template name="OutputTlcChar">
	<xsl:param name="count" select="0"/>
	<xsl:param name="tlc" select="' '"/>
	<xsl:value-of select="$tlc"/>
	<xsl:if test="$count > 1">
		<xsl:call-template name="OutputTlcChar">
			<xsl:with-param name="count" select="$count - 1"/>
			<xsl:with-param name="tlc" select="$tlc"/>
		</xsl:call-template>
	</xsl:if>
</xsl:template>
		
<xsl:template match="w:tab">
	<xsl:if test="@WX:cTlc">
		<xsl:call-template name="OutputTlcChar">
			<xsl:with-param name="tlc">
				<xsl:choose>
					<xsl:when test="@WX:tlc='dot'">
						<xsl:text>.</xsl:text>
					</xsl:when>
					<xsl:when test="@WX:tlc='hyphen'">
						<xsl:text>-</xsl:text>
					</xsl:when>
					<xsl:when test="@WX:tlc='underscore'">
						<xsl:text>_</xsl:text>
					</xsl:when>
					<xsl:when test="@WX:tlc='heavy'">
						<xsl:text>_</xsl:text>
					</xsl:when>
					<xsl:when test="@WX:tlc='middle-dot'">
						<xsl:text>&#183;</xsl:text>
					</xsl:when>
					<xsl:otherwise>
						<xsl:text>&#160;</xsl:text>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:with-param>
			<xsl:with-param name="count" select="@WX:cTlc"/>
		</xsl:call-template>
	</xsl:if>
</xsl:template>
						
<!-- soft-hyphen -->
<xsl:template match="w:softHyphen">
	<xsl:text>&#xAD;</xsl:text>
</xsl:template>

<!-- no-break-hyphen -->
<xsl:template match="w:noBreakHyphen">
	<xsl:text disable-output-escaping="yes">&amp;#8209;</xsl:text>
</xsl:template>

<!-- display items within an w:r (or list index at w:pPr) -->
<xsl:template name="DisplayRContent">
	<xsl:choose>
		<!-- display list numbering/bullet -->
		<xsl:when test="w:listPr">
			<xsl:value-of select="w:listPr[1]/WX:t/@WX:val"/>
			<!-- do the *after* tab... -->
			<xsl:if test="w:listPr[1]/WX:t/@WX:wTabAfter">
				<span style="font: 7pt 'Times New Roman';text-decoration:none;">
				<xsl:call-template name="OutputTlcChar">
					<xsl:with-param name="count" select="(w:listPr[1]/WX:t/@WX:wTabAfter div 30) - 0.50"/>
					<xsl:with-param name="tlc">&#160;</xsl:with-param>
				</xsl:call-template>
				</span>
			</xsl:if>
		</xsl:when>	
		<!-- otherwise, display r content items -->			 
		<xsl:otherwise>
			<xsl:apply-templates select="*"/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- apply special character properties as CSS styles
		should be called only once on the specific r -->
<xsl:template name="ApplyRPr.once">
	<xsl:param name="rStyleId"/>
	<xsl:param name="b.bidi"/>
	<xsl:param name="prs.r"/>
	<!-- complex script -->
	<xsl:variable name="b.complexScript">
		<xsl:choose>
			<xsl:when test="w:rPr[1]/w:cs[1] or w:rPr[1]/w:rtl[1]"><xsl:value-of select="$on"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$off"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:if test="$b.complexScript = $on">
		<xsl:variable name="suffix.complexScript">-CS</xsl:variable>
		<xsl:variable name="b.font-weight" select="substring($prs.r,$i.bcs,1)"/>
		<xsl:variable name="b.font-style" select="substring($prs.r,$i.ics,1)"/>
		<xsl:variable name="pr.sz" select="substring($prs.r,$i.szcs)"/>
		<!-- font-style CS -->
		<xsl:choose>
			<xsl:when test="$b.font-style = $on">font-style:italic;</xsl:when>
			<xsl:otherwise>font-style:normal;</xsl:otherwise>		
		</xsl:choose>
		<!-- font-weight CS -->
		<xsl:choose>
			<xsl:when test="$b.font-weight = $on">font-weight:bold;</xsl:when>
			<xsl:otherwise>font-weight:normal;</xsl:otherwise>		
		</xsl:choose>
		<!-- font-size CS -->
		<xsl:choose>
			<xsl:when test="$pr.sz = ''">font-size:12pt;</xsl:when>
			<xsl:otherwise>font-size:<xsl:value-of select="$pr.sz div 2"/>pt;</xsl:otherwise>		
		</xsl:choose>
	</xsl:if>
	<!-- direction ($b.bidi from pPr) -->
	<xsl:if test="not($b.bidi = '')">
		<xsl:choose>
			<xsl:when test="$b.bidi = $on and not($b.complexScript = $on)">direction:ltr;</xsl:when>
			<xsl:when test="not($b.bidi = $on) and $b.complexScript = $on">direction:rtl;</xsl:when>
		</xsl:choose>
	</xsl:if>
	<!-- color -->
	<xsl:if test="substring($prs.r,$i.emboss-imprint,1) = $on">color:gray;</xsl:if>
	<!-- text-decoration -->
	<xsl:variable name="b.line-through" select="substring($prs.r,$i.strike-dstrike,1)"/>
	<xsl:variable name="b.underline" select="substring($prs.r,$i.u-em,1)"/>
	<xsl:choose>
		<xsl:when test="$b.line-through = $off and $b.underline = $off">text-decoration:none;</xsl:when>
		<xsl:when test="$b.line-through = $on and $b.underline = $on">text-decoration:line-through underline;</xsl:when>
		<xsl:when test="$b.line-through = $on">text-decoration:none line-through;</xsl:when>
		<xsl:when test="$b.underline = $on">text-decoration:none underline;</xsl:when>
	</xsl:choose>
	<!-- vertical-align -->
	<xsl:variable name="b.sup" select="substring($prs.r,$i.sup,1)"/>
	<xsl:variable name="b.sub" select="substring($prs.r,$i.sub,1)"/>
	<xsl:choose>
		<xsl:when test="$b.sup = $on and $b.sub = $on">vertical-align:baseline;</xsl:when>
		<xsl:when test="$b.sub = $on">vertical-align:sub;</xsl:when>
		<xsl:when test="$b.sup = $on">vertical-align:super;</xsl:when>
	</xsl:choose>
	<!-- display -->
	<xsl:if test="not($rStyleId='CommentReference')">
		<xsl:if test="substring($prs.r,$i.vanish-webhidden,1) = $on">display:none;</xsl:if>
	</xsl:if>
</xsl:template>

<!-- apply all simple character properties as CSS styles
		that can also be implemented using CSS class selector -->
<xsl:template name="ApplyRPr.class">
	<xsl:for-each select="w:rPr[1]">
		<!-- background-color -->
		<xsl:choose>
			<xsl:when test="w:highlight">background-color:<xsl:call-template name="ConvColor"><xsl:with-param name="value" select="w:hightlight[1]/@w:val"/></xsl:call-template>;</xsl:when>
			<xsl:otherwise><xsl:for-each select="w:shd[1]"><xsl:call-template name="ApplyShd"/></xsl:for-each></xsl:otherwise>
		</xsl:choose>
		<!-- all other properties -->
		<xsl:apply-templates select="*" mode="rpr"/>
	</xsl:for-each>
</xsl:template>
	
<!-- rpr highlight -->
<xsl:template match="w:highlight" mode="rpr">background:<xsl:call-template name="ConvColor"><xsl:with-param name="value" select="@w:val"/></xsl:call-template>;</xsl:template>

<!-- rpr color -->
<xsl:template match="w:color" mode="rpr">color:<xsl:call-template name="ConvHexColor"><xsl:with-param name="value" select="@w:val"/></xsl:call-template>;</xsl:template>
<!-- rpr font-family -->
<xsl:template match="WX:font" mode="rpr">font-family:<xsl:value-of select="@WX:val"/>;</xsl:template>
<!-- rpr font-variant -->
<xsl:template match="w:smallCaps" mode="rpr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">font-variant:normal;</xsl:when>
		<xsl:otherwise>font-variant:small-caps;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- rpr layout-flow -->
<xsl:template match="w:asianLayout" mode="rpr">
	<xsl:choose>
		<xsl:when test="@w:vert = 'on'">layout-flow:horizontal;</xsl:when>
		<xsl:when test="@w:vert-compress = 'on'">layout-flow:horizontal;</xsl:when>
		<xsl:when test="@w:vert = 'off' or @w:vert-compress = 'off'">layout-flow:normal;</xsl:when>
	</xsl:choose>
	<xsl:if test="@w:combine = 'lines'">text-combine:lines;</xsl:if>
</xsl:template>
<!-- rpr letter-spacing -->
<xsl:template match="w:spacing" mode="rpr">letter-spacing:<xsl:value-of select="@w:val div 20"/>pt;</xsl:template>
<!-- rpr position -->
<xsl:template match="w:position" mode="rpr">
	<xsl:variable name="fDropCap">
		 <xsl:value-of select="ancestor::w:p[1]/w:pPr/w:framePr/@w:drop-cap"/>
	</xsl:variable>
	<xsl:if test="$fDropCap=''">
		<xsl:text>position:relative;top:</xsl:text>
		<xsl:value-of select="@w:val div -2"/>
		<xsl:text>pt;</xsl:text>
	</xsl:if>
</xsl:template>
<xsl:template match="w:fitText" mode="rpr">text-fit:<xsl:value-of select="@w:val div 20"/>pt;</xsl:template>
<xsl:template match="w:shadow" mode="rpr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">text-shadow:none;</xsl:when>
		<xsl:otherwise>text-shadow:0.2em 0.2em;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- rpr text-transform -->
<xsl:template match="w:caps" mode="rpr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">text-transform:none;</xsl:when>
		<xsl:otherwise>text-transform:uppercase;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- rpr font-size -->
<xsl:template match="w:sz" mode="rpr">font-size:<xsl:value-of select="@w:val div 2"/>pt;</xsl:template>
<!-- rpr font-weight -->
<xsl:template match="w:b" mode="rpr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">font-weight:normal;</xsl:when>
		<xsl:otherwise>font-weight:bold;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- rpr font-style -->
<xsl:template match="w:i" mode="rpr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">font-style:normal;</xsl:when>
		<xsl:otherwise>font-style:italic;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- rpr close mode -->
<xsl:template match="*" mode="rpr"/>

<!-- display context w:r (both properties and content) -->
<!-- calling on w:pPr is used to display the list index -->
<xsl:template name="DisplayR">
	<xsl:param name="b.bidi"/>
	<xsl:param name="prs.r"/>
	<!-- get style id -->
	<xsl:variable name="rStyleId" select="string(w:rPr/wrStyle/@w:val)"/>
	<!-- update prs.r -->
	<xsl:variable name="prs.r.updated">
		<!-- update at rstyle rPr -->
		<xsl:variable name="prs.r.updated1">
			<xsl:call-template name="UpdateRPr">
				<xsl:with-param name="p.style" select="($ns.styles[@w:styleId=$rStyleId])[1]"/>
				<xsl:with-param name="prs.r" select="$prs.r"/>
			</xsl:call-template>
		</xsl:variable>
		<!-- update at direct rPr -->
		<xsl:variable name="prs.r.updated2">
			<xsl:call-template name="UpdateRPr">
				<xsl:with-param name="prs.r" select="$prs.r.updated1"/>
			</xsl:call-template>
		</xsl:variable>
		<!-- update at list index rPr -->
		<xsl:variable name="prs.r.temp3">
			<xsl:call-template name="GetListPr">
				<xsl:with-param name="type" select="$t.updateRPr"/>
				<xsl:with-param name="prs.r" select="$prs.r.updated2"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$prs.r.temp3=''">
				<xsl:value-of select="$prs.r.updated2"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$prs.r.temp3"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<!-- list suff property is used to determine the spacing of list index -->
	<xsl:variable name="pr.listSuff">
		<xsl:call-template name="GetListPr">
			<xsl:with-param name="type" select="$t.listSuff"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:variable name="styleMod">
		<xsl:call-template name="ApplyRPr.class"/>
		<xsl:if test="w:listPr">
			<xsl:text>font-style:normal;text-decoration:none;font-weight:normal;background:transparent;</xsl:text>
		</xsl:if>
		<!-- apply font-family of list index -->
		<xsl:apply-templates select="w:listPr[1]/WX:font[1]" mode="rpr"/>
		<!-- call ApplyRPr.class on listPr's rPr -->
		<xsl:call-template name="GetListPr">
			<xsl:with-param name="type" select="$t.applyRPr"/>
		</xsl:call-template>
		<xsl:call-template name="ApplyRPr.once">
			<xsl:with-param name="rStyleId" select="$rStyleId"/>
			<xsl:with-param name="b.bidi" select="$b.bidi"/>
			<xsl:with-param name="prs.r" select="$prs.r.updated"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:choose>
		<xsl:when test="$rStyleId='' and $styleMod=''">
			<xsl:call-template name="DisplayRContent"/>
			<!-- when list suff is set to space, add a space after list numbering/bullet -->
			<xsl:if test="$pr.listSuff = $pr.listSuff.space"><xsl:text> </xsl:text></xsl:if>
		</xsl:when>
		<xsl:otherwise>
			<span>
			<!-- class attribute -->
			<xsl:if test="not($rStyleId='')">
				<xsl:attribute name="class"><xsl:value-of select="$rStyleId"/><xsl:value-of select="$styleSuffix.char"/></xsl:attribute>
			</xsl:if>
			<!-- style attribute -->
			<xsl:if test="not($styleMod='')">
					<xsl:attribute name="style"><xsl:value-of select="$styleMod"/></xsl:attribute>
			</xsl:if>
			<xsl:call-template name="DisplayRContent"/>
			<!-- when list suff is set to space, add a space after list numbering/bullet -->
			<xsl:if test="$pr.listSuff = $pr.listSuff.space"><xsl:text> </xsl:text></xsl:if>
			</span>
		</xsl:otherwise>	
	</xsl:choose>
</xsl:template>

<!-- match template for w:r, call "DisplayR" -->
<xsl:template match="w:r">
	<xsl:param name="b.bidi" select="''"/>
	<xsl:param name="prs.r" select="$prs.r.default"/>
	<xsl:call-template name="DisplayR">
		<xsl:with-param name="b.bidi" select="$b.bidi"/>
		<xsl:with-param name="prs.r" select="$prs.r"/>
	</xsl:call-template>
</xsl:template>

<!-- match template for w:pPr, call "DisplayR" -->
<xsl:template match="w:pPr">
	<xsl:param name="b.bidi" select="''"/>
	<xsl:param name="prs.r" select="$prs.r.default"/>
	<xsl:call-template name="DisplayR">
		<xsl:with-param name="b.bidi" select="$b.bidi"/>
		<xsl:with-param name="prs.r" select="$prs.r"/>
	</xsl:call-template>
</xsl:template>

<!-- display hyper-link -->
<xsl:template name="DisplayHlink">
	<xsl:param name="b.bidi"/>
	<xsl:param name="prs.r"/>
	<a>
	<xsl:variable name="href">
		<xsl:for-each select="@w:dest"><xsl:value-of select="."/></xsl:for-each>
		<xsl:choose>
			<xsl:when test="@w:bookmark">#<xsl:value-of select="@w:bookmark"/></xsl:when>
			<xsl:when test="@w:arbLocation"># <xsl:value-of select="@w:arbLocation"/></xsl:when>
		</xsl:choose>
	</xsl:variable>
	<xsl:if test="not(href='')"><xsl:attribute name="href"><xsl:value-of select="$href"/></xsl:attribute></xsl:if>
	<xsl:for-each select="@w:target"><xsl:attribute name="target"><xsl:value-of select="."/></xsl:attribute></xsl:for-each>
	<xsl:for-each select="@w:screenTip"><xsl:attribute name="title"><xsl:value-of select="."/></xsl:attribute></xsl:for-each>
	<xsl:call-template name="DisplayPContent"><xsl:with-param name="b.bidi" select="$b.bidi"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
	</a>
</xsl:template>

<!-- match template for w:hlink, call "DisplayHlink" -->
<xsl:template match="w:hlink">
	<xsl:param name="b.bidi" select="''"/>
	<xsl:param name="prs.r" select="$prs.r.default"/>
	<xsl:call-template name="DisplayHlink">
		<xsl:with-param name="b.bidi" select="$b.bidi"/>
		<xsl:with-param name="prs.r" select="$prs.r"/>
	</xsl:call-template>
</xsl:template>

<!-- apply sepcial paragraph properties
		needs to be called once at the paragraph level -->
<xsl:template name="ApplyPPr.once">
	<xsl:param name="i.bdrRange.this"/>
	<xsl:param name="i.bdrRange.last"/>
	<xsl:param name="pr.bdrBetween"/>
	<xsl:param name="prs.p"/>
	<xsl:param name="b.bidi"/>
	<!-- border-bottom -->
	<xsl:if test="not($i.bdrRange.this = $i.bdrRange.last)">
		<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$pr.bdrBetween"/><xsl:with-param name="bdrSide" select="$bdrSide.bottom"/></xsl:call-template>
	</xsl:if>
	<!-- padding -->
	<xsl:if test="not($pr.bdrBetween = '')">
		<xsl:choose>
			<xsl:when test="$i.bdrRange.this = 1">padding:0 0 1pt;</xsl:when>
			<xsl:when test="$i.bdrRange.this = i.bdrRange.last">padding:1pt 0 0;</xsl:when>
			<xsl:otherwise>padding:1pt 0 1pt;</xsl:otherwise>
		</xsl:choose>
	</xsl:if>
	<!-- direction, unicode-bidi -->
	<xsl:choose>
		<xsl:when test="$b.bidi = $off">direction:ltr;unicode-bidi:normal;</xsl:when>
		<xsl:when test="$b.bidi = $on">direction:rtl;unicode-bidi:embed;text-align:right;</xsl:when>
	</xsl:choose>
	<!-- margin-right, margin-left, text-indent -->
	<xsl:variable name="pr.ind" select="substring($prs.p,$i.ind)"/>
	<xsl:variable name="pr.listInd">
		<xsl:for-each select="w:pPr">
			<xsl:call-template name="GetListPr">
				<xsl:with-param name="type" select="$t.listInd"/>
			</xsl:call-template>
		</xsl:for-each>
	</xsl:variable>
	<xsl:if test="not($pr.ind='' and $pr.listInd='')">
		<!-- sort out all the indent and margin properties -->
		<xsl:variable name="pr.ind.left" select="substring-before($pr.ind,$sepa2)"/><xsl:variable name="temp1" select="substring-after($pr.ind,$sepa2)"/>
		<xsl:variable name="pr.ind.leftChars" select="substring-before($temp1,$sepa2)"/><xsl:variable name="temp2" select="substring-after($temp1,$sepa2)"/>
		<xsl:variable name="pr.ind.right" select="substring-before($temp2,$sepa2)"/><xsl:variable name="temp3" select="substring-after($temp2,$sepa2)"/>
		<xsl:variable name="pr.ind.rightChars" select="substring-before($temp3,$sepa2)"/><xsl:variable name="temp4" select="substring-after($temp3,$sepa2)"/>
		<xsl:variable name="pr.ind.hanging" select="substring-before($temp4,$sepa2)"/><xsl:variable name="temp5" select="substring-after($temp4,$sepa2)"/>
		<xsl:variable name="pr.ind.hangingChars" select="substring-before($temp5,$sepa2)"/><xsl:variable name="temp6" select="substring-after($temp5,$sepa2)"/>
		<xsl:variable name="pr.ind.firstLine" select="substring-before($temp6,$sepa2)"/>
		<xsl:variable name="pr.ind.firstLineChars" select="substring-after($temp6,$sepa2)"/>
		<xsl:variable name="pr.listInd.left" select="substring-before($pr.listInd,$sepa2)"/><xsl:variable name="temp1a" select="substring-after($pr.listInd,$sepa2)"/>
		<xsl:variable name="pr.listInd.leftChars" select="substring-before($temp1a,$sepa2)"/><xsl:variable name="temp2a" select="substring-after($temp1a,$sepa2)"/>
		<xsl:variable name="pr.listInd.hanging" select="substring-before($temp2a,$sepa2)"/>
		<xsl:variable name="pr.listInd.hangingChars" select="substring-after($temp2a,$sepa2)"/>	
		<!-- figure out if which side is left depending on bidi -->
		<xsl:variable name="marginSide.before">margin-<xsl:choose><xsl:when test="$b.bidi=$on">right</xsl:when><xsl:otherwise>left</xsl:otherwise></xsl:choose>:</xsl:variable>
		<xsl:variable name="marginSide.after">margin-<xsl:choose><xsl:when test="$b.bidi=$on">left</xsl:when><xsl:otherwise>right</xsl:otherwise></xsl:choose>:</xsl:variable>
		<!-- before margin (left, or right for bidi) -->
		<xsl:choose>
			<!-- paragraph -->
			<xsl:when test="not($pr.ind.left = '')"><xsl:value-of select="$marginSide.before"/><xsl:value-of select="$pr.ind.left div 20"/>pt;</xsl:when>
			<xsl:when test="not($pr.ind.leftChars = '' and $pr.ind.hangingChars='')">
				<xsl:value-of select="$marginSide.before"/>
				<xsl:variable name="leftchars"><xsl:choose><xsl:when test="$pr.ind.leftChars=''">0</xsl:when><xsl:otherwise><xsl:value-of select="$pr.ind.leftChars div 100"/></xsl:otherwise></xsl:choose></xsl:variable>
				<xsl:variable name="hangingchars"><xsl:choose><xsl:when test="$pr.ind.hangingChars=''">0</xsl:when><xsl:otherwise><xsl:value-of select="$pr.ind.hangingChars div 100"/></xsl:otherwise></xsl:choose></xsl:variable>
				<xsl:value-of select="$leftchars + $hangingchars"/>
				<xsl:text>em;</xsl:text>
			</xsl:when>
			<!-- if paragraph margin not defined, then list -->
			<xsl:when test="not($pr.listInd.left = '')"><xsl:value-of select="$marginSide.before"/><xsl:value-of select="$pr.listInd.left div 20"/>pt;</xsl:when>
			<xsl:when test="not($pr.listInd.leftChars = '' and $pr.listInd.hangingChars='')">
				<xsl:value-of select="$marginSide.before"/>
				<xsl:variable name="leftchars"><xsl:choose><xsl:when test="$pr.listInd.leftChars=''">0</xsl:when><xsl:otherwise><xsl:value-of select="$pr.listInd.leftChars div 100 * 12"/></xsl:otherwise></xsl:choose></xsl:variable>
				<xsl:variable name="hangingchars"><xsl:choose><xsl:when test="$pr.listInd.hangingChars=''">0</xsl:when><xsl:otherwise><xsl:value-of select="$pr.listInd.hangingChars div 100 * 12"/></xsl:otherwise></xsl:choose></xsl:variable>
				<xsl:value-of select="$leftchars + $hangingchars"/>
				<xsl:text>pt;</xsl:text>
			</xsl:when>
		</xsl:choose>
		<!-- after margin (right, or left for bidi) -->
		<xsl:choose>
			<xsl:when test="not($pr.ind.right = '')"><xsl:value-of select="$marginSide.after"/><xsl:value-of select="$pr.ind.right div 20"/>pt;</xsl:when>
			<xsl:when test="not($pr.ind.rightChars = '')"><xsl:value-of select="$marginSide.after"/><xsl:value-of select="$pr.ind.rightChars div 100"/>em;</xsl:when>
		</xsl:choose>
		<!-- text-indent -->
		<xsl:choose>
			<xsl:when test="not($pr.ind.hanging='')">text-indent:<xsl:value-of select="$pr.ind.hanging div -20"/>pt;</xsl:when>
			<xsl:when test="not($pr.ind.hangingChars='')">text-indent:<xsl:value-of select="$pr.ind.hangingChars div -100"/>em;</xsl:when>
			<xsl:when test="not($pr.ind.firstLine='')">text-indent:<xsl:value-of select="$pr.ind.firstLine div 20"/>pt;</xsl:when>
			<xsl:when test="not($pr.ind.firstLineChars='')">text-indent:<xsl:value-of select="$pr.ind.firstLineChars div 100"/>em;</xsl:when>
			<xsl:when test="not($pr.listInd.hanging='')">text-indent:<xsl:value-of select="$pr.listInd.hanging div -20"/>pt;</xsl:when>
			<xsl:when test="not($pr.listInd.hangingChars='')">text-indent:<xsl:value-of select="$pr.listInd.hangingChars div -100 * 12"/>pt;</xsl:when>
		</xsl:choose>
	</xsl:if>
	<!-- text-autospace -->
	<xsl:variable name="b.textAutospace.o" select="substring($prs.p,$i.textAutospace.o,1)"/>
	<xsl:variable name="b.textAutospace.n" select="substring($prs.p,$i.textAutospace.n,1)"/>
	<xsl:choose>
		<xsl:when test="not($b.textAutospace.n = $off) and $b.textAutospace.o = $off">text-autospace:ideograph-numeric;</xsl:when>
		<xsl:when test="not($b.textAutospace.o = $off) and $b.textAutospace.n = $off">text-autospace:ideograph-other;</xsl:when>
		<xsl:when test="$b.textAutospace.o = $off and $b.textAutospace.n = $off">text-autospace:none;</xsl:when>
	</xsl:choose>
</xsl:template>

<!-- apply paragraph properties as non-inheriting CSS
		this template needs to be called once for each inheritance level -->
<xsl:template name="ApplyPPr.many">
	<xsl:param name="t.cSpacing" select="$t.cSpacing.all"/>
	<!-- margin-top, margin-bottom -->
	<xsl:variable name="spacing" select="w:pPr[1]/w:spacing[1]"/>
	<xsl:choose>
		<xsl:when test="($spacing/@w:before-autospacing and not($spacing/@w:before-autospacing = 'off')) or $t.cSpacing = $t.cSpacing.none or $t.cSpacing = $t.cSpacing.bottom">
<!--	I've removed this (rlittle) - - if we *do* have an autospacing value *and* its not 'off', then we don't want a margin value written...
			<xsl:text>margin-top:</xsl:text>
				<xsl:value-of select="$pMargin.default.top"/>
			<xsl:text>;</xsl:text>
-->
		</xsl:when>
		<xsl:when test="$spacing/@w:before">margin-top:<xsl:value-of select="$spacing/@w:before div 20"/>pt;</xsl:when>
		<xsl:when test="$spacing/@w:before-lines">margin-top:<xsl:value-of select="$spacing/@w:before-lines *.12"/>pt;</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="($spacing/@w:after-autospacing and not($spacing/@w:after-autospacing = 'off')) or $t.cSpacing = $t.cSpacing.none or $t.cSpacing = $t.cSpacing.top"> 
<!--	I've removed this (rlittle) - - if we *do* have an autospacing value *and* its not 'off', then we don't want a margin value written...
			<xsl:text>margin-bottom:</xsl:text>
				<xsl:value-of select="$pMargin.default.bottom"/>
			<xsl:text>;</xsl:text>
-->
		</xsl:when>

		<xsl:when test="$spacing/@w:after">margin-bottom:<xsl:value-of select="$spacing/@w:after div 20"/>pt;</xsl:when>
		<xsl:when test="$spacing/@w:after-lines">margin-bottom:<xsl:value-of select="$spacing/@w:after-lines *.12"/>pt;</xsl:when>
	</xsl:choose>
	<xsl:for-each select="w:pPr[1]">
		<!-- layout-grid-mode -->
		<xsl:for-each select="w:snapToGrid[1]">
			<xsl:choose>
				<xsl:when test="@w:val = 'off'">layout-grid-mode:char;</xsl:when>
				<xsl:otherwise>layout-grid-mode:both;</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
		<!-- page-break-after -->
		<xsl:for-each select="w:keepNext[1]">
			<xsl:choose>
				<xsl:when test="@w:val = 'off'">page-break-after:auto;</xsl:when>
				<xsl:otherwise>page-break-after:avoid;</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
		<!-- page-break-before -->
		<xsl:for-each select="w:pageBreakBefore[1]">
			<xsl:choose>
				<xsl:when test="@w:val = 'off'">page-break-before:auto;</xsl:when>
				<xsl:otherwise>page-break-before:always;</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
	</xsl:for-each>
</xsl:template>

<!-- apply paragraph properties as CSS
		these properties inherit in CSS so they can be used by classes -->
<xsl:template name="ApplyPPr.class">
	<xsl:apply-templates select="w:pPr[1]/*" mode="ppr"/>
</xsl:template>

<!-- ppr background-color -->
<xsl:template match="w:shd" mode="ppr"><xsl:call-template name="ApplyShd"/></xsl:template>

<!-- background-color hint -->
<xsl:template match="WX:shd" mode="ppr"><xsl:call-template name="ApplyShdHint"/></xsl:template>

<!-- ppr layout-flow -->
<xsl:template match="w:textDirection" mode="ppr"><xsl:call-template name="ApplyTextDirection"/></xsl:template>
<!-- ppr line-height -->
<xsl:template match="w:spacing[@w:line-rule or @w:line]" mode="ppr">
	<xsl:choose>
		<xsl:when test="not(@w:line-rule)">line-height:<xsl:value-of select="@w:line div 240"/>;</xsl:when>
		<xsl:when test="@w:line-rule = 'auto'">line-height:normal;</xsl:when>
		<xsl:otherwise>line-height:<xsl:value-of select="@w:line div 20"/>pt;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- ppr punctuation-trim -->
<xsl:template match="w:topLinePunct" mode="ppr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">punctuation-trim:none;</xsl:when>
		<xsl:otherwise>punctuation-trim:leading;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- ppr punctuation-wrap -->
<xsl:template match="w:overflowPunct" mode="ppr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">punctuation-wrap:simple;</xsl:when>
		<xsl:otherwise>punctuation-wrap:hanging;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- ppr text-align -->
<xsl:template match="w:jc" mode="ppr">
	<xsl:choose>
		<xsl:when test="@w:val = 'left'">text-align:left;</xsl:when>
		<xsl:when test="@w:val = 'center'">text-align:center;</xsl:when>
		<xsl:when test="@w:val = 'right'">text-align:right;</xsl:when>
		<xsl:when test="@w:val = 'both'">text-align:justify;text-justify:inter-ideograph;</xsl:when>
		<xsl:when test="@w:val = 'distribute'">text-align:justify;text-justify:distribute-all-lines;</xsl:when>
		<xsl:when test="@w:val = 'low-kashida'">text-align:justify;text-justify:kashida;text-kashida:0%;</xsl:when>
		<xsl:when test="@w:val = 'medium-kashida'">text-align:justify;text-justify:kashida;text-kashida:10%;</xsl:when>
		<xsl:when test="@w:val = 'high-kashida'">text-align:justify;text-justify:kashida;text-kashida:20%;</xsl:when>
		<xsl:when test="@w:val = 'thai-distribute'">text-align:justify;text-justify:inter-cluster;</xsl:when>
	</xsl:choose>
</xsl:template>
<!-- ppr vertical-align -->
<xsl:template match="w:textAlignment" mode="ppr">
	<xsl:choose>
		<xsl:when test="@w:val = 'top'">vertical-align:top;</xsl:when>
		<xsl:when test="@w:val = 'center'">vertical-align:middle;</xsl:when>
		<xsl:when test="@w:val = 'baseline'">vertical-align:baseline;</xsl:when>
		<xsl:when test="@w:val = 'bottom'">vertical-align:bottom;</xsl:when>
		<xsl:when test="@w:val = 'auto'">vertical-align:baseline;</xsl:when>
	</xsl:choose>
</xsl:template>
<!-- word-break -->
<xsl:template match="w:wordWrap" mode="ppr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">word-break:break-all;</xsl:when>
		<xsl:otherwise>word-break:normal;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- ppr mode close -->
<xsl:template match="*" mode="ppr"/>

<!-- display the elements within w:p -->
<xsl:template name="DisplayPContent">
	<xsl:param name="b.bidi"/>
	<xsl:param name="prs.r"/>
	<xsl:call-template name="DisplayRBorder">
		<xsl:with-param name="b.bidi" select="$b.bidi"/>
		<xsl:with-param name="prs.r" select="$prs.r"/>
	</xsl:call-template>
	<!-- &nbsp; -->
	<xsl:if test="count(*[not(name()='w:pPr')])=0"><xsl:text disable-output-escaping="yes">&#160;</xsl:text></xsl:if>
</xsl:template>

<xsl:template name="GetPStyleId">
	<xsl:choose>
		<xsl:when test="w:pPr/w:pStyle/@w:val">
			<xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="$pStyleId.default"/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- display <p> with CSS class and style attributes -->
<xsl:template match="w:p">
	<xsl:param name="bdrBetween" select="''"/>	
	<xsl:param name="prs.pMany" select="''"/>
	<xsl:param name="prs.p" select="$prs.p.default"/>
	<xsl:param name="prs.r" select="$prs.r.default"/>
	<xsl:if test="not(w:pPr/w:pStyle/@w:val='z-TopofForm') and not(w:pPr/w:pStyle/@w:val='z-BottomofForm')">
	<p>
	<!-- get and apply the class name -->
	<xsl:variable name="pStyleId">
		<xsl:call-template name="GetPStyleId"/>
	</xsl:variable>
	<xsl:attribute name="class"><xsl:value-of select="$pStyleId"/><xsl:value-of select="$styleSuffix.para"/></xsl:attribute>
	<xsl:variable name="p.pStyle" select="($ns.styles[@w:styleId=$pStyleId])[1]"/>
	<xsl:variable name="b.bidi">
		<xsl:for-each select="w:pPr[1]/w:bidi[1]"><xsl:choose><xsl:when test="@w:val = 'off'"><xsl:value-of select="$off"/></xsl:when><xsl:otherwise><xsl:value-of select="$on"/></xsl:otherwise></xsl:choose></xsl:for-each>
	</xsl:variable>
	<!-- update the encoded r properties -->
	<xsl:variable name="prs.r.updated">
		<xsl:call-template name="UpdateRPr">
			<xsl:with-param name="p.style" select="$p.pStyle"/>
			<xsl:with-param name="prs.r" select="$prs.r"/>
		</xsl:call-template>
	</xsl:variable>
	<!-- update the encoded p properties -->
	<xsl:variable name="prs.p.updated">
		<xsl:variable name="prs.p.updated1">
			<xsl:call-template name="UpdatePPr">
				<xsl:with-param name="p.style" select="$p.pStyle"/>
				<xsl:with-param name="prs.p" select="$prs.p"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:call-template name="UpdatePPr">
			<xsl:with-param name="prs.p" select="$prs.p.updated1"/>
		</xsl:call-template>
	</xsl:variable>
	<!-- get and apply the CSS properties -->
	<xsl:variable name="styleMod">
		<!-- properties defined in table style that don't inherit -->
		<xsl:value-of select="$prs.pMany"/>
		<!-- properties defined in paragraph style that don't inherit -->
		<xsl:for-each select="$p.pStyle"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
		<!-- direct properties that don't inherit -->
		<xsl:call-template name="ApplyPPr.many">
			<xsl:with-param name="t.cSpacing">
				<xsl:variable name="cspacing" select="$p.pStyle/w:pPr[1]/w:contextualSpacing[1]"/>
				<xsl:if test="$cspacing and not($cspacing/@w:val = 'off')">
					<xsl:if test="following-sibling::*[1]/w:pPr[1]/w:pStyle[1]/@w:val = $pStyleId"><xsl:value-of select="$t.cSpacing.top"/></xsl:if>
					<xsl:if test="preceding-sibling::*[1]/w:pPr[1]/w:pStyle[1]/@w:val = $pStyleId"><xsl:value-of select="$t.cSpacing.bottom"/></xsl:if>
				</xsl:if>
			</xsl:with-param>
		</xsl:call-template>
		<!-- normal properties defined inline -->
		<xsl:call-template name="ApplyPPr.class"/>
		<!-- special properties which context matters -->
		<xsl:call-template name="ApplyPPr.once">
			<xsl:with-param name="b.bidi" select="$b.bidi"/>
			<xsl:with-param name="prs.p" select="$prs.p.updated"/>
			<xsl:with-param name="i.bdrRange.this" select="position()"/>
			<xsl:with-param name="i.bdrRange.last" select="last()"/>
			<xsl:with-param name="pr.bdrBetween" select="$bdrBetween"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:if test="not($styleMod='')"><xsl:attribute name="style"><xsl:value-of select="$styleMod"/></xsl:attribute></xsl:if>
	<!-- wrap inherit class and display content -->
	<span>
	<xsl:attribute name="class"><xsl:value-of select="$pStyleId"/><xsl:value-of select="$styleSuffix.char"/></xsl:attribute>
	<xsl:call-template name="DisplayPContent"><xsl:with-param name="b.bidi" select="$b.bidi"/><xsl:with-param name="prs.r" select="$prs.r.updated"/></xsl:call-template>
	</span>
	</p>
	</xsl:if>
</xsl:template>

<!-- body content includes tables, paragraphs, and others -->
<xsl:template name="DisplayBodyContent">
	<!-- ns.content of "descendant::*[(parent::WX:sect or parent::WX:sub-section) and not(name()='WX:sub-section')]" is used for borders to cross over sub-sections -->
	<xsl:param name="ns.content" select="descendant::*[(parent::WX:sect or parent::WX:sub-section) and not(name()='WX:sub-section')]"/>
	<xsl:param name="prs.pMany" select="''"/>
	<xsl:param name="prs.p" select="$prs.p.default"/>
	<xsl:param name="prs.r" select="$prs.r.default"/>
	<xsl:apply-templates>
		<xsl:with-param name="ns.content" select="$ns.content"/>
		<xsl:with-param name="prs.pMany" select="$prs.pMany"/>	
		<xsl:with-param name="prs.p" select="$prs.p"/>	
		<xsl:with-param name="prs.r" select="$prs.r"/>	
	</xsl:apply-templates>
	<!-- print &nbsp; if there is nothing within the body -->
	<xsl:if test="count($ns.content)=0"><xsl:text disable-output-escaping="yes">&#160;</xsl:text></xsl:if>
</xsl:template>

<!-- apply table cell properties as CSS Class -->
<xsl:template name="ApplyTcPr.class">
	<xsl:apply-templates select="w:tcPr[1]/*" mode="tcpr"/>
</xsl:template>

<!-- tcpr background-color -->
<xsl:template match="w:shd" mode="tcpr"><xsl:call-template name="ApplyShd"/></xsl:template>
<!-- tcpr layout-flow -->
<xsl:template match="w:textFlow" mode="tcpr"><xsl:call-template name="ApplyTextDirection"/></xsl:template>
<!-- tcpr text-fit -->
<xsl:template match="w:tcFitText" mode="tcpr">
	<xsl:if test="not(@w:val = 'off')">text-fit:100%;</xsl:if>
</xsl:template>
<!-- tcpr vertical-align -->
<xsl:template match="w:vAlign" mode="tcpr">
	<xsl:choose>
		<xsl:when test="@w:val = 'center'">vertical-align:middle;</xsl:when>
		<xsl:when test="@w:val = 'bottom'">vertical-align:bottom;</xsl:when>
	</xsl:choose>
</xsl:template>
<!-- tcpr white-space -->
<xsl:template match="w:noWrap" mode="tcpr">
	<xsl:choose>
		<xsl:when test="@w:val = 'off'">white-space:normal;</xsl:when>
		<xsl:otherwise>white-space:nowrap;</xsl:otherwise>
	</xsl:choose>
</xsl:template>
<!-- tcpr width -->
<xsl:template match="w:tcW" mode="tcpr">width:<xsl:call-template name="EvalTableWidth"/>;</xsl:template>
<xsl:template match="*" mode="tcpr"/>

<!-- apply table cell properties as CSS -->
<xsl:template name="ApplyTcPr.once">
	<xsl:param name="cellspacing"/>
	<xsl:param name="cellpadding.default"/>
	<xsl:param name="cellpadding.custom"/>
	<xsl:param name="bdr.top"/>
	<xsl:param name="bdr.left"/>
	<xsl:param name="bdr.bottom"/>
	<xsl:param name="bdr.right"/>
	<xsl:param name="bdr.insideV"/>
	<xsl:param name="thisRow"/>
	<xsl:param name="lastRow"/>
	<xsl:param name="p.tStyle"/>
	<xsl:param name="cnfRow"/>
	<xsl:param name="cnfCol"/>
	<xsl:param name="b.bidivisual"/>
	<!-- get the cnf border properties -->
	<xsl:variable name="cnfType">
		<xsl:if test="not($cnfRow='' and $cnfCol='')">
			<xsl:call-template name="GetCnfType"><xsl:with-param name="cnfRow" select="$cnfRow"/><xsl:with-param name="cnfCol" select="$cnfCol"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:variable>
	<!-- borders (update with cnf) -->
	<xsl:variable name="tcborders" select="w:tcPr[1]/w:tcBorders[1]"/>
	<xsl:variable name="thisBdr.top">
		<xsl:choose>
			<xsl:when test="$tcborders/w:top"><xsl:for-each select="$tcborders/w:top[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
			<xsl:when test="not($cnfType='')">
				<xsl:choose>
					<xsl:when test="$cnfType=$cnfType.band1Vert or $cnfType=$cnfType.band2Vert or $cnfType=$cnfType.firstCol or $cnfType=$cnfType.lastCol">
						<xsl:variable name="p.cnfFirstRow" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.firstRow][1]"/>
						<xsl:choose>
							<xsl:when test="$p.cnfFirstRow and $thisRow=2"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:top[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:when test="not($p.cnfFirstRow) and $thisRow=1"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:top[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:otherwise><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:insideH[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:otherwise>
						</xsl:choose> 
					</xsl:when>
					<xsl:otherwise>
						<xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:top[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise><xsl:value-of select="$bdr.top"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="thisBdr.bottom">
		<xsl:choose>
			<xsl:when test="$tcborders/w:bottom"><xsl:for-each select="$tcborders/w:bottom[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
			<xsl:when test="not($cnfType='')">
				<xsl:choose>
					<xsl:when test="$cnfType=$cnfType.band1Vert or $cnfType=$cnfType.band2Vert or $cnfType=$cnfType.firstCol or $cnfType=$cnfType.lastCol">
						<xsl:variable name="p.cnfLastRow" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.lastRow][1]"/>
						<xsl:choose>
							<xsl:when test="$p.cnfLastRow and $thisRow=$lastRow - 1"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:bottom[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:when test="not($p.cnfLastRow) and $thisRow=$lastRow"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:bottom[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:otherwise><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:insideH[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:otherwise>
						</xsl:choose> 
					</xsl:when>
					<xsl:otherwise>
						<xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:bottom[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise><xsl:value-of select="$bdr.bottom"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="thisBdr.left">
		<xsl:choose>
			<xsl:when test="$tcborders/w:left"><xsl:for-each select="$tcborders/w:left[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
			<xsl:when test="not($cnfType='')">
				<xsl:choose>
					<xsl:when test="$cnfType=$cnfType.band1Horz or $cnfType=$cnfType.band2Horz">
						<xsl:variable name="p.cnfFirstCol" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.firstCol][1]"/>
						<xsl:choose>
							<xsl:when test="$p.cnfFirstCol and position()=2"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:left[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:when test="not($p.cnfFirstCol) and position()=1"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:left[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:otherwise><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:insideV[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:otherwise>
						</xsl:choose> 
					</xsl:when>
					<xsl:when test="$cnfType=$cnfType.firstRow or $cnfType=$cnfType.lastRow">
						<xsl:choose>
							<xsl:when test="position()=1"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:left[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:otherwise><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:insideV[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:otherwise>
						</xsl:choose> 
					</xsl:when>
					<xsl:otherwise>
						<xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:left[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise><xsl:value-of select="$bdr.left"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="thisBdr.right">
		<xsl:choose>
			<xsl:when test="$tcborders/w:right"><xsl:for-each select="$tcborders/w:right[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
			<xsl:when test="not($cnfType='')">
				<xsl:choose>
					<xsl:when test="$cnfType=$cnfType.band1Horz or $cnfType=$cnfType.band2Horz">
						<xsl:variable name="p.cnfLastCol" select="$p.tStyle/w:tStylePr[@w:type=$cnfType.lastCol][1]"/>
						<xsl:choose>
							<xsl:when test="$p.cnfLastCol and position()=last() - 1"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:right[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:when test="not($p.cnfLastCol) and position()=last()"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:right[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:otherwise><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:insideV[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="$cnfType=$cnfType.firstRow or $cnfType=$cnfType.lastRow">
						<xsl:choose>
							<xsl:when test="position()=last()"><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:right[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:when>
							<xsl:otherwise><xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:insideV[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each></xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:for-each select="$p.tStyle/w:tStylePr[@w:type=$cnfType][1]/w:tcPr[1]/w:tcBorders[1]/w:right[1]"><xsl:call-template name="GetBorderPr"/></xsl:for-each>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise><xsl:value-of select="$bdr.right"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="bdrSide.left.bidi">
		<xsl:choose>
			<xsl:when test="$b.bidivisual = $on"><xsl:value-of select="$bdrSide.right"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$bdrSide.left"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="bdrSide.right.bidi">
		<xsl:choose>
			<xsl:when test="$b.bidivisual = $on"><xsl:value-of select="$bdrSide.left"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$bdrSide.right"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$thisBdr.top"/><xsl:with-param name="bdrSide" select="$bdrSide.top"/></xsl:call-template>	
	<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$thisBdr.right"/><xsl:with-param name="bdrSide" select="$bdrSide.right.bidi"/></xsl:call-template>	
	<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$thisBdr.bottom"/><xsl:with-param name="bdrSide" select="$bdrSide.bottom"/></xsl:call-template>	
	<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$thisBdr.left"/><xsl:with-param name="bdrSide" select="$bdrSide.left.bidi"/></xsl:call-template>	
	<!-- padding -->
	<xsl:variable name="cellpadding.custom.merged">
		<!-- directly applied cellpadding -->
		<xsl:variable name="temp.direct">
			<xsl:for-each select="w:tcPr[1]/w:tcMar[1]"><xsl:call-template name="ApplyCellMar"/></xsl:for-each>
		</xsl:variable>
		<xsl:value-of select="$temp.direct"/>
		<xsl:if test="$temp.direct=''">
			<!-- cellpadding from cnf -->
			<xsl:variable name="temp.cnf">
				<xsl:for-each select="$p.tStyle">
					<xsl:call-template name="GetCnfPr.cell">
						<xsl:with-param name="type" select="$t.customCellpadding"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
					</xsl:call-template>
				</xsl:for-each>
			</xsl:variable>
			<xsl:value-of select="$temp.cnf"/>
			<xsl:if test="$temp.cnf=''">
				<!-- cellpadding from table style -->
				<xsl:value-of select="$cellpadding.custom"/>
			</xsl:if>
		</xsl:if>
	</xsl:variable>
	<xsl:variable name="cellpadding.default.merged">
		<!-- default cellpadding from cnf -->
		<xsl:variable name="temp.cnf">
			<xsl:for-each select="$p.tStyle">
				<xsl:call-template name="GetCnfPr.cell">
					<xsl:with-param name="type" select="$t.defaultCellpadding"/><xsl:with-param name="cnfCol" select="$cnfCol"/><xsl:with-param name="cnfRow" select="$cnfRow"/>
				</xsl:call-template>
			</xsl:for-each>
		</xsl:variable>
		<xsl:value-of select="$temp.cnf"/>
		<xsl:if test="$temp.cnf=''">
			<!-- default cellpadding from table style -->
			<xsl:value-of select="$cellpadding.default"/>
		</xsl:if>
	</xsl:variable>
	<xsl:choose>
		<xsl:when test="$cellpadding.custom.merged = 'none' and not($cellpadding.default.merged='')"><xsl:value-of select="$cellpadding.default.merged"/></xsl:when>
		<xsl:when test="not($cellpadding.custom.merged='')"><xsl:value-of select="$cellpadding.custom.merged"/></xsl:when>
		<xsl:when test="not($cellpadding.default.merged='')"><xsl:value-of select="$cellpadding.default.merged"/></xsl:when>	
	</xsl:choose>
</xsl:template>

<!-- table cell -->
<xsl:template match="w:tc">
	<xsl:param name="p.tStyle"/>
	<xsl:param name="prs.pMany"/>
	<xsl:param name="prs.p"/>
	<xsl:param name="prs.r"/>
	<xsl:param name="cellspacing"/>
	<xsl:param name="cellpadding.default"/>
	<xsl:param name="cellpadding.custom"/>
	<xsl:param name="bdr.top"/>
	<xsl:param name="bdr.left"/>
	<xsl:param name="bdr.bottom"/>
	<xsl:param name="bdr.right"/>
	<xsl:param name="bdr.insideV"/>
	<xsl:param name="bdr.insideH"/>
	<xsl:param name="thisRow"/>
	<xsl:param name="lastRow"/>
	<xsl:param name="cnfRow"/>
	<xsl:param name="b.bidivisual"/>
	<xsl:variable name="cnfCol" select="string(w:tcPr[1]/WX:cnfStyle[1]/@WX:val)"/>
	<xsl:variable name="vmerge" select="w:tcPr[1]/w:vmerge[1]"/>
	<xsl:if test="not($vmerge and not($vmerge/@WX:rowspan))">	
		<td>
		<!-- apply class selector -->
		<xsl:attribute name="class">
			<xsl:value-of select="$p.tStyle/@w:styleId"/><xsl:value-of select="$styleSuffix.cell"/>
		</xsl:attribute>
		<!-- apply colspan attribute -->
		<xsl:for-each select="w:tcPr[1]/w:gridSpan[1]/@w:val">
			<xsl:attribute name="colspan">
				<xsl:value-of select="."/>
			</xsl:attribute>
		</xsl:for-each>
		<!-- apply rowspan attribute -->
		<xsl:variable name="rowspan">
			<xsl:choose>
				<xsl:when test="$vmerge/@WX:rowspan"><xsl:value-of select="$vmerge/@WX:rowspan"/></xsl:when>
				<xsl:otherwise>1</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:if test="not($rowspan=1)">
			<xsl:attribute name="rowspan">
				<xsl:value-of select="$rowspan"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:variable name="lastRow.updated" select="$lastRow - $rowspan + 1"/>
		<!-- update borders for the cell
			choose between internal borders or external borders -->
		<xsl:variable name="bdr.bottom.updated">
			<xsl:choose>
				<xsl:when test="$cellspacing='' and $thisRow=$lastRow.updated"><xsl:value-of select="$bdr.bottom"/></xsl:when>
				<xsl:otherwise><xsl:value-of select="$bdr.insideH"/></xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="bdr.left.updated">
			<xsl:choose>
				<xsl:when test="$cellspacing='' and position()=1"><xsl:value-of select="$bdr.left"/></xsl:when>
				<xsl:otherwise><xsl:value-of select="$bdr.insideV"/></xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="bdr.right.updated">
			<xsl:choose>
				<xsl:when test="$cellspacing='' and position()=last()"><xsl:value-of select="$bdr.right"/></xsl:when>
				<xsl:otherwise><xsl:value-of select="$bdr.insideV"/></xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<!-- apply td style -->
		<xsl:attribute name="style">
			<!-- call ApplyTcPr.class for each applicable conditional formatting -->
			<xsl:if test="not($cnfRow='' and $cnfCol='')">
				<xsl:for-each select="$p.tStyle">
					<xsl:call-template name="GetCnfPr.all">
						<xsl:with-param name="type" select="$t.applyTcPr"/>
						<xsl:with-param name="cnfRow" select="$cnfRow"/><xsl:with-param name="cnfCol" select="$cnfCol"/>
					</xsl:call-template>
				</xsl:for-each>
			</xsl:if>
			<!-- for directly applied tc properties -->
			<xsl:call-template name="ApplyTcPr.class"/>
			<xsl:call-template name="ApplyTcPr.once">
				<xsl:with-param name="thisRow" select="$thisRow"/><xsl:with-param name="lastRow" select="$lastRow.updated"/>
				<xsl:with-param name="cellspacing" select="$cellspacing"/><xsl:with-param name="cellpadding.default" select="$cellpadding.default"/><xsl:with-param name="cellpadding.custom" select="$cellpadding.custom"/>
				<xsl:with-param name="bdr.top" select="$bdr.top"/><xsl:with-param name="bdr.left" select="$bdr.left.updated"/><xsl:with-param name="bdr.right" select="$bdr.right.updated"/><xsl:with-param name="bdr.bottom" select="$bdr.bottom.updated"/>
				<xsl:with-param name="bdr.insideV" select="$bdr.insideV"/>
				<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfRow" select="$cnfRow"/><xsl:with-param name="cnfCol" select="$cnfCol"/>
				<xsl:with-param name="b.bidivisual" select="$b.bidivisual"/>
			</xsl:call-template>
		</xsl:attribute>
		<xsl:choose>
			<xsl:when test="$cnfRow='' and $cnfCol=''">
				<!-- display content within as body content -->
				<xsl:call-template name="DisplayBodyContent"><xsl:with-param name="ns.content" select="*"/><xsl:with-param name="prs.pMany" select="$prs.pMany"/><xsl:with-param name="prs.p" select="$prs.p"/><xsl:with-param name="prs.r" select="$prs.r"/></xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<!-- wrap conditional formatting around and display content -->
				<xsl:call-template name="WrapCnf">
					<xsl:with-param name="p.tStyle" select="$p.tStyle"/><xsl:with-param name="cnfRow" select="$cnfRow"/><xsl:with-param name="cnfCol" select="$cnfCol"/>
					<xsl:with-param name="prs.pMany" select="$prs.pMany"/><xsl:with-param name="prs.p" select="$prs.p"/><xsl:with-param name="prs.r" select="$prs.r"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>		
		</td>
	</xsl:if>
</xsl:template>

<!-- apply table row properties as CSS -->
<xsl:template name="ApplyTrPr.class">
	<xsl:for-each select="w:trPr">
		<!-- height -->
		<xsl:text>height:</xsl:text>
		<xsl:choose><xsl:when test="w:trHeight/@w:val"><xsl:value-of select="w:trHeight[1]/@w:val div 20"/>pt</xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose>
		<xsl:text>;</xsl:text>
		<!-- page-break-inside -->
		<xsl:for-each select="w:cantSplit[1]">
			<xsl:choose>
				<xsl:when test="@w:val = 'off'">page-break-inside:auto;</xsl:when>
				<xsl:otherwise>page-break-inside:avoid;</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
	</xsl:for-each>
</xsl:template>

<!-- display i empty table cells, used by w:tr -->
<xsl:template name="DisplayEmptyCell">
	<xsl:param name="i" select="1"/>
	<td colspan="$i"></td>
</xsl:template>

<!-- table row -->
<xsl:template match="w:tr">
	<xsl:param name="p.tStyle"/>
	<xsl:param name="prs.pMany"/>
	<xsl:param name="prs.p"/>
	<xsl:param name="prs.r"/>
	<xsl:param name="cellspacing"/>
	<xsl:param name="cellpadding.default"/>
	<xsl:param name="cellpadding.custom"/>
	<xsl:param name="bdr.top"/>
	<xsl:param name="bdr.left"/>
	<xsl:param name="bdr.bottom"/>
	<xsl:param name="bdr.right"/>
	<xsl:param name="bdr.insideH"/>
	<xsl:param name="bdr.insideV"/>
	<xsl:param name="b.bidivisual"/>
	<tr>
	<!-- class attribute -->
	<xsl:attribute name="class">
		<xsl:value-of select="$p.tStyle/@w:styleId"/><xsl:value-of select="$styleSuffix.row"/>
	</xsl:attribute>
	<!-- get the conditional formatting hints -->
	<xsl:variable name="cnfRow" select="string(w:trPr[1]/WX:cnfStyle[1]/@WX:val)"/>
	<!-- style attribute -->
	<xsl:variable name="styleMod">
		<!-- get tr pagebreak inside propertie from conditional formatting -->
		<xsl:if test="not($cnfRow='')">
			<xsl:for-each select="$p.tStyle">
				<xsl:call-template name="GetCnfPr.row"><xsl:with-param name="type" select="$t.trCantSplit"/><xsl:with-param name="cnfRow" select="$cnfRow"/></xsl:call-template>
			</xsl:for-each>
		</xsl:if>
		<!-- for directly applied tr properties -->
		<xsl:call-template name="ApplyTrPr.class"/>
	</xsl:variable>
	<xsl:if test="not($styleMod='')">
		<xsl:attribute name="style"><xsl:value-of select="$styleMod"/></xsl:attribute>
	</xsl:if>
	<!-- record current row number -->
	<xsl:variable name="thisRow" select="position()"/>
	<xsl:variable name="lastRow" select="last()"/>
	<!-- update borders for the row -->
	<xsl:variable name="bdr.top.updated">
		<xsl:choose>
			<xsl:when test="$cellspacing='' and $thisRow=1"><xsl:value-of select="$bdr.top"/></xsl:when>
			<xsl:otherwise><xsl:value-of select="$bdr.insideH"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<!-- empty grid before -->
	<xsl:for-each select="w:trPr[1]/w:gridBefore[1]/@w:val">
		<xsl:call-template name="DisplayEmptyCell"><xsl:with-param name="i"><xsl:value-of select="."/></xsl:with-param></xsl:call-template>
	</xsl:for-each>
	<!-- apply table cell templates -->
	<xsl:apply-templates select="*[not(name()='w:trPr')]">
		<xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
		<xsl:with-param name="prs.p" select="$prs.p"/>
		<xsl:with-param name="prs.r" select="$prs.r"/>
		<xsl:with-param name="thisRow" select="$thisRow"/><xsl:with-param name="lastRow" select="$lastRow"/>
		<xsl:with-param name="cellspacing" select="$cellspacing"/><xsl:with-param name="cellpadding.default" select="$cellpadding.default"/><xsl:with-param name="cellpadding.custom" select="$cellpadding.custom"/>
		<xsl:with-param name="bdr.top" select="$bdr.top.updated"/><xsl:with-param name="bdr.left" select="$bdr.left"/><xsl:with-param name="bdr.right" select="$bdr.right"/><xsl:with-param name="bdr.bottom" select="$bdr.bottom"/><xsl:with-param name="bdr.insideV" select="$bdr.insideV"/><xsl:with-param name="bdr.insideH" select="$bdr.insideH"/>
		<xsl:with-param name="cnfRow" select="$cnfRow"/>
		<xsl:with-param name="b.bidivisual" select="$b.bidivisual"/>
	</xsl:apply-templates>
	<!-- empty grid after -->
	<xsl:for-each select="w:trPr[1]/w:gridAfter[1]/@w:val">
		<xsl:call-template name="DisplayEmptyCell"><xsl:with-param name="i"><xsl:value-of select="."/></xsl:with-param></xsl:call-template>
	</xsl:for-each>
	</tr>
</xsl:template>

<!-- apply table properties as CSS -->
<xsl:template name="ApplyTblPr.class">
	<xsl:for-each select="w:tblPr[1]">
		<!-- margin -->
		<xsl:variable name="tblppr" select="w:tblpPr[1]"/>
		<xsl:if test="$tblppr/@w:LeftFromText or $tblppr/@w:RightFromText or $tblppr/@w:TopFromText or $tblppr/@w:BottomFromText">
			<xsl:text>margin:</xsl:text>
			<xsl:choose><xsl:when test="$tblppr/@w:TopFromText"><xsl:value-of select="$tblppr/@w:TopFromText[1] div 20"/>pt</xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
			<xsl:choose><xsl:when test="$tblppr/@w:RightFromText"><xsl:value-of select="$tblppr/@w:RightFromText[1] div 20"/>pt</xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
			<xsl:choose><xsl:when test="$tblppr/@w:BottomFromText"><xsl:value-of select="$tblppr/@w:BottomFromText[1] div 20"/>pt</xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose><xsl:text> </xsl:text>
			<xsl:choose><xsl:when test="$tblppr/@w:LeftFromText"><xsl:value-of select="$tblppr/@w:LeftFromText[1] div 20"/>pt</xsl:when><xsl:otherwise>0</xsl:otherwise></xsl:choose><xsl:text>;</xsl:text>
		</xsl:if>
		<!-- width -->
		<xsl:for-each select="w:tblW[1]">width:<xsl:call-template name="EvalTableWidth"/>;</xsl:for-each>
	</xsl:for-each>
</xsl:template>

<!-- table -->
<xsl:template name="tblCore">
	<table>
	<!-- get and apply the class name -->
	<xsl:variable name="tStyleId">
		<xsl:choose>
			<xsl:when test="w:tblPr[1]/w:tblStyle[1]/@w:val">
				<xsl:value-of select="w:tblPr[1]/w:tblStyle[1]/@w:val"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$tStyleId.default"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:attribute name="class"><xsl:value-of select="$tStyleId"/><xsl:value-of select="$styleSuffix.table"/></xsl:attribute>
	<xsl:variable name="p.tStyle" select="($ns.styles[@w:styleId=$tStyleId])[1]"/>
	<!-- get cellspacing and cellpadding for setting default and passing onto td -->
	<xsl:variable name="cellspacingTEMP">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.cellspacing"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:variable name="cellspacing">
		<xsl:choose>
			<!-- KL: this needs to be changed in the future, cellspacing 0 is not the same as no cellspacing -->
			<xsl:when test="$cellspacingTEMP='0'"></xsl:when>
			<xsl:otherwise><xsl:value-of select="$cellspacingTEMP"/></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="cellpadding.default">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.defaultCellpadding"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:variable name="cellpadding.custom">
		<xsl:for-each select="$p.tStyle/w:tcPr[1]/w:tcMar[1]">
			<xsl:call-template name="ApplyCellMar"/>
		</xsl:for-each>	
	</xsl:variable>
	<!-- get table indentation (right or left depending on bidivisual) -->
	<xsl:variable name="tblInd">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.tblInd"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<!-- get outside borders properties for collaspe -->
	<xsl:variable name="bdr.top">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.bdrPr.top"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:variable name="bdr.left">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.bdrPr.left"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:variable name="bdr.bottom">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.bdrPr.bottom"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:variable name="bdr.right">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.bdrPr.right"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<!-- get inside borders properties and passing onto td -->
	<xsl:variable name="bdr.insideH">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.bdrPr.insideH"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:variable name="bdr.insideV">
		<xsl:call-template name="GetTblPr">
			<xsl:with-param name="type" select="$t.bdrPr.insideV"/><xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<!-- bidivisual direction -->
	<xsl:variable name="b.bidivisual">
		<xsl:for-each select="w:tblPr[1]/w:bidiVisual[1]">
			<xsl:choose>
				<xsl:when test="@w:val = 'off'"><xsl:value-of select="$off"/></xsl:when>
				<xsl:otherwise><xsl:value-of select="$on"/></xsl:otherwise>
			</xsl:choose>	
		</xsl:for-each>
	</xsl:variable>
	<!-- align attribute -->
	<xsl:variable name="align"><xsl:for-each select="w:tblPr[1]/w:tblpPr[1]/@w:tblpXSpec"><xsl:value-of select="."/></xsl:for-each></xsl:variable>
	<xsl:if test="not($align='')"><xsl:attribute name="align"><xsl:choose><xsl:when test="$align = 'right' or $align = 'outside'">right</xsl:when><xsl:otherwise>left</xsl:otherwise></xsl:choose></xsl:attribute></xsl:if>
	<!-- cellspacing attribute -->
	<xsl:attribute name="cellspacing">
		<xsl:choose>
			<xsl:when test="$cellspacing=''">0</xsl:when>
			<xsl:otherwise><xsl:value-of select="($cellspacing div 1440) * $pixelsPerInch"/></xsl:otherwise>
		</xsl:choose>
	</xsl:attribute>
	<xsl:if test="$cellspacing=''"><xsl:attribute name="cellspacing">0</xsl:attribute></xsl:if>
	<!-- apply the CSS properties -->
	<xsl:variable name="styleMod">
		<xsl:call-template name="ApplyTblPr.class"/>
		<!-- border-collapse -->
		<xsl:choose>
			<xsl:when test="$cellspacing=''">border-collapse:collapse;</xsl:when>
			<xsl:otherwise>
				<xsl:text>border-collapse:separate;</xsl:text>
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$bdr.top"/><xsl:with-param name="bdrSide" select="$bdrSide.top"/></xsl:call-template>
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$bdr.left"/><xsl:with-param name="bdrSide" select="$bdrSide.left"/></xsl:call-template>
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$bdr.bottom"/><xsl:with-param name="bdrSide" select="$bdrSide.bottom"/></xsl:call-template>
				<xsl:call-template name="ApplyBorderPr"><xsl:with-param name="pr.bdr" select="$bdr.right"/><xsl:with-param name="bdrSide" select="$bdrSide.right"/></xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
		<!-- direction:rtl; -->
		<xsl:if test="$b.bidivisual=$on">direction:rtl;</xsl:if>
		<!-- margin-left/margin-right -->
		<xsl:if test="not($tblInd='')">
			<xsl:text>margin-</xsl:text>
			<xsl:choose>
				<xsl:when test="$b.bidivisual=$on">right</xsl:when>
				<xsl:otherwise>left</xsl:otherwise>
			</xsl:choose>
			<xsl:text>:</xsl:text>
			<xsl:value-of select="$tblInd"/>
			<xsl:text>;</xsl:text>
		</xsl:if>
	</xsl:variable>
	<xsl:if test="not($styleMod='')">
		<xsl:attribute name="style"><xsl:value-of select="$styleMod"/></xsl:attribute>
	</xsl:if>
	<!-- table properties that will be passed on to cells/paragraphs -->
	<xsl:variable name="prs.pMany">
		<!-- direction of table should not be inherited by paragraph -->
		<xsl:for-each select="w:tblPr[1]/w:bidiVisual[1]"><xsl:if test="not(@w:val = 'off')">direction:ltr;</xsl:if></xsl:for-each>
		<!-- non-inherit paragraph properties -->
		<xsl:for-each select="$p.tStyle"><xsl:call-template name="ApplyPPr.many"/></xsl:for-each>
	</xsl:variable>
	<!-- table properties that will be passed on to runs (via a special record) -->
	<xsl:variable name="prs.r">
		<xsl:call-template name="UpdateRPr">
			<xsl:with-param name="p.style" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<!-- table properties that will be passed on to paargraphs (via a special record) -->
	<xsl:variable name="prs.p">
		<xsl:call-template name="UpdatePPr">
			<xsl:with-param name="p.style" select="$p.tStyle"/>
		</xsl:call-template>
	</xsl:variable>
	<!-- display the rows -->
	<xsl:apply-templates select="*[not(name()='w:tblPr' or name()='w:tblGrid')]">
		<xsl:with-param name="p.tStyle" select="$p.tStyle"/>
		<xsl:with-param name="prs.pMany" select="$prs.pMany"/>
		<xsl:with-param name="prs.p" select="$prs.p"/>
		<xsl:with-param name="prs.r" select="$prs.r"/>
		<xsl:with-param name="cellspacing" select="$cellspacing"/><xsl:with-param name="cellpadding.default" select="$cellpadding.default"/><xsl:with-param name="cellpadding.custom" select="$cellpadding.custom"/>
		<xsl:with-param name="bdr.top" select="$bdr.top"/><xsl:with-param name="bdr.left" select="$bdr.left"/><xsl:with-param name="bdr.right" select="$bdr.right"/><xsl:with-param name="bdr.bottom" select="$bdr.bottom"/>
		<xsl:with-param name="bdr.insideH" select="$bdr.insideH"/><xsl:with-param name="bdr.insideV" select="$bdr.insideV"/>
		<xsl:with-param name="b.bidivisual" select="$b.bidivisual"/>
	</xsl:apply-templates>
	<!-- display table grid -->
	<xsl:for-each select="w:tblGrid[1]">
		<xsl:text disable-output-escaping="yes">&lt;![if !supportMisalignedColumns]&gt;</xsl:text>
		<tr height="0">
		<xsl:for-each select="w:gridCol">
			<xsl:variable name="gridStyle">margin:0;padding:0;border:none;width:<xsl:call-template name="EvalTableWidth"/>;</xsl:variable>
			<td style="{$gridStyle}"/>
		</xsl:for-each>
		</tr>
		<xsl:text disable-output-escaping="yes">&lt;![endif]&gt;</xsl:text>
	</xsl:for-each>
	</table>
</xsl:template>

<xsl:template match="w:tbl[w:tblPr/w:jc/@w:val]">
	<xsl:variable name="p.Jc" select="w:tblPr/w:jc/@w:val"/>
	<div>
		<xsl:attribute name="align"><xsl:value-of select="$p.Jc"/></xsl:attribute>
<!--			<xsl:choose>
				<xsl:when test="$p.Jc='center'"><xsl:text>center</xsl:text></xsl:when>
				<xsl:when test="$p.Jc='left'"><xsl:text>left</xsl:text></xsl:when>
				<xsl:when test="$p.Jc='right'"><xsl:text>right</xsl:text></xsl:when>
			</xsl:choose>-->
		<xsl:call-template name="tblCore"/>
	</div>
</xsl:template>

<xsl:template match="w:tbl">
	<xsl:call-template name="tblCore"/>
</xsl:template>

<xsl:template name="hrCore">
	<xsl:param name="p.Hr"/>
		<hr>
			<xsl:attribute name="style"><xsl:value-of select="substring-after($p.Hr/@style, ';')"/></xsl:attribute>
			<xsl:attribute name="align"><xsl:value-of select="$p.Hr/@o:hralign"/></xsl:attribute>
			<xsl:if test="$p.Hr/@o:hrnoshade='t'">
				<xsl:attribute name="noshade">
					<xsl:text>1</xsl:text>
				</xsl:attribute>
				<xsl:attribute name="color">
					<xsl:value-of select="$p.Hr/@fillcolor"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="$p.Hr/@o:hrpct">
				<xsl:attribute name="width">
					<xsl:value-of select="$p.Hr/@o:hrpct div 10"/>
					<xsl:text>%</xsl:text>
				</xsl:attribute>
			</xsl:if>
		</hr>
</xsl:template>

<xsl:template match="w:p[w:r[1]//v:rect/@o:hrstd and not(w:r[2])]">
	<xsl:call-template name="hrCore">
		<xsl:with-param name="p.Hr" select="w:r//v:rect"/>
	</xsl:call-template>
</xsl:template>

<xsl:template match="v:rect[@o:hrstd]">
	<xsl:call-template name="hrCore">
		<xsl:with-param name="p.Hr" select="."/>
	</xsl:call-template>
</xsl:template>

<!-- sub-section -->
<!-- this will normally not happen, as WX:sect should take care of everything -->
<xsl:template match="WX:sub-section">
	<xsl:call-template name="DisplayBodyContent"/>
</xsl:template>

<!-- section -->
<xsl:template match="WX:sect">
	<xsl:variable name="thisSect" select="."/>
	<div>
	<!-- apply section properties and display body content within the section -->
	<xsl:for-each select="//WX:sect"><xsl:if test=".=$thisSect"><xsl:attribute name="class">Section<xsl:value-of select="position()"/></xsl:attribute></xsl:if></xsl:for-each>
	<xsl:call-template name="DisplayBodyContent"/>
	</div>
</xsl:template>

<!-- body -->
<xsl:template match="w:body">
	<!-- we might need to ouput a margin from the "body div". -->
	<xsl:attribute name="style">
		<xsl:variable name="divBody" select="/w:wordDocument/w:divs/w:div[w:bodyDiv/@w:val='on']"/>
		<xsl:variable name="dxaBodyLeft">
			<xsl:value-of select="$divBody/w:marLeft/@w:val"/>
		</xsl:variable>
		<xsl:variable name="dxaBodyRight">
			<xsl:value-of select="$divBody/w:marRight/@w:val"/>
		</xsl:variable>
		<xsl:if test="not($dxaBodyLeft='')">
				<xsl:text>margin-left:</xsl:text><xsl:value-of select="$dxaBodyLeft div 20"/><xsl:text>pt;</xsl:text>
		</xsl:if>		
		<xsl:if test="not($dxaBodyRight='')">
			<xsl:text>margin-right:</xsl:text><xsl:value-of select="$dxaBodyRight div 20"/><xsl:text>pt;</xsl:text>
		</xsl:if>		
	</xsl:attribute>
	<xsl:apply-templates select="*"/>
</xsl:template>

<!-- font -->
<xsl:template match="w:font">
	<xsl:text>@font-face{font-family:"</xsl:text>
	<xsl:value-of select="@w:name"/>
	<xsl:text>";panose-1:</xsl:text>
	<xsl:variable name="panose1">
		<xsl:call-template name="ConvHex2Dec">
			<xsl:with-param name="value" select="w:panose-1[1]/@w:val"/>
			<xsl:with-param name="i" select="2"/>
			<xsl:with-param name="s" select="2"/>
		</xsl:call-template>
	</xsl:variable>
	<xsl:value-of select="substring($panose1,2)"/>
	<xsl:text>;}</xsl:text>
</xsl:template>

<!-- generate a character CSS class -->
<xsl:template name="MakeRStyle">
	<xsl:param name="basetype"/>
	<xsl:text>.</xsl:text><xsl:value-of select="@w:styleId"/><xsl:value-of select="$styleSuffix.char"/>
	<xsl:text>{</xsl:text>
	<xsl:choose>
		<xsl:when test="$basetype='paragraph'">
			<xsl:text>font-size: 10pt;</xsl:text>
		</xsl:when>
	</xsl:choose>
	<xsl:call-template name="ApplyRPr.class"/>
	<xsl:text>} </xsl:text>
</xsl:template>

<!-- setup the style as CSS class -->
<xsl:template match="w:style">
	<xsl:choose>
		<!-- for character style -->
		<xsl:when test="@w:type = 'character'">
			<xsl:call-template name="MakeRStyle"/>
		</xsl:when>
		<!-- for paragraph style -->
		<xsl:when test="@w:type = 'paragraph'">
			<!-- define the paragraph style -->
			<xsl:text>.</xsl:text><xsl:value-of select="@w:styleId"/><xsl:value-of select="$styleSuffix.para"/>
			<xsl:text>{margin-left:</xsl:text><xsl:value-of select="$pMargin.default.left"/>
			<xsl:text>;margin-right:</xsl:text><xsl:value-of select="$pMargin.default.right"/>
			<!-- be carful to only put out the default iff it is not an 'auto' value -->
			<xsl:variable name="spacing" select="w:pPr[1]/w:spacing[1]"/>
		
			<xsl:if test="(not($spacing/@w:before-autospacing) or $spacing/@w:before-autospacing = 'off')">
				<xsl:text>;margin-top:</xsl:text><xsl:value-of select="$pMargin.default.top"/>
			</xsl:if>

			<xsl:if test="(not($spacing/@w:after-autospacing) or $spacing/@w:after-autospacing = 'off')">
				<xsl:text>;margin-bottom:</xsl:text><xsl:value-of select="$pMargin.default.bottom"/>
			</xsl:if>

			<xsl:text>;font-size:10.0pt;font-family:"Times New Roman";</xsl:text>
			<xsl:call-template name="ApplyPPr.class"/>
			<xsl:text>} </xsl:text>
			<!-- define a class for character properties within paragraph style, used to wrap around -->
			<xsl:call-template name="MakeRStyle"><xsl:with-param name="basetype" select="'paragraph'"/></xsl:call-template>
		</xsl:when>
		<!-- for table style (table, row, cell) -->
		<xsl:when test="@w:type = 'table'">
			<xsl:variable name="styleId" select="@w:styleId"/>
			<!-- define the table style -->
			<xsl:text>.</xsl:text><xsl:value-of select="$styleId"/><xsl:value-of select="$styleSuffix.table"/>
			<xsl:text>{</xsl:text><xsl:call-template name="ApplyTblPr.class"/><xsl:text>} </xsl:text>
			<!-- define the row style -->
			<xsl:text>.</xsl:text><xsl:value-of select="$styleId"/><xsl:value-of select="$styleSuffix.row"/>
			<xsl:text>{</xsl:text><xsl:call-template name="ApplyTrPr.class"/><xsl:text>} </xsl:text>
			<!-- define the cell style -->
			<xsl:text>.</xsl:text><xsl:value-of select="$styleId"/><xsl:value-of select="$styleSuffix.cell"/>
			<xsl:text>{vertical-align:top;</xsl:text>
			<!-- table cell properties -->
			<xsl:call-template name="ApplyTcPr.class"/>
			<!-- table cell properties include inheritable paragraph properties -->
			<xsl:call-template name="ApplyPPr.class"/>
			<!-- table cell properties include inheritable character properties -->
			<xsl:call-template name="ApplyRPr.class"/>
			<xsl:text>} </xsl:text>
			<!-- define conditional formatting sub styles (characters and paragraphs) -->
			<xsl:for-each select="w:tStylePr">
				<xsl:text>.</xsl:text><xsl:value-of select="$styleId"/>-<xsl:value-of select="@w:type"/>
				<xsl:text>{vertical-align:top;</xsl:text>
				<xsl:call-template name="ApplyPPr.class"/>
				<xsl:call-template name="ApplyRPr.class"/>
				<xsl:text>} </xsl:text>
			</xsl:for-each>
		</xsl:when>
	</xsl:choose>
</xsl:template>

<!-- MAIN ROOT TRANSFORM -->
<xsl:template match="/w:wordDocument">

<html>

<head>
<!-- meta tags -->
<meta name="Generator" content="Microsoft Word 11 XSLT"/>

<!-- link base -->
<xsl:for-each select="$p.docInfo/w:linkBase[1]/@w:val"><base href="{.}"/></xsl:for-each>
<!-- title -->
<title><xsl:value-of select="$p.docInfo/w:title[1]/@w:val"/></title>

<!-- javascript used for displaying annotation comment -->
<xsl:call-template name="DisplayAnnotationScript"/>


<xsl:comment><xsl:text disable-output-escaping="yes">[if !mso]&gt;</xsl:text>
<xsl:text disable-output-escaping="yes">&lt;style&gt;</xsl:text>
/*vml*/
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w10\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
<xsl:text disable-output-escaping="yes">&lt;/style&gt;</xsl:text>
<xsl:text disable-output-escaping="yes">&lt;![endif]</xsl:text></xsl:comment>

<style>
<xsl:comment>

/*font definitions*/
<xsl:apply-templates select="w:fonts[1]/w:font"/>

/*element styles*/

<xsl:choose>
	<xsl:when test="w:docPr/w:revisionView/@w:markup = 'off'">
		del {display:none;}
		ins {text-decoration:none;}
	</xsl:when>
	<xsl:otherwise>
		del {text-decoration:line-through;color:red;}
		ins {text-decoration:underline;color:teal;}
	</xsl:otherwise>
</xsl:choose>

a:link {color:blue;text-decoration:underline;text-underline:single;}
a:visited {color:purple;text-decoration:underline;text-underline:single;}

/*class styles*/
<xsl:apply-templates select="$ns.styles"/>

/*sections*/
<xsl:for-each select="//WX:sect">
	<xsl:variable name="sectName" select="concat('Section',position())"/>
	<xsl:text>@page </xsl:text><xsl:value-of select="$sectName"/><xsl:text>{</xsl:text>
	<xsl:for-each select="w:sectPr[1]">
		<!-- section size and margin -->
		<xsl:variable name="pgsz" select="w:pgSz[1]"/>
		<xsl:choose>
			<xsl:when test="$pgsz">size:<xsl:value-of select="$pgsz/@w:w div 20"/>pt <xsl:value-of select="$pgsz/@w:h div 20"/>pt;</xsl:when>
			<xsl:otherwise>size:8.5in 11in;</xsl:otherwise>
		</xsl:choose>
		<xsl:variable name="pgmar" select="w:pgMar[1]"/>
		<xsl:choose>
			<xsl:when test="$pgmar">margin:<xsl:value-of select="$pgmar/@w:top div 20"/>pt <xsl:value-of select="$pgmar/@w:right div 20"/>pt <xsl:value-of select="$pgmar/@w:bottom div 20"/>pt <xsl:value-of select="$pgmar/@w:left div 20"/>pt;</xsl:when>
			<xsl:otherwise>margin:1in 1.25in 1in 1.25in;</xsl:otherwise>
		</xsl:choose>
		<!-- layout-flow -->
		<xsl:for-each select="w:textFlow[1]"><xsl:call-template name="ApplyTextDirection"/></xsl:for-each>
	</xsl:for-each>
	<xsl:text>} div.</xsl:text><xsl:value-of select="$sectName"/>{page:<xsl:value-of select="$sectName"/>;}
</xsl:for-each>

</xsl:comment>
</style>

</head>

<!-- body -->
<body>
<!-- fetch the background if we can -->
<xsl:if test="w:bgPict/w:background/@w:bgcolor">
	<xsl:attribute name="bgcolor">
		<xsl:value-of select="w:bgPict/w:background/@w:bgcolor"/>
	</xsl:attribute>
</xsl:if>

<xsl:apply-templates select="w:body"/>

<!-- hidden annotation comments -->
<xsl:for-each select="//aml:annotation[@w:type='Word.Comment']">
	<xsl:call-template name="DisplayAnnotationText"/>
</xsl:for-each>
</body>

</html>
</xsl:template>

<!-- annotation-bookmark -->
<xsl:template match="aml:annotation[@w:type='Word.Bookmark.Start']">
<a name="{@w:name}"/>
</xsl:template>

<!-- annotation-insertion -->
<xsl:template match="aml:annotation[@w:type='Word.Insertion']">
	<xsl:param name="b.bidi"/>
	<xsl:param name="prs.r"/>
	<ins>
	<xsl:for-each select="aml:content">
		<xsl:call-template name="DisplayPContent">
			<xsl:with-param name="b.bidi" select="$b.bidi"/>
			<xsl:with-param name="prs.r" select="$prs.r"/>
		</xsl:call-template>
	</xsl:for-each>
	</ins>
</xsl:template>

<!-- annotation-deletion -->
<xsl:template match="aml:annotation[@w:type='Word.Deletion']">
	<xsl:param name="b.bidi"/>
	<xsl:param name="prs.r"/>
	<del>
	<xsl:for-each select="aml:content">
		<xsl:call-template name="DisplayPContent">
			<xsl:with-param name="b.bidi" select="$b.bidi"/>
			<xsl:with-param name="prs.r" select="$prs.r"/>
		</xsl:call-template>
	</xsl:for-each>
	</del>
</xsl:template>

<!-- link for annotation (comment) -->
<xsl:template match="aml:annotation[@w:type='Word.Comment']">
	<xsl:variable name="id" select="@aml:id + 1"/>
	<a class="msocomanchor" id="_anchor_{$id}" onmouseover="msoCommentShow('_anchor_{$id}','_com_{$id}')" onmouseout="msoCommentHide('_com_{$id}')" href="#_msocom_{$id}" language="JavaScript" name="_msoanchor_{$id}">
	<xsl:value-of select="concat('[',@w:initials,$id,']')"/>
	</a>
</xsl:template>

<!-- text for annotation (comment) -->
<xsl:template name="DisplayAnnotationText">
	<xsl:variable name="id" select="@aml:id + 1"/>
	<div id="_com_{$id}" class="msocomtxt" language="JavaScript" onmouseover="msoCommentShow('_anchor_{$id}','_com_{$id}')" onmouseout="msoCommentHide('_com_{$id}')">
	<a name="_msocom_{$id}"></a>
	<a href="#_msoanchor_{$id}" class="msocomoff">
	<xsl:value-of select="concat('[',@w:initials,$id,']')"/>
	</a>
	<xsl:for-each select="aml:content">
		<xsl:call-template name="DisplayBodyContent">
			<xsl:with-param name="ns.content" select="*"/>
		</xsl:call-template>
	</xsl:for-each>
	</div>
</xsl:template>

<!-- javascript for annotation (comment) -->
<xsl:template name="DisplayAnnotationScript">
<xsl:text disable-output-escaping="yes">&lt;![if !supportAnnotations]&gt;</xsl:text>
<style id="dynCom" type="text/css"></style>
<script language="JavaScript">
<xsl:comment>
<xsl:text disable-output-escaping="yes">
function msoCommentShow(anchor_id, com_id)
{
	if(msoBrowserCheck()) 
		{
		c = document.all(com_id);
		a = document.all(anchor_id);
		if (null != c &amp;&amp; null == c.length &amp;&amp; null != a &amp;&amp; null == a.length)
			{
			var cw = c.offsetWidth;
			var ch = c.offsetHeight;
			var aw = a.offsetWidth;
			var ah = a.offsetHeight;
			var x  = a.offsetLeft;
			var y  = a.offsetTop;
			var el = a;
			while (el.tagName != "BODY") 
				{
				el = el.offsetParent;
				x = x + el.offsetLeft;
				y = y + el.offsetTop;
				}
			var bw = document.body.clientWidth;
			var bh = document.body.clientHeight;
			var bsl = document.body.scrollLeft;
			var bst = document.body.scrollTop;
			if (x + cw + ah / 2 > bw + bsl &amp;&amp; x + aw - ah / 2 - cw >= bsl ) 
				{ c.style.left = x + aw - ah / 2 - cw; }
			else 
				{ c.style.left = x + ah / 2; }
			if (y + ch + ah / 2 > bh + bst &amp;&amp; y + ah / 2 - ch >= bst ) 
				{ c.style.top = y + ah / 2 - ch; }
			else 
				{ c.style.top = y + ah / 2; }
			c.style.visibility = "visible";
}	}	}
function msoCommentHide(com_id) 
{
	if(msoBrowserCheck())
		{
		c = document.all(com_id);
		if (null != c &amp;&amp; null == c.length)
		{
		c.style.visibility = "hidden";
		c.style.left = -1000;
		c.style.top = -1000;
		} } 
}
function msoBrowserCheck()
{
	ms = navigator.appVersion.indexOf("MSIE");
	vers = navigator.appVersion.substring(ms + 5, ms + 6);
	ie4 = (ms > 0) &amp;&amp; (parseInt(vers) >= 4);
	return ie4;
}
if (msoBrowserCheck())
{
	document.styleSheets.dynCom.addRule(".msocomanchor","background: infobackground");
	document.styleSheets.dynCom.addRule(".msocomoff","display: none");
	document.styleSheets.dynCom.addRule(".msocomtxt","visibility: hidden");
	document.styleSheets.dynCom.addRule(".msocomtxt","position: absolute");
	document.styleSheets.dynCom.addRule(".msocomtxt","top: -1000");
	document.styleSheets.dynCom.addRule(".msocomtxt","left: -1000");
	document.styleSheets.dynCom.addRule(".msocomtxt","width: 33%");
	document.styleSheets.dynCom.addRule(".msocomtxt","background: infobackground");
	document.styleSheets.dynCom.addRule(".msocomtxt","color: infotext");
	document.styleSheets.dynCom.addRule(".msocomtxt","border-top: 1pt solid threedlightshadow");
	document.styleSheets.dynCom.addRule(".msocomtxt","border-right: 2pt solid threedshadow");
	document.styleSheets.dynCom.addRule(".msocomtxt","border-bottom: 2pt solid threedshadow");
	document.styleSheets.dynCom.addRule(".msocomtxt","border-left: 1pt solid threedlightshadow");
	document.styleSheets.dynCom.addRule(".msocomtxt","padding: 3pt 3pt 3pt 3pt");
	document.styleSheets.dynCom.addRule(".msocomtxt","z-index: 100");
}
</xsl:text>
</xsl:comment>
</script>
<xsl:text disable-output-escaping="yes">&lt;![endif]&gt;</xsl:text>
</xsl:template>

<!-- general template matching -->

<!-- pass everything else not defined here -->
<xsl:template name="copyElements">
	<xsl:element name="{name()}" namespace="{namespace-uri()}">
		<xsl:for-each select="@*">
			<xsl:attribute name="{name()}">
				<xsl:value-of select="."/>
			</xsl:attribute>
		</xsl:for-each>
		<xsl:apply-templates/>	
	</xsl:element>
</xsl:template>
	
<xsl:template match="*">
	<xsl:call-template name="copyElements"/>
</xsl:template>

<!-- ignore everything in the word namespace not defined here -->
<xsl:template match="w:*"/>

<xsl:template match="WX:*"/>

<xsl:template match="v:*">
	<xsl:choose>
		<xsl:when test=".//w10:wrap[@type='topAndBottom']">
			<o:wrapblock>
				<xsl:call-template name="copyElements"/>
			</o:wrapblock>
			<br style="mso-ignore:vglayout" clear='ALL'/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:call-template name="copyElements"/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template match="o:WordFieldCodes"/>

</xsl:stylesheet>
