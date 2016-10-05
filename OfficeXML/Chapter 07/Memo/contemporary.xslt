<?xml  version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:w="http://schemas.microsoft.com/office/word/2003/2/wordml" 
	xmlns:ns0="urn:schemas-microsoft-com.office.demos.memo" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:v="urn:schemas-microsoft-com:vml" 
	xmlns:w10="urn:schemas-microsoft-com:office:word" 
	xmlns:o="urn:schemas-microsoft-com:office:office">
	
    <xsl:template match="/">
        <w:wordDocument>
        
            <!-- ************************            ***            ***	Define the styles to be referenced in the document            ***            ****************************  -->
            
            <w:styles>
                <w:versionOfBuiltInStylenames w:val="3"/>
                
                <w:style w:type="paragraph" w:default="on" w:styleId="Normal">
                    <w:name w:val="Normal"/>
                    <w:pPr>
                        <w:ind w:left="835"/>
                    </w:pPr>
                </w:style>
                                
                <w:style w:type="paragraph" w:styleId="BodyText">
                    <w:name w:val="Body Text"/>
                    <w:basedOn w:val="Normal"/>
                    <w:pPr>
                        <w:pStyle w:val="BodyText"/>
                        <w:spacing w:after="220" w:line="220" w:line-rule="at-least"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="DocumentLabel">
                    <w:name w:val="Document Label"/>
                    <w:next w:val="Normal"/>
                    <w:pPr>
                        <w:pStyle w:val="DocumentLabel"/>
                        <w:spacing w:before="140" w:after="540" w:line="600" w:line-rule="at-least"/>
                        <w:ind w:left="840"/>
                    </w:pPr>
                    <w:rPr>
                        <w:spacing w:val="-38"/>
                        <w:sz w:val="60"/>
                        <w:lang w:val="EN-US" w:fareast="EN-US" w:bidi="AR-SA"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeader">
                    <w:name w:val="Message Header"/>
                    <w:basedOn w:val="BodyText"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeader"/>
                        <w:keepLines/>
                        <w:spacing w:after="0" w:line="415" w:line-rule="at-least"/>
                        <w:ind w:left="1560" w:hanging="720"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeaderFirst">
                    <w:name w:val="Message Header First"/>
                    <w:basedOn w:val="MessageHeader"/>
                    <w:next w:val="MessageHeader"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeaderFirst"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="character" w:styleId="MessageHeaderLabel">
                    <w:name w:val="Message Header Label"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:h-ansi="Arial"/>
                        <w:b/>
                        <w:spacing w:val="-4"/>
                        <w:sz w:val="18"/>
                        <w:vertAlign w:val="baseline"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeaderLast">
                    <w:name w:val="Message Header Last"/>
                    <w:basedOn w:val="MessageHeader"/>
                    <w:next w:val="BodyText"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeaderLast"/>
                        <w:pBdr>
                            <w:bottom w:val="single" w:sz="6" w:space="22" w:color="auto"/>
                        </w:pBdr>
                        <w:spacing w:after="400"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="character" w:styleId="PageNumber">
                    <w:name w:val="page number"/>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="Slogan">
                    <w:name w:val="Slogan"/>
                    <w:basedOn w:val="Normal"/>
                    <w:pPr>
                        <w:pStyle w:val="Slogan"/>
                        <w:framePr w:w="5170" w:h="1800" w:hspace="187" w:vspace="187" w:wrap="not-beside" w:vanchor="page" w:hanchor="page" w:x="966" w:y-align="bottom" w:anchor-lock="on"/>
                        <w:ind w:left="0"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Impact" w:h-ansi="Impact"/>
                        <w:caps/>
                        <w:color w:val="DFDFDF"/>
                        <w:spacing w:val="20"/>
                        <w:sz w:val="48"/>
                    </w:rPr>
                </w:style>
                
            </w:styles>
            <w:docPr>
                <w:view w:val="print"/>
                <w:zoom w:val="best-fit" w:percent="148"/>
                <w:validateAgainstSchema/>
                
                <!-- Don't allow save if the document is not valid according to the schema -->
                <w:saveInvalidXML w:val="off"/>
                
                <!-- Make sure the tag view is off -->
                <w:showXMLTags w:val="off"/>
                
                <!-- Don't include Mixed Content in the validation -->
                <w:ignoreMixedContent/>
            </w:docPr>
            <w:body>
                <xsl:apply-templates select="/ns0:memo"/>
            </w:body>
        </w:wordDocument>
    </xsl:template>
    <xsl:template match="ns0:memo">
        <ns0:memo>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="DocumentLabel"/>
                </w:pPr>
                
                <!-- Here, we will output an image that will be anchored to the bottom of the memo -->
                <w:r>
                    <w:pict>
                        <v:group id="_x0000_s1037" style="position:absolute;margin-left:36.1pt;margin-top:324pt;width:146.4pt;height:146.4pt;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" coordorigin="722,6480" coordsize="2928,2928" o:allowincell="f">
                            <v:oval id="_x0000_s1038" style="position:absolute;left:722;top:6480;width:2928;height:2928" fillcolor="#f2f2f2" stroked="f" strokeweight=".25pt"/>
                            <v:rect id="_x0000_s1039" style="position:absolute;left:1695;top:6480;width:941;height:2928" stroked="f" strokeweight=".25pt"/>
                            <w10:wrap anchorx="page" anchory="page"/>
                            <w10:anchorlock/>
                        </v:group>
                    </w:pict>
                </w:r>
                
                <w:r>
                    <w:t>Memorandum</w:t>
                </w:r>
            </w:p>
            
            <!-- Create the To: Line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeaderFirst"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                        <w:spacing w:val="-20"/>
                    </w:rPr>
                    <w:t>To:</w:t>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve">	</w:t>
                </w:r>
                <ns0:to>
                    <w:r>
                        <w:t>
                            <xsl:value-of select="ns0:to"/>
                        </w:t>
                    </w:r>
                </ns0:to>
            </w:p>
            
            <!-- Create the From: Line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeader"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t>From:</w:t>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve">	</w:t>
                </w:r>
                <ns0:from>
                    <w:r>
                        <w:t>
                            <xsl:value-of select="ns0:from"/>
                        </w:t>
                    </w:r>
                </ns0:from>
            </w:p>
            
            <!-- Create the Priority: Line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeader"/>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                        <w:b w:val="off"/>
                    </w:rPr>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t xml:space="preserve">Priority:	</w:t>
                </w:r>
                <ns0:priority>
                    <w:r>
                        <w:rPr>
                            <w:rStyle w:val="MessageHeaderLabel"/>
                            <w:rFonts w:ascii="Times New Roman" w:h-ansi="Times New Roman"/>
                            <w:b w:val="off"/>
                            <w:sz w:val="20"/>
                        </w:rPr>
                        <w:t>
                            <xsl:value-of select="ns0:priority"/>
                        </w:t>
                    </w:r>
                </ns0:priority>
            </w:p>
            
            <!-- Create the Date: Line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeader"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t>Date:</w:t>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve">	</w:t>
                </w:r>
                <ns0:date>
					<w:r>
						<w:rPr>
							<w:noProof/>
						</w:rPr>
						<w:t><xsl:value-of select="ns0:date"/></w:t>
					</w:r>
				</ns0:date>
            </w:p>
            
            <!-- Create the Subject: line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeaderLast"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t>Re:</w:t>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve">	</w:t>
                </w:r>
                <ns0:subject>
                    <w:r>
                        <w:t>
                            <xsl:value-of select="ns0:subject"/>
                        </w:t>
                    </w:r>
                </ns0:subject>
            </w:p>
            
            <!-- ******************            ****            ****  Here is where the body of the memo will go            ****            ****            ********************* -->
            
            <ns0:body>
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="BodyText"/>
                        <w:ind w:right="25"/>
                    </w:pPr>
                    <w:r>
                        <w:t>
                            <xsl:value-of select="ns0:body"/>></w:t>
                    </w:r>
                </w:p>
            </ns0:body>
            
            <!-- Create a small frame that will go under the body of the document, and contain the text: "Confidential" -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="Slogan"/>
                    <w:framePr w:wrap="not-beside"/>
                    <w:rPr>
                        <w:color w:val="C0C0C0"/>
                    </w:rPr>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:color w:val="C0C0C0"/>
                    </w:rPr>
                    <w:t>Confidential</w:t>
                </w:r>
            </w:p>
            
        </ns0:memo>
    </xsl:template>
</xsl:stylesheet>