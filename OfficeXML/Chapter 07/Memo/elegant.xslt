<?xml  version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:w="http://schemas.microsoft.com/office/word/2003/2/wordml" 
	xmlns:ns0="urn:schemas-microsoft-com.office.demos.memo" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    
    <xsl:template match="/">
        <w:wordDocument>
        
            <!-- ************************            ***            ***	Define the styles to be referenced in the document            ***            ****************************  -->
            
            <w:styles>
                <w:versionOfBuiltInStylenames w:val="3"/>
                
                <w:style w:type="paragraph" w:default="on" w:styleId="Normal">
                    <w:name w:val="Normal"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Garamond" w:h-ansi="Garamond"/>
                        <w:sz w:val="22"/>
                        <w:lang w:val="EN-US" w:fareast="EN-US" w:bidi="AR-SA"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="character" w:default="on" w:styleId="DefaultParagraphFont">
                    <w:name w:val="Default Paragraph Font"/>
                    <w:semiHidden/>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="BodyText">
                    <w:name w:val="Body Text"/>
                    <w:basedOn w:val="Normal"/>
                    <w:pPr>
                        <w:pStyle w:val="BodyText"/>
                        <w:spacing w:after="240" w:line="240" w:line-rule="at-least"/>
                        <w:ind w:first-line="360"/>
                        <w:jc w:val="both"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="DocumentLabel">
                    <w:name w:val="Document Label"/>
                    <w:next w:val="Normal"/>
                    <w:pPr>
                        <w:pStyle w:val="DocumentLabel"/>
                        <w:pBdr>
                            <w:top w:val="double" w:sz="6" w:space="8" w:color="808080"/>
                            <w:bottom w:val="double" w:sz="6" w:space="8" w:color="808080"/>
                        </w:pBdr>
                        <w:spacing w:after="40" w:line="240" w:line-rule="at-least"/>
                        <w:jc w:val="center"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Garamond" w:h-ansi="Garamond"/>
                        <w:b/>
                        <w:caps/>
                        <w:spacing w:val="20"/>
                        <w:sz w:val="18"/>
                        <w:lang w:val="EN-US" w:fareast="EN-US" w:bidi="AR-SA"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeader">
                    <w:name w:val="Message Header"/>
                    <w:basedOn w:val="BodyText"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeader"/>
                        <w:keepLines/>
                        <w:spacing w:after="120"/>
                        <w:ind w:left="1080" w:hanging="1080"/>
                        <w:jc w:val="left"/>
                    </w:pPr>
                    <w:rPr>
                        <w:caps/>
                        <w:sz w:val="18"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeaderFirst">
                    <w:name w:val="Message Header First"/>
                    <w:basedOn w:val="MessageHeader"/>
                    <w:next w:val="MessageHeader"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeaderFirst"/>
                        <w:spacing w:before="360"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="character" w:styleId="MessageHeaderLabel">
                    <w:name w:val="Message Header Label"/>
                    <w:rPr>
                        <w:b/>
                        <w:sz w:val="18"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeaderLast">
                    <w:name w:val="Message Header Last"/>
                    <w:basedOn w:val="MessageHeader"/>
                    <w:next w:val="BodyText"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeaderLast"/>
                        <w:pBdr>
                            <w:bottom w:val="single" w:sz="6" w:space="18" w:color="808080"/>
                        </w:pBdr>
                        <w:spacing w:after="360"/>
                    </w:pPr>
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
                <w:r>
                    <w:t>interoffice memorandum</w:t>
                </w:r>
            </w:p>
            
            <!-- This paragraph will be the "to:" line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeaderFirst"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t>to:</w:t>
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
            
            <!-- This paragraph will be the "from:" line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeader"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t>from:</w:t>
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
            
            <!-- This paragraph will be the "Priority:" line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeader"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t>Priority:</w:t>
                </w:r>
                <w:r>
                    <w:t xml:space="preserve">	</w:t>
                </w:r>
                <ns0:priority>
                    <w:r>
                        <w:t>
                            <xsl:value-of select="ns0:priority"/>
                        </w:t>
                    </w:r>
                </ns0:priority>
            </w:p>
            
            <!-- This paragraph will be the "Subject:" line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeader"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t>Subject:</w:t>
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
            
            <!-- This paragraph will be the "date:" line -->
            
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeaderLast"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                    </w:rPr>
                    <w:t>date:</w:t>
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
            
            <!-- This paragraph will be the body of the memo -->
            
            <ns0:body>
                <w:p>
                    <w:pPr>
                        <w:pStyle w:val="BodyText"/>
                        <w:ind w:right="25"/>
                    </w:pPr>
                    <w:r>
                        <w:t>
                            <xsl:value-of select="ns0:body"/>
                        </w:t>
                    </w:r>
                </w:p>
            </ns0:body>
        </ns0:memo>
    </xsl:template>
</xsl:stylesheet>