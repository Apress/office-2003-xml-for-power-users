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
                    <w:rsid w:val="C93968"/>
                    <w:pPr>
                        <w:ind w:left="835" w:right="835"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:h-ansi="Arial"/>
                        <w:spacing w:val="-5"/>
                        <w:lang w:val="EN-US" w:fareast="EN-US" w:bidi="AR-SA"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="character" w:default="on" w:styleId="DefaultParagraphFont">
                    <w:name w:val="Default Paragraph Font"/>
                    <w:semiHidden/>
                    <w:rsid w:val="C93968"/>
                </w:style>
                
                <w:style w:type="table" w:default="on" w:styleId="TableNormal">
                    <w:name w:val="Normal Table"/>
                    <w:semiHidden/>
                    <w:rsid w:val="C93968"/>
                    <w:tblPr>
                        <w:tblInd w:w="0" w:type="dxa"/>
                        <w:tblCellMar>
                            <w:top w:w="0" w:type="dxa"/>
                            <w:left w:w="108" w:type="dxa"/>
                            <w:bottom w:w="0" w:type="dxa"/>
                            <w:right w:w="108" w:type="dxa"/>
                        </w:tblCellMar>
                    </w:tblPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="BodyText">
                    <w:name w:val="Body Text"/>
                    <w:basedOn w:val="Normal"/>
                    <w:rsid w:val="C93968"/>
                    <w:pPr>
                        <w:pStyle w:val="BodyText"/>
                        <w:spacing w:after="220" w:line="180" w:line-rule="at-least"/>
                        <w:jc w:val="both"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="CompanyName">
                    <w:name w:val="Company Name"/>
                    <w:basedOn w:val="Normal"/>
                    <w:rsid w:val="C93968"/>
                    <w:pPr>
                        <w:pStyle w:val="CompanyName"/>
                        <w:keepLines/>
                        <w:shd w:val="solid" w:color="auto" w:fill="auto"/>
                        <w:spacing w:line="320" w:line-rule="exact"/>
                        <w:ind w:left="0"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial Black" w:h-ansi="Arial Black"/>
                        <w:color w:val="FFFFFF"/>
                        <w:spacing w:val="-15"/>
                        <w:sz w:val="32"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="DocumentLabel">
                    <w:name w:val="Document Label"/>
                    <w:basedOn w:val="Normal"/>
                    <w:next w:val="Normal"/>
                    <w:rsid w:val="C93968"/>
                    <w:pPr>
                        <w:pStyle w:val="DocumentLabel"/>
                        <w:keepNext/>
                        <w:keepLines/>
                        <w:spacing w:before="400" w:after="120" w:line="240" w:line-rule="at-least"/>
                        <w:ind w:left="0"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial Black" w:h-ansi="Arial Black"/>
                        <w:kern w:val="28"/>
                        <w:sz w:val="96"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeader">
                    <w:name w:val="Message Header"/>
                    <w:basedOn w:val="BodyText"/>
                    <w:rsid w:val="C93968"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeader"/>
                        <w:keepLines/>
                        <w:spacing w:after="120"/>
                        <w:ind w:left="1555" w:hanging="720"/>
                        <w:jc w:val="left"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeaderFirst">
                    <w:name w:val="Message Header First"/>
                    <w:basedOn w:val="MessageHeader"/>
                    <w:next w:val="MessageHeader"/>
                    <w:rsid w:val="C93968"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeaderFirst"/>
                        <w:spacing w:before="220"/>
                    </w:pPr>
                </w:style>
                
                <w:style w:type="character" w:styleId="MessageHeaderLabel">
                    <w:name w:val="Message Header Label"/>
                    <w:rsid w:val="C93968"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial Black" w:h-ansi="Arial Black"/>
                        <w:spacing w:val="-10"/>
                        <w:sz w:val="18"/>
                    </w:rPr>
                </w:style>
                
                <w:style w:type="paragraph" w:styleId="MessageHeaderLast">
                    <w:name w:val="Message Header Last"/>
                    <w:basedOn w:val="MessageHeader"/>
                    <w:next w:val="BodyText"/>
                    <w:rsid w:val="C93968"/>
                    <w:pPr>
                        <w:pStyle w:val="MessageHeaderLast"/>
                        <w:pBdr>
                            <w:bottom w:val="single" w:sz="6" w:space="15" w:color="auto"/>
                        </w:pBdr>
                        <w:spacing w:after="320"/>
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
                <xsl:apply-templates select="ns0:memo"/>
            </w:body>
            
        </w:wordDocument>
    </xsl:template>
    <xsl:template match="ns0:memo">
        <ns0:memo>
        
			<!-- This memo uses a table for much of the layout  -->
			
            <w:tbl>
                <w:tblPr>
                    <w:tblW w:w="0" w:type="auto"/>
                    <w:tblInd w:w="835" w:type="dxa"/>
                    <w:tblCellMar>
                        <w:left w:w="187" w:type="dxa"/>
                        <w:right w:w="187" w:type="dxa"/>
                    </w:tblCellMar>
                </w:tblPr>
                <w:tblGrid>
                    <w:gridCol w:w="5040"/>
                    <w:gridCol w:w="3845"/>
                </w:tblGrid>
                <w:tr>
                    <w:trPr>
                        <w:trHeight w:val="720"/>
                    </w:trPr>
                    <w:tc>
                        <w:tcPr>
                            <w:tcW w:w="5040" w:type="dxa"/>
                            <w:tcMar>
                                <w:left w:w="0" w:type="dxa"/>
                                <w:right w:w="0" w:type="dxa"/>
                            </w:tcMar>
                        </w:tcPr>
                        <w:p>
                            <w:pPr>
                                <w:pStyle w:val="ReturnAddress"/>
                            </w:pPr>
                        </w:p>
                    </w:tc>
                    <w:tc>
                        <w:tcPr>
                            <w:tcW w:w="3845" w:type="dxa"/>
                            <w:shd w:val="solid" w:color="auto" w:fill="auto"/>
                            <w:vAlign w:val="center"/>
                        </w:tcPr>
                        <w:p>
                            <w:pPr>
                                <w:pStyle w:val="CompanyName"/>
                                <w:ind w:right="25"/>
                            </w:pPr>
                            <w:r>
                                <w:t>Microsoft Corporation </w:t>
                            </w:r>
                        </w:p>
                    </w:tc>
                </w:tr>
            </w:tbl>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="DocumentLabel"/>
                </w:pPr>
                <w:r>
                    <w:t>Memo</w:t>
                </w:r>
            </w:p>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeaderFirst"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
                        <w:spacing w:val="-25"/>
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
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="MessageHeader"/>
                    <w:rPr>
                        <w:rStyle w:val="MessageHeaderLabel"/>
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
                            <w:rFonts w:ascii="Arial" w:h-ansi="Arial" w:cs="Arial"/>
                            <w:sz w:val="20"/>
                        </w:rPr>
                        <w:t>
                            <xsl:value-of select="ns0:priority"/>
                        </w:t>
                    </w:r>
                </ns0:priority>
            </w:p>
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