<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"   
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" 
    exclude-result-prefixes="office text style fo"
    >

    <xsl:output method="xml" indent="yes" encoding="utf-8"/>

    <xsl:key name="auto-styles" match="style:style" use="@style:name"/>
    <xsl:key name="paragraph-styles" match="text:h | text:p" use="@text:style-name"/>
    <xsl:key name="paragraph-parent-styles" match="style:style"  
        use="@style:parent-style-name"/>

    <xsl:variable name="styles_file" select="'styles.xml'"/>
    <xsl:variable name="content_file" select="/"/>

    <xsl:template match="/">
        <xsl:apply-templates select="/office:document-content"/>
    </xsl:template>
    
    <!-- =================== main() ===================================== -->
    <xsl:template match="/office:document-content">
        <xsl:processing-instruction name="aid">style="50" type="snippet" readerVersion="6.0" featureSet="513" product="8.0(370)"</xsl:processing-instruction>
        <xsl:processing-instruction name="aid">SnippetType="InCopyInterchange"</xsl:processing-instruction>
        <Document>
            <xsl:attribute name="DOMVersion" select="'8.0'"/>
            <xsl:attribute name="Self" select="'d'"/>
            <RootCharacterStyleGroup Self="u77">
                <CharacterStyle Self="CharacterStyle/$ID/[No character style]" Imported="false" Name="$ID/[No character style]" />
            </RootCharacterStyleGroup>
            <RootParagraphStyleGroup >
                <xsl:apply-templates 
                    select="document($styles_file)/office:document-styles/office:styles/style:style"/>
            </RootParagraphStyleGroup>

            <Story>
                <xsl:attribute name="Self" select="'sdf4'"/>
                <xsl:apply-templates select="office:body/office:text"/>
            </Story>
        </Document>
    </xsl:template>


    <!-- =================== getting contents =========================== -->
    <xsl:template match="office:text">
        <xsl:apply-templates select="table:table | text:h | text:p"/>
    </xsl:template>


    <xsl:template match="text:p | text:h">
        <ParagraphStyleRange>
            <xsl:attribute name="AppliedParagraphStyle">

                <xsl:text>ParagraphStyle/</xsl:text>
                <xsl:call-template name="getParaStyle">
                    <xsl:with-param name="style-name" select="@text:style-name"/>
                </xsl:call-template>
            </xsl:attribute>

            <xsl:call-template name="getCustomFormatting">
                <xsl:with-param name="styleName" select="@text:style-name"/>
            </xsl:call-template>

            <xsl:apply-templates/>
            <xsl:if test="not(position()=last())">
                <Br/>
            </xsl:if>
        </ParagraphStyleRange>
    </xsl:template>

    <xsl:template match="table:table">
        <Table>
            <xsl:attribute name="BodyRowCount">
                <xsl:value-of select="count(table:table-row)"/>
            </xsl:attribute>
            <xsl:attribute name="ColumnCount">
                <xsl:value-of select="count(table:table-column)"/>
            </xsl:attribute>
            <xsl:apply-templates select="table:table-row"/>
        </Table>
    </xsl:template>

    <xsl:template match="table:table-row">
        <xsl:variable name="rowNum" select="position()"/>
        <xsl:apply-templates select="table:table-cell | table:covered-table-cell">
            <xsl:with-param name="rowNum" select="$rowNum"/>
        </xsl:apply-templates>
    </xsl:template>

    <xsl:template match="table:table-cell | table:covered-table-cell">
        <xsl:param name="rowNum"/>
        <xsl:variable name="colNum" select="position()"/>
        <xsl:if test="not(name()='table:covered-table-cell')">
            <!-- we are in table-cell -->
            <Cell>
                <xsl:attribute name="Name">
                    <xsl:value-of select="$colNum - 1"/>
                    <xsl:text>:</xsl:text>
                    <xsl:value-of select="$rowNum - 1"/>
                </xsl:attribute>
                <xsl:if test="@table:number-rows-spanned">
                    <xsl:attribute name="RowSpan" 
                        select="@table:number-rows-spanned"/>
                </xsl:if>
                <xsl:if test="@table:number-columns-spanned">
                    <xsl:attribute name="ColumnSpan" 
                        select="@table:number-columns-spanned"/>
                </xsl:if>
                <xsl:apply-templates/>
            </Cell>
        </xsl:if>
    </xsl:template>



    <xsl:template match="text:span">
        <CharacterStyleRange AppliedCharacterStyle="CharacterStyle/$ID/[No character style]">
            <xsl:call-template name="getCustomFormatting">
                <xsl:with-param name="styleName" select="@text:style-name"/>
            </xsl:call-template>

            <xsl:apply-templates/>
        </CharacterStyleRange>
    </xsl:template>
    <xsl:template match="text:line-break">
        <Content><xsl:text>&#2028;</xsl:text></Content>
    </xsl:template>

    <xsl:template match="text()">
        <Content>
            <xsl:value-of select="."/>
        </Content>
    </xsl:template>

    <!-- ============================ getting styles ====================   -->
    <xsl:template name="getCustomFormatting">
        <xsl:param name="styleName"/>
        <xsl:if test="key('auto-styles', $styleName)">
            <xsl:variable name="fname">
                <xsl:value-of
                    select="key('auto-styles', $styleName)/style:text-properties/@style:font-name"/>
            </xsl:variable>  
            <xsl:variable name="fposition">
                <xsl:value-of 
                    select="key('auto-styles', $styleName)/style:text-properties/@style:text-position"/>
            </xsl:variable>
            <xsl:variable name="funderline">
                <xsl:value-of 
                    select="key('auto-styles', $styleName)/style:text-properties/@style:text-underline-style"/>
            </xsl:variable>
            <xsl:variable name="fweight">
                <xsl:value-of 
                    select="key('auto-styles', $styleName)/style:text-properties/@fo:font-weight"/>
            </xsl:variable>
            <xsl:variable name="fstyle">
                <xsl:value-of 
                    select="key('auto-styles', $styleName)/style:text-properties/@fo:font-style"/>
            </xsl:variable>
            <xsl:variable name="fbold">
                <xsl:value-of select="$fweight='bold'"/>
            </xsl:variable>
            <xsl:variable name="fitalic">
                <xsl:value-of select="$fstyle='italic' 
                    or $fstyle='oblique'"/>
            </xsl:variable>

            <xsl:variable name="font-style">
                <xsl:choose>
                    <xsl:when test="$fbold='true' and $fitalic='true'">
                        <xsl:text>Bold Italic</xsl:text>
                    </xsl:when>

                    <xsl:when test="$fbold='true'">
                        <xsl:text>Bold</xsl:text>
                    </xsl:when>

                    <xsl:when test="$fitalic='true'">
                        <xsl:text>Italic</xsl:text>
                    </xsl:when>
                </xsl:choose>
            </xsl:variable>
            <xsl:if test="string-length($font-style)">
                    <xsl:attribute name="FontStyle" select="$font-style"/>
            </xsl:if>
            <xsl:if test="string-length($funderline)">
                <xsl:attribute name="Underline" select="'true'"/>
            </xsl:if>
            <xsl:if test="string-length($fposition)">
                <xsl:choose>
                    <xsl:when test="starts-with($fposition, 'sub')">
                        <xsl:attribute name="Position" select="'Subscript'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($fposition, 'super')">
                        <xsl:attribute name="Position" select="'Superscript'"/>
                    </xsl:when>
                </xsl:choose>
            </xsl:if>
            <xsl:if test="$fname='Symbol'">
                <Properties>
                        <AppliedFont type="string">Symbol</AppliedFont>
                </Properties>
            </xsl:if>

        </xsl:if>
    </xsl:template>

    <xsl:template name="getParaStyle">
        <xsl:param name="style-name"/>
        <xsl:choose>
            
            <xsl:when test="key('auto-styles', $style-name)">
                <xsl:value-of 
                    select="key('auto-styles', $style-name)/@style:parent-style-name"/>
            </xsl:when>

            <xsl:otherwise>
                <xsl:value-of select="$style-name"/>
            </xsl:otherwise>

        </xsl:choose>
    </xsl:template>
        
    <xsl:template match="office:document-styles/office:styles/style:style">
        <!-- we should be inside the styles.xml now -->
        <xsl:variable name="style_name" >
            <xsl:value-of select="string(@style:name)"/>
        </xsl:variable>
        <xsl:apply-templates select="$content_file/*/office:body">
            <xsl:with-param name="style_name" select="$style_name"/>
        </xsl:apply-templates>
    </xsl:template>

    <xsl:template match="office:body">
        <!-- just looking for used styles here -->
        <xsl:param name="style_name"/>
        <xsl:choose>
            <xsl:when test="key('paragraph-styles', $style_name)">

                <ParagraphStyle>
                    <xsl:attribute name="Name" select="$style_name"/>
                    <xsl:attribute name="Self" 
                        select="concat('ParagraphStyle/', $style_name)"/>
                </ParagraphStyle>

            </xsl:when>
            <xsl:otherwise>
                <xsl:if test="key('paragraph-parent-styles', $style_name)">
                    <ParagraphStyle>
                        <xsl:attribute name="Name" select="$style_name"/>
                        <xsl:attribute name="Self" 
                            select="concat('ParagraphStyle/', $style_name)"/>
                    </ParagraphStyle>
                </xsl:if>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
