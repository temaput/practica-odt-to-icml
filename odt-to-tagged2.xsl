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

    <xsl:output method="text" encoding="utf-16le"/>

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
        <xsl:text>&lt;UNICODE-WIN&gt;&#xD;&#xA;&lt;Version:8&gt;</xsl:text>
        <!-- go pick all styles and make style definitions at the begining -->
        <xsl:apply-templates 
            select="document($styles_file)/office:document-styles/office:styles/style:style"/>
        <!-- main processing -->
        <xsl:apply-templates select="office:body/office:text"/>
    </xsl:template>

    <xsl:template match="office:text">
        <xsl:apply-templates select="table:table | text:h | text:list | text:p"/>
    </xsl:template>


    <!-- ===========================list processing =============  -->
    <xsl:template match="text:list">
        <xsl:apply-templates select="descendant::text:list-item"/>
    </xsl:template>

    <xsl:template match="text:list-item">
        <!-- here we could do something to preserve bullets and numbers but
             instead we just do nothing here -->
        <xsl:apply-templates select="text:p | text:h"/>
    </xsl:template>


    <!-- ========================== text processing =================== -->
    <xsl:template match="text:p | text:h">
        <!-- processing main para tags -->
        <!-- insert return if the closest preceding is also para or heading or list -->
        <xsl:if test="preceding-sibling::*[1][self::text:p|self::text:h|self::text:list]"><xsl:text>&#xD;&#xA;</xsl:text></xsl:if>
        <!-- also insert CR before every first para in list-item -->
        <xsl:if test="parent::text:list-item and count(parent::text:list-item/*[1]|.)=1">
            <xsl:text>&#xD;&#xA;</xsl:text>
            <xsl:message><xsl:value-of select="."/></xsl:message>
        </xsl:if>
        <xsl:text>&lt;ParaStyle:</xsl:text>

        <xsl:call-template name="getParaStyle">
            <xsl:with-param name="style-name" select="@text:style-name"/>
        </xsl:call-template>

        <xsl:text>&gt;</xsl:text>
        <xsl:apply-templates/>
    </xsl:template>

    <xsl:template match="text:span">
        <!-- processing span tags -->
        <xsl:call-template name="getCustomFormatting"/>
        <xsl:apply-templates/>
        <xsl:call-template name="getCustomFormatting">
            <xsl:with-param name="clear" select="'1'"/>
        </xsl:call-template>
    </xsl:template>

    <xsl:template match="text:alphabetical-index-mark">
        <!-- index mark -->
        <!-- what we receive:
            <text:alphabetical-index-mark text:string-value="Большой индекс" text:key1="Раздел" text:key2="Подраздел"/>
             what we create here:
            <IndexEntry:=<IndexEntryType:IndexPageEntry><IndexEntryRangeType:kCurrentPage><IndexEntryDisplayString:Чубайс>>
        -->

        <xsl:text>&lt;IndexEntry:=&lt;IndexEntryType:IndexPageEntry&gt;</xsl:text>
        <xsl:text>&lt;IndexEntryRangeType:kCurrentPage&gt;</xsl:text>
        <!-- lets break apart the 3 levels -->
        <xsl:if test="@text:key1">
            <xsl:text>&lt;IndexEntryDisplayString:</xsl:text>
            <xsl:value-of select="@text:key1"/>
            <xsl:text>&gt;</xsl:text>
        </xsl:if>
        <xsl:if test="@text:key2">
            <xsl:text>&lt;IndexEntryDisplayString:</xsl:text>
            <xsl:value-of select="@text:key2"/>
            <xsl:text>&gt;</xsl:text>
        </xsl:if>
        <xsl:text>&lt;IndexEntryDisplayString:</xsl:text>
        <xsl:value-of select="@text:string-value"/>
        <xsl:text>&gt;&gt;</xsl:text>
    </xsl:template>

    <xsl:template match="text:line-break">
        <xsl:text>&#xA;</xsl:text>
    </xsl:template>
    <xsl:template match="text:tab">
        <xsl:text>&#09;</xsl:text>
    </xsl:template>
    <xsl:template match="text:s">
            <xsl:call-template name="printWhitespaces">
                <xsl:with-param name="count" select="number(@text:c)"/>
            </xsl:call-template>
    </xsl:template>

    <xsl:template match="text()">
        <xsl:variable name="sub1">
            <xsl:call-template name="search-and-replace">
                <xsl:with-param name="input" select="."/>
                <xsl:with-param name="search-string" select="'\'"/>
                <xsl:with-param name="replace-string" select="'\\'"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="sub2">
            <xsl:call-template name="search-and-replace">
                <xsl:with-param name="input" select="$sub1"/>
                <xsl:with-param name="search-string" select="'&gt;'"/>
                <xsl:with-param name="replace-string" select="'\&gt;'"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="search-and-replace">
            <xsl:with-param name="input" select="$sub2"/>
            <xsl:with-param name="search-string" select="'&lt;'"/>
            <xsl:with-param name="replace-string" select="'\&lt;'"/>
        </xsl:call-template>
    </xsl:template>

    <!-- ======================= FINISHED text processing ============== -->

    <!-- ======================= table processing ===================== -->
    
    
    <xsl:template match="table:table">
        <!-- processing table start tags 
             what we expect is:
            <TableStart:2,2:0:0<tCellDefaultCellType:Text>><ColStart:<tColAttrWidth:269.5>><ColStart:<tColAttrWidth:269.5>>
        -->
        <!-- we need to count columns right -->
        <xsl:variable name="colCount" select="count(table:table-column[not(@table:number-columns-repeated)])"/>
        <xsl:variable name="repeatedColSum">
            <xsl:value-of select="sum(table:table-column/@table:number-columns-repeated)"/>
        </xsl:variable>

        <xsl:text>&lt;TableStart:</xsl:text>
        <xsl:value-of select="count(table:table-row)"/>
        <xsl:text>,</xsl:text>
        <xsl:value-of select="$colCount + $repeatedColSum"/>
        <xsl:text>&gt;</xsl:text>

        <!-- go process the rows -->
        <xsl:apply-templates select="table:table-row"/>
        <!-- finishing table -->
        <xsl:text>&lt;TableEnd:&gt;&#xD;&#xA;</xsl:text>
    </xsl:template>

    <xsl:template match="table:table-row">
        <xsl:text>&lt;RowStart:&gt;</xsl:text>
        <xsl:apply-templates select="table:table-cell | table:covered-table-cell">
        </xsl:apply-templates>
        <xsl:text>&lt;RowEnd:&gt;</xsl:text>
    </xsl:template>

    <xsl:template match="table:table-cell | table:covered-table-cell">
        <xsl:if test="not(name()='table:covered-table-cell')">
            <!-- we are in table-cell -->
            <xsl:text>&lt;CellStart:</xsl:text>
                <xsl:if test="@table:number-rows-spanned">
                    <xsl:value-of
                        select="@table:number-rows-spanned"/>
                </xsl:if>
                <xsl:if test="not(@table:number-rows-spanned)">
                    <xsl:text>1</xsl:text>
                </xsl:if>
                <xsl:text>,</xsl:text>
                <xsl:if test="@table:number-columns-spanned">
                    <xsl:value-of
                        select="@table:number-columns-spanned"/>
                </xsl:if>
                <xsl:if test="not(@table:number-columns-spanned)">
                    <xsl:text>1</xsl:text>
                </xsl:if>
            <xsl:text>&gt;</xsl:text>

            <xsl:apply-templates/>
            <xsl:text>&lt;CellEnd:&gt;</xsl:text>
        </xsl:if>
        <xsl:if test="name()='table:covered-table-cell'">
            <xsl:text>&lt;CellStart:1,1&gt;&lt;CellEnd:&gt;</xsl:text>
        </xsl:if>
    </xsl:template>
    
    

    <!-- ======================= FINISHED table processing ===================== -->



    <!-- =========================== utilities =========================== -->
    <xsl:template name="printWhitespaces">
        <xsl:param name="count"/>
        <xsl:text> </xsl:text>
        <xsl:if test="$count &gt; 1">
            <xsl:call-template name="printWhitespaces">
                <xsl:with-param name="count" select="$count - 1"/>
            </xsl:call-template>
        </xsl:if>
    </xsl:template>
    <xsl:template name="getCustomFormatting">
        <!-- checking for accidental formatting like Bold, Subscript -->
        <xsl:param name="clear"/>

        <!-- first we find out the style name and the parent (Paragraph) style -->
        <xsl:variable name="styleName">
            <xsl:choose>
                <xsl:when test="self::text()">
                    <xsl:value-of select="parent::*[1]/@text:style-name"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="@text:style-name"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="parentStyleName">
            <xsl:value-of select="parent::*[1][self::text:p|self::text:h]/@text:style-name"/>
        </xsl:variable>

        <!-->
            <xsl:message>
                <xsl:text>tag is: </xsl:text>
                <xsl:value-of select="name(.)"/>
                <xsl:text>&#10;styleName: </xsl:text>
                <xsl:value-of select="$styleName"/>
                <xsl:text>&#10;parentStyleName: </xsl:text>
                <xsl:value-of select="$parentStyleName"/>
            </xsl:message>
        </!-->

        <!-- now we aquire values and pick the right (recent) ones -->
        <xsl:if test="key('auto-styles', $styleName)">
            <xsl:variable name="fname">
                <!-- checking for symbol, not implemented yet -->
                <xsl:value-of
                    select="key('auto-styles', $styleName)/style:text-properties/@style:font-name"/>
            </xsl:variable>  
            <xsl:variable name="fposition">
                <!-- subscript or super -->
                <xsl:value-of 
                    select="concat(
                    key('auto-styles', $styleName)/style:text-properties/@style:text-position,
                    key('auto-styles', $parentStyleName)/style:text-properties/@style:text-position)"/>
            </xsl:variable>
            <xsl:variable name="funderline">
                <xsl:value-of 
                    select="concat(
                    key('auto-styles', $styleName)/style:text-properties/@style:text-underline-style,
                    key('auto-styles', $parentStyleName)/style:text-properties/@style:text-underline-style)"/>
            </xsl:variable>
            <xsl:variable name="fweight">
                <xsl:value-of 
                    select="concat(
                    key('auto-styles', $styleName)/style:text-properties/@fo:font-weight,
                    key('auto-styles', $parentStyleName)/style:text-properties/@fo:font-weight)"/>
            </xsl:variable>
            <xsl:variable name="fstyle">
                <xsl:value-of 
                    select="concat(
                    key('auto-styles', $styleName)/style:text-properties/@fo:font-style,
                    key('auto-styles', $parentStyleName)/style:text-properties/@fo:font-style)"/>
            </xsl:variable>
            <xsl:variable name="fbold">
                <xsl:value-of select="starts-with($fweight, 'bold')"/>
            </xsl:variable>
            <xsl:variable name="fitalic">
                <xsl:value-of select="starts-with($fstyle,'italic') 
                    or starts-with($fstyle, 'oblique')"/>
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

            <!-- Here we print the aquired values -->
            <xsl:if test="string-length($font-style)">
                <xsl:text>&lt;cTypeface:</xsl:text>
                <xsl:if test="not($clear)">
                    <!-- if clear is set then we should give empty tag here -->
                    <xsl:value-of select="$font-style"/>
                </xsl:if>
                <xsl:text>&gt;</xsl:text>
            </xsl:if>
            <xsl:if test="string-length($funderline)">
                <xsl:if test="not(starts-with($funderline, 'none'))">
                    <xsl:text>&lt;cUnderline:</xsl:text>
                    <xsl:if test="not($clear)">
                        <!-- if clear is set then we should give empty tag here -->
                        <xsl:text>1</xsl:text>
                    </xsl:if>
                    <xsl:text>&gt;</xsl:text>
                </xsl:if>
            </xsl:if>
            <xsl:if test="string-length($fposition)">
                <xsl:text>&lt;cPosition:</xsl:text>
                <xsl:if test="not($clear)">
                    <xsl:choose>
                        <xsl:when test="starts-with($fposition, 'sub') or 
                            number(substring-before($fposition,'%')) &lt; 0">
                            <xsl:text>Subscript</xsl:text>
                        </xsl:when>
                        <xsl:when test="starts-with($fposition, 'super') or
                            number(substring-before($fposition,'%')) &gt; 0">
                            <xsl:text>Superscript</xsl:text>
                        </xsl:when>
                    </xsl:choose>
                </xsl:if>
                <xsl:text>&gt;</xsl:text>
            </xsl:if>
            <!-->
                <xsl:if test="$fname='Symbol'">
                    <Properties>
                            <AppliedFont type="string">Symbol</AppliedFont>
                    </Properties>
                </xsl:if>
            </!-->
        </xsl:if>
    </xsl:template>

    <xsl:template name="getParaStyle">
        <!-- returns the style name or parent style name if auto-style -->
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

    <xsl:template name="search-and-replace">
        <xsl:param name="input"/>
        <xsl:param name="search-string"/>
        <xsl:param name="replace-string"/>
        <xsl:choose>
            <!-- See if the input contains the search string -->
            <xsl:when test="$search-string and 
                            contains($input,$search-string)">
            <!-- If so, then concatenate the substring before the search
            string to the replacement string and to the result of
            recursively applying this template to the remaining substring.
            -->
                <xsl:value-of 
                        select="substring-before($input,$search-string)"/>
                <xsl:value-of select="$replace-string"/>
                <xsl:call-template name="search-and-replace">
                        <xsl:with-param name="input"
                        select="substring-after($input,$search-string)"/>
                        <xsl:with-param name="search-string" 
                        select="$search-string"/>
                        <xsl:with-param name="replace-string" 
                            select="$replace-string"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <!-- There are no more occurrences of the search string so 
                just return the current input string -->
                <xsl:value-of select="$input"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    <!-- ========================= FINISHED utilities =================== -->
        


    <!-- ============================ getting styles ====================   -->
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
                <xsl:text>&#xD;&#xA;&lt;DefineParaStyle:</xsl:text>
                <xsl:value-of select="$style_name"/>
                <xsl:text>&gt;</xsl:text>
            </xsl:when>
            <xsl:otherwise>
                <xsl:if test="key('paragraph-parent-styles', $style_name)">
                    <xsl:text>&#xD;&#xA;&lt;DefineParaStyle:</xsl:text>
                    <xsl:value-of select="$style_name"/>
                    <xsl:text>&gt;</xsl:text>
                </xsl:if>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    <!--  ======================== FINISHED getting styles ================ -->
</xsl:stylesheet>
