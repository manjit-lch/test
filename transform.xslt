<?xml version="1.0" ?>
<xsl:stylesheet version="1.0" 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:office:xslt"
    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:html="http://www.w3.org/TR/REC-html40">

    <xsl:key name="names" match="/*/*/*" use="local-name(.)"/>

    <xsl:template match="/*">
        <ss:Workbook>
            <ss:Worksheet ss:Name="hello">
                <ss:Table>
                    <ss:Row>
                        <xsl:for-each select="/*/*/*[generate-id(.) = generate-id(key('names', local-name(.)))]">
                            <ss:Cell>
                                <ss:Data ss:Type="String"><xsl:value-of select="local-name(.)" /></ss:Data>
                            </ss:Cell>
                        </xsl:for-each>
                    </ss:Row>
                    <xsl:apply-templates/>
                </ss:Table>
            </ss:Worksheet>
        </ss:Workbook>
    </xsl:template>

    <xsl:template match="@*|node()">
        <xsl:choose>
            <!-- If the element has multiple childs, call this template 
                on its children to flatten it-->
            <xsl:when test="count(child::*) > 0">
                <ss:Row>
                    <xsl:apply-templates select="@*|node()"/>
                </ss:Row>
            </xsl:when>
            <xsl:otherwise>
                <ss:Cell>
                    <ss:Data ss:Type="String"><xsl:value-of select="text()" /></ss:Data>
                </ss:Cell>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>

