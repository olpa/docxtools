<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0"
    xmlns:xsl		= "http://www.w3.org/1999/XSL/Transform"
    xmlns:xs		= "http://www.w3.org/2001/XMLSchema"
    xmlns:xsldoc	= "http://www.bacman.net/XSLdoc"
    xmlns:saxon		= "http://saxon.sf.net/"
    xmlns:letex		= "http://www.le-tex.de/namespace"
    xmlns:saxExtFn	= "java:saxonExtensionFunctions"
    xmlns:hub		= "http://www.le-tex.de/namespace/hub"
    xmlns:xlink		= "http://www.w3.org/1999/xlink"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:css           = "http://www.w3.org/1996/css"
    xmlns:o		= "urn:schemas-microsoft-com:office:office"
    xmlns:w		= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:m		= "http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:wp		= "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:r		= "http://schemas.openxmlformats.org/package/2006/relationships"

    xpath-default-namespace = "http://docbook.org/ns/docbook"

    exclude-result-prefixes = "xsl xs xsldoc saxon letex saxExtFn hub xlink o w m wp r css"
>


<!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
<!-- mode="hub:default" -->
<!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <xsl:template  match="bibliography"  mode="hub:default" >
    <xsl:apply-templates  select="node()"  mode="#current" />
  </xsl:template>


  <xsl:template  match="bibliography/title"  mode="hub:default" >
    <xsl:variable name="pPr" as="element(*)*">
      <xsl:apply-templates  select="@css:page-break-after, @css:page-break-inside, @css:page-break-before, @css:text-indent, (@css:widows, @css:orphans)[1], @css:margin-bottom, @css:margin-top, @css:line-height, @css:text-align"  mode="props" />
      <w:pStyle w:val="Heading{letex:headinglevel(ancestor::*[self::section or self::chapter][1]/title) + 1}"/>
    </xsl:variable>
    <w:p>
      <xsl:if  test="$pPr">
        <w:pPr>
          <xsl:sequence  select="$pPr" />
        </w:pPr>
      </xsl:if>
      <xsl:apply-templates  select="node()"  mode="#current" />
    </w:p>
  </xsl:template>


  <xsl:template  match="bibliodiv"  mode="hub:default" >
    <xsl:apply-templates  mode="#current" />
  </xsl:template>


  <xsl:template  match="bibliodiv/title"  mode="hub:default" >
    <xsl:variable name="pPr" as="element(*)*">
      <xsl:apply-templates  select="@css:page-break-after, @css:page-break-inside, @css:page-break-before, @css:text-indent, (@css:widows, @css:orphans)[1], @css:margin-bottom, @css:margin-top, @css:line-height, @css:text-align"  mode="props" />
      <w:pStyle w:val="Heading{letex:headinglevel(ancestor::*[self::section or self::chapter][1]/title) + 2}"/>
    </xsl:variable>
    <w:p>
      <xsl:if  test="$pPr">
        <w:pPr>
          <xsl:sequence  select="$pPr" />
        </w:pPr>
      </xsl:if>
      <xsl:apply-templates  select="node()"  mode="#current" />
    </w:p>
  </xsl:template>


  <xsl:template  match="bibliolist"     mode="hub:default">
    <xsl:apply-templates mode="#current" />
  </xsl:template>

  <xsl:template  match="biblioentry" mode="hub:default">
    <xsl:apply-templates mode="#current" />
  </xsl:template>

  <xsl:template  match="bibliomixed"  mode="hub:default">
    <xsl:variable  name="bibliomixedId"  select="generate-id()"/>
    <w:bookmarkStart w:id="{$bibliomixedId}" w:name="bm_{$bibliomixedId}_"/>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="bib"/>
      </w:pPr>
      <xsl:apply-templates  select="node()"  mode="#current" />
    </w:p>
    <w:bookmarkEnd w:id="{$bibliomixedId}"/>
  </xsl:template>

  <xsl:template  match="bibliomisc"    mode="hub:default">
    <w:p>
      <w:pPr>
        <xsl:choose>
          <xsl:when  test="@role eq 'numberedRef'">
            <w:pStyle w:val="bibnum"/>
            <w:numPr>
              <w:ilvl w:val="0"/>
              <w:numId w:val="{letex:getNumId( ancestor::*[self::*[ local-name() = ( 'itemizedlist' , 'orderedlist', 'bibliography', 'bibliodiv' )]][1]/generate-id() )}" />
            </w:numPr>
          </xsl:when>
          <xsl:otherwise>
            <w:pStyle w:val="bib"/>
          </xsl:otherwise>
        </xsl:choose>
      </w:pPr>
      <xsl:apply-templates  select="node()"  mode="#current" />
    </w:p>
  </xsl:template>


</xsl:stylesheet>
