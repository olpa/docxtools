<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"

  xmlns:xsl		= "http://www.w3.org/1999/XSL/Transform"
  xmlns:fn              = "http://www.w3.org/2005/xpath-functions"
  xmlns:xs		= "http://www.w3.org/2001/XMLSchema"
  xmlns:w		= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:word200x	= "http://schemas.microsoft.com/office/word/2003/wordml"
  xmlns:v		= "urn:schemas-microsoft-com:vml" 
  xmlns:dbk		= "http://docbook.org/ns/docbook"
  xmlns:wx		= "http://schemas.microsoft.com/office/word/2003/auxHint"
  xmlns:o		= "urn:schemas-microsoft-com:office:office"
  xmlns:pkg		= "http://schemas.microsoft.com/office/2006/xmlPackage"
  xmlns:r		= "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel		= "http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:exsl		= 'http://exslt.org/common'
  xmlns:saxon		= "http://saxon.sf.net/"
  xmlns:letex		= "http://www.le-tex.de/namespace"
  xmlns:mml             = "http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
 xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
              xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
              xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
              xmlns:w10="urn:schemas-microsoft-com:office:word"
              xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"

  >
  
  <xsl:function name="docx2hub:srcpath" as="xs:string">
    <xsl:param name="elt" as="element(*)?"/>
    <xsl:sequence select="string-join(
                            (
                              if ($elt/.. instance of element(*)) then docx2hub:srcpath($elt/..) else concat(base-uri($elt), '?xpath='),
                              '/',
                              name($elt),
                              '[',
                              xs:string(index-of(for $s in $elt/../*[name() = name($elt)] return generate-id($s), generate-id($elt))),
                              ']'
                            ),
                            ''
                          )"/>
  </xsl:function>

  <xsl:template match="*[ self::w:p or self::w:t or self::w:tbl or self::w:tc or self::w:hyperlink or self::w:r ]
                        [ /*/name() = ('w:document', 'w:footnotes', 'w:endnotes', 'w:comments')]" mode="insert-xpath">
    <xsl:copy copy-namespaces="no">
      <xsl:if test="$srcpaths eq 'yes'">
        <xsl:attribute name="srcpath" select="docx2hub:srcpath(.)"/>
      </xsl:if>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="/w:document" mode="insert-xpath" as="document-node(element(w:root))" priority="2">
    <xsl:document>
      <w:root>
        <xsl:attribute name="xml:base" select="replace($base-dir, 'word/$', '')" />
        <xsl:apply-templates select="document(resolve-uri('styles.xml', base-uri()))/w:styles" mode="#current"/>
        <xsl:if test="doc-available(resolve-uri('numbering.xml', base-uri()))">
          <xsl:sequence select="document(resolve-uri('numbering.xml', base-uri()))/w:numbering" />
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('footnotes.xml', base-uri()))">
          <xsl:apply-templates select="document(resolve-uri('footnotes.xml', base-uri()))/w:footnotes" mode="#current" />
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('endnotes.xml', base-uri()))">
          <xsl:apply-templates select="document(resolve-uri('endnotes.xml', base-uri()))/w:endnotes" mode="#current" />
        </xsl:if>
        <xsl:sequence select="document(resolve-uri('settings.xml', base-uri()))/w:settings" />
        <xsl:if test="doc-available(resolve-uri('comments.xml', base-uri()))">
          <xsl:apply-templates select="document(resolve-uri('comments.xml', base-uri()))/w:comments" mode="#current" />
        </xsl:if>
        <xsl:sequence select="document(resolve-uri('fontTable.xml', base-uri()))/w:fonts" />
        <w:docRels>
          <xsl:sequence select="document(resolve-uri('_rels/document.xml.rels', base-uri()))/rel:Relationships" />
        </w:docRels>
        <xsl:if test="doc-available(resolve-uri('_rels/footnotes.xml.rels', base-uri()))">
          <w:footnoteRels>
            <xsl:sequence select="document(resolve-uri('_rels/footnotes.xml.rels', base-uri()))/rel:Relationships" />
          </w:footnoteRels>
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('_rels/comments.xml.rels', base-uri()))">
          <w:commentRels>
            <xsl:sequence select="document(resolve-uri('_rels/comments.xml.rels', base-uri()))/rel:Relationships" />
          </w:commentRels>
        </xsl:if>
        <!-- reproduce the document (with srcpaths), using the default identity template from catch-all.xsl: -->
        <xsl:next-match/>
      </w:root>
    </xsl:document>
  </xsl:template>


  <xsl:template match="/w:styles" mode="insert-xpath">
    <xsl:copy copy-namespaces="no">
      <!-- Font des Standardtextes -->
      <xsl:variable name="normal" select="key('docx2hub:style', 'Normal')" as="element(w:style)?" />
      <xsl:variable name="default-font" as="xs:string"
        select="if ($normal/w:rPr/w:rFonts/@w:ascii)
                then $normal/w:rPr/w:rFonts/@w:ascii
                else w:docDefaults/w:rPrDefault/w:rPr/w:rFonts/@w:ascii" />
      <!-- Font-size des Standardtextes -->
      <xsl:variable name="default-font-size" as="xs:string"
        select="if ($normal/w:rPr/w:sz/@w:val)
                then $normal/w:rPr/w:sz/@w:val
                else '20'" />
      <xsl:apply-templates select="@*, node() except w:latentStyles" mode="#current" >
        <xsl:with-param name="default-font" select="$default-font" tunnel="yes"/>
        <xsl:with-param name="default-font-size" select="$default-font-size" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="key('docx2hub:style', 'Normal')/w:rPr" mode="insert-xpath">
    <xsl:param name="default-font" as="xs:string?" tunnel="yes"/>
    <xsl:param name="default-font-size" as="xs:string" tunnel="yes"/>
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@* | node()" mode="#current"/>
      <xsl:if test="not(w:sz)">
        <w:sz w:val="{$default-font-size}"/>
      </xsl:if>
      <xsl:if test="not(w:rFonts) and $default-font">
        <w:rFonts w:ascii="{$default-font}"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>
  
</xsl:stylesheet>