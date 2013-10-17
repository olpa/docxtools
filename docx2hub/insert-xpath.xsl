<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"
  xmlns:xsl		= "http://www.w3.org/1999/XSL/Transform"
  xmlns:xs		= "http://www.w3.org/2001/XMLSchema"
  xmlns:w		= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:dbk		= "http://docbook.org/ns/docbook"
  xmlns:rel		= "http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:letex		= "http://www.le-tex.de/namespace"
  xmlns:mml             = "http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  exclude-result-prefixes="xs docx2hub mml letex dbk"
              
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

  <xsl:template match="*[ self::w:p or self::w:t or self::w:tbl or self::w:tc or self::w:hyperlink or self::w:r or self::w:br ]
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
        <xsl:variable name="rels" as="document-node(element(rel:Relationships))"
          select="document(resolve-uri(concat($base-dir,'_rels/document.xml.rels')))"/>
        <!-- At the moment, we only need themes for default font resolution that takes place
             in the current mode. Therefore, we don’t include the theme documents below  
             /w:root yet. We rather pass them as a tunneled variable. -->
        <xsl:variable name="themes" as="document-node(element(a:theme))*"
          select="for $t in $rels/rel:Relationships/rel:Relationship[@Type eq 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme']/@Target
                  return document(resolve-uri($t, base-uri()))"/>
        <xsl:apply-templates select="document(resolve-uri('styles.xml', base-uri()))/w:styles" mode="#current">
          <xsl:with-param name="themes" select="$themes" tunnel="yes"/>
        </xsl:apply-templates>
        <xsl:if test="doc-available(resolve-uri('numbering.xml', base-uri()))">
          <xsl:apply-templates select="document(resolve-uri('numbering.xml', base-uri()))/w:numbering" mode="#current"/>
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('footnotes.xml', base-uri()))">
          <xsl:apply-templates select="document(resolve-uri('footnotes.xml', base-uri()))/w:footnotes" mode="#current" />
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('endnotes.xml', base-uri()))">
          <xsl:apply-templates select="document(resolve-uri('endnotes.xml', base-uri()))/w:endnotes" mode="#current" />
        </xsl:if>
        <xsl:apply-templates select="document(resolve-uri('settings.xml', base-uri()))/w:settings" mode="#current"/>
        <xsl:if test="doc-available(resolve-uri('comments.xml', base-uri()))">
          <xsl:apply-templates select="document(resolve-uri('comments.xml', base-uri()))/w:comments" mode="#current" />
        </xsl:if>
        <xsl:apply-templates select="document(resolve-uri('fontTable.xml', base-uri()))/w:fonts" mode="#current"/>
        <w:docTypes>
          <xsl:apply-templates select="document(resolve-uri('../%5BContent_Types%5D.xml', base-uri()))/ct:Types" mode="#current"/>
        </w:docTypes>
        <w:docRels>
          <xsl:apply-templates select="document(resolve-uri('_rels/document.xml.rels', base-uri()))/rel:Relationships" mode="#current"/>
          <xsl:for-each select="document(resolve-uri('_rels/document.xml.rels', base-uri()))/rel:Relationships/rel:Relationship[
                                  @Type = (
                                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
                                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'
                                  )]">
            <xsl:if test="doc-available(resolve-uri(concat(@Target, '.rels'), base-uri()))">
              <xsl:apply-templates select="document(resolve-uri(concat(@Target, '.rels'), base-uri()))/rel:Relationships" mode="#current"/>
            </xsl:if>
          </xsl:for-each>
        </w:docRels>
        <xsl:if test="doc-available(resolve-uri('_rels/footnotes.xml.rels', base-uri()))">
          <w:footnoteRels>
            <xsl:apply-templates select="document(resolve-uri('_rels/footnotes.xml.rels', base-uri()))/rel:Relationships" mode="#current"/>
          </w:footnoteRels>
        </xsl:if>
        <xsl:if test="doc-available(resolve-uri('_rels/comments.xml.rels', base-uri()))">
          <w:commentRels>
            <xsl:apply-templates select="document(resolve-uri('_rels/comments.xml.rels', base-uri()))/rel:Relationships" mode="#current"/>
          </w:commentRels>
        </xsl:if>
        <!-- reproduce the document (with srcpaths), using the default identity template from catch-all.xsl: -->
        <xsl:next-match/>
      </w:root>
    </xsl:document>
  </xsl:template>

  <xsl:template match="w:document | w:numbering | w:endnotes | w:footnotes | w:settings | w:fonts | rel:Relationships | w:comments | ct:Types" mode="insert-xpath">
    <xsl:copy copy-namespaces="no">
      <xsl:attribute name="xml:base" select="base-uri()" />
      <xsl:apply-templates select="@*, node()" mode="#current"/>      
    </xsl:copy>
  </xsl:template>

  <!-- theme support incomplete … -->
  <xsl:function name="letex:theme-font" as="xs:string">
    <xsl:param name="rFonts" as="element(w:rFonts)?"/>
    <xsl:param name="themes" as="document-node(element(a:theme))*"/>
    <xsl:choose>
      <xsl:when test="not($themes | $rFonts)">
        <xsl:sequence select="'Arial'"/>
      </xsl:when>
      <xsl:when test="$rFonts/@w:asciiTheme">
        <!-- minor font is for the bulk text (major is for the headings).
             Spec sez dat w:asciiTheme has precedence over w:ascii (I don’t find it now, and it wasn’t all clear there) -->
        <xsl:sequence select="($themes/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface)[1]"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$rFonts/@w:ascii"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="/w:styles" mode="insert-xpath">
    <xsl:param name="themes" as="document-node(element(a:theme))*" tunnel="yes"/>
    <xsl:copy copy-namespaces="no">
      <xsl:attribute name="xml:base" select="base-uri()" />
      <!-- Font des Standardtextes -->
      <xsl:variable name="normal" select="key('docx2hub:style', 'Normal')" as="element(w:style)?" />
      <xsl:variable name="default-font" as="xs:string"
        select="if ($normal/w:rPr/w:rFonts/@w:ascii)
                then $normal/w:rPr/w:rFonts/@w:ascii
                else letex:theme-font(w:docDefaults/w:rPrDefault/w:rPr/w:rFonts, $themes)" />
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
