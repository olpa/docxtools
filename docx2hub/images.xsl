<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:letex		= "http://www.le-tex.de/namespace"
  xmlns="http://docbook.org/ns/docbook"
  exclude-result-prefixes = "w v xs r rel letex"
  version="2.0">
  
  <xsl:template match="w:drawing[descendant::a:blip]" mode="wml-to-dbk">
    <mediaobject>
      <xsl:apply-templates select="@srcpath" mode="#current"/>
      <xsl:if test=".//wp:docPr/@descr ne ''">
        <alt>
          <xsl:value-of select=".//wp:docPr/@descr"/>
        </alt>
      </xsl:if>
      <xsl:apply-templates select="descendant::a:blip" mode="wml-to-dbk"/>
    </mediaobject>
    <!-- 
      I have it on good authority that Word stores file paths in {container}/word/document.xml.rels
      and NOT in wp:inline/wp:docPr/@descr. As the spelling indicates, @descr is used for 
      description. Word stores only the former os system path of the image in this attribute, if no 
      description is assigned. From the Office Open XML Spec:
      
      @descr: Specifies alternative text for the current DrawingML object, for use by assistive
              technologies or applications which do not display the current object.
              
      The code below is commented, as it derives image paths from @descr instead of document.xml.rels.  
    -->
    <!--<xsl:choose>
      <xsl:when test="exists(wp:inline/wp:docPr[@descr and matches(@descr,'\.([Jj][pP][gG]|[gG][iI][fF]|[pP][nN][gG]|[tT][iI][fF][fF]?)$')])">
        <mediaobject>
          <imageobject>
            <imagedata fileref="{wp:inline/wp:docPr/@descr}"/>
          </imageobject>
        </mediaobject>
      </xsl:when>
      <xsl:when test="exists(wp:inline/wp:docPr/@name)
                      and (: don't destroy ref to the embedded file name :)
                      not(descendant::a:blip[@r:embed])">
        <mediaobject>
          <imageobject>
            <imagedata fileref="{wp:inline/wp:docPr/@name}"/>
          </imageobject>
        </mediaobject>
      </xsl:when>
      <xsl:otherwise>
      <mediaobject>
        <xsl:apply-templates select="@srcpath" mode="#current"/>
        <xsl:if test=".//wp:docPr/@descr ne ''">
          <alt>
            <xsl:value-of select=".//wp:docPr/@descr"/>
          </alt>
        </xsl:if>
        <xsl:apply-templates select="descendant::a:blip" mode="wml-to-dbk"/>
        <xsl:if test=".//wp:docPr/@title ne ''">
          <caption>
            <para>
              <xsl:value-of select=".//wp:docPr/@title"/>
            </para>
          </caption>
        </xsl:if>
      </mediaobject>
    </xsl:otherwise>
   </xsl:choose>-->
  </xsl:template>
  
  <xsl:template match="w:drawing[not(descendant::a:blip)]" mode="wml-to-dbk">
    <phrase role="hub:foreign">
      <xsl:apply-templates select="." mode="foreign"/>
    </phrase>
  </xsl:template>
  
  <!-- images embedded in word zip container, usually stored in {docx}/word/media/ -->
  <xsl:template match="a:blip[@r:embed]" mode="wml-to-dbk">
    <xsl:call-template name="create-imageobject">
      <xsl:with-param name="image-id" select="@r:embed"/>
      <xsl:with-param name="embedded" select="true()"/>
    </xsl:call-template>
  </xsl:template>
  
  <!-- parent v:shape is processed in vml mode, see objects.xsl -->
  <xsl:template match="v:imagedata" mode="vml">
    <xsl:param name="inline" select="false()" tunnel="yes"/>
    <xsl:element name="{if ($inline) then 'inlinemediaobject' else 'mediaobject'}">
      <xsl:if test="ancestor::w:object">
        <xsl:attribute name="annotations" select="concat('object_',generate-id(ancestor::w:object[1]))"/>
      </xsl:if>
      <xsl:apply-templates select="@srcpath, parent::v:shape/@style" mode="#current"/>
      <xsl:call-template name="create-imageobject">
        <xsl:with-param name="image-id" select="@r:id"/>
        <xsl:with-param name="embedded" select="true()"/>
      </xsl:call-template>
    </xsl:element>
  </xsl:template>

  <xsl:template match="@style" mode="vml" priority="4">
    <xsl:analyze-string select="." regex="\s*;\s*">
      <xsl:non-matching-substring>
        <xsl:analyze-string select="." regex="(.+)\s*:\s*(.+)">
          <xsl:matching-substring>
            <xsl:attribute name="css:{regex-group(1)}" select="regex-group(2)"/>
          </xsl:matching-substring>
        </xsl:analyze-string>
      </xsl:non-matching-substring>
    </xsl:analyze-string>
  </xsl:template>

  <!-- externally referenced images -->
  <xsl:template match="a:blip[@r:link]" mode="wml-to-dbk">
    <xsl:call-template name="create-imageobject">
      <xsl:with-param name="image-id" select="@r:link"/>
      <xsl:with-param name="embedded" select="false()"/>
    </xsl:call-template>
  </xsl:template>
  
  <xsl:template name="create-imageobject">
    <xsl:param name="image-id" as="xs:string"/>
    <xsl:param name="embedded" as="xs:boolean"/>
    <!-- the file reference of an image object is stored in {docx}/_rels/document.xml.rels -->
    <xsl:variable name="relationships" 
      select="//w:docRels/rel:Relationships/rel:Relationship[@Type eq 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image']" 
      as="element(rel:Relationship)+"/>
    <xsl:variable name="file-uri" select="$relationships[@Id eq $image-id]/@Target" as="xs:string"/>
    <!-- include container prefix for files embedded in docx -->
    <xsl:variable name="patched-file-uri" select="
      if($embedded) 
      then concat('container:word/', $file-uri)
      else replace(replace($file-uri, '(file:/)//(.+)', '$1$2'),'\\', '/')" as="xs:string"/>
    <imageobject>
      <xsl:apply-templates select="../a:srcRect" mode="wml-to-dbk"/>
      <imagedata fileref="{$patched-file-uri}"/>
    </imageobject>
  </xsl:template>
  
  <!-- despite ISO 29100-1, which states that the clipping coordinates are always percentages, you’ll find
    large numbers as in <a:srcRect l="57262" t="34190" r="7344" b="42485"/> 
    These are called emu. 1 pt = 12700 emu, 1 mm = 36000 emu  -->
  <xsl:template match="a:srcRect[@l][@t][@r][@b][every $a in (@l, @t, @r, @b) satisfies (matches($a, '^\d+$'))]" mode="wml-to-dbk">
    <xsl:attribute name="css:clip" select="concat('rect(', string-join(for $a in (@t, @r, @b, @l) return  concat(round(number($a) * 0.0015748) * 0.05, 'pt'), ', '), ')')"/>
  </xsl:template>
  
</xsl:stylesheet>