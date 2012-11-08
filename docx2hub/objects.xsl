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
  xmlns:w10		= "urn:schemas-microsoft-com:office:word"
  xmlns:mml             = "http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"

  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn letex w10 mml"
  >


  <!-- Grundlegendes Problem: wie können displayed und inline voneinander unterschieden werden?
       Da Word nur malt, und hier keine explizite Unterscheidung vornimmt, geht es wohl nur mit Heuristik -->
  <xsl:template match="w:object" mode="wml-to-dbk">
    <inlineequation>
      <xsl:apply-templates mode="vml">
        <xsl:with-param name="inline" select="true()" tunnel="yes"/>
      </xsl:apply-templates>
    </inlineequation>
  </xsl:template>

  <xsl:template match="w:pict" mode="wml-to-dbk">
    <xsl:apply-templates mode="vml">
      <xsl:with-param name="inline" select="false()" tunnel="yes"/>
    </xsl:apply-templates>
  </xsl:template>

  <xsl:template match="@xpath" mode="vml"/>

  <xsl:template match="v:fill" mode="vml">
    <xsl:apply-templates select="@* except (@o:detectmouseclick)" mode="#current"/>
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="v:group" mode="vml">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="v:imagedata" mode="vml">
    <xsl:param name="inline" select="false()" as="xs:boolean" tunnel="yes"/>
    <xsl:variable name="key-name" as="xs:string"
      select="if (ancestor::w:footnote)
              then 'footnote-rel-by-id'
              else 'doc-rel-by-id'" />
    <xsl:variable name="styles" select="tokenize(parent::v:shape/@style,';')" as="xs:string*"/>
    <xsl:element name="{if ($inline) then 'inlinemediaobject' else 'mediaobject'}">
      <xsl:attribute name="role" select="'imagedata'"/>
      <imageobject>
        <!-- If these values are relevant, they have to be parsed and transformed to css: attributes.
        <xsl:attribute name="style" select="parent::v:shape/@style"/>
        -->
        <xsl:apply-templates select="@* except @r:id" mode="#current"/>
        <imagedata fileref="{key($key-name, current()/@r:id)/@Target}">
          <xsl:if test="$styles[matches(.,'^width')]">
            <xsl:attribute name="width" select="tokenize($styles[matches(.,'^width')],':')[2]"/>
          </xsl:if>
          <xsl:if test="$styles[matches(.,'^height')]">
            <xsl:attribute name="depth" select="tokenize($styles[matches(.,'^height')],':')[2]"/>
          </xsl:if>
        </imagedata>
      </imageobject>
      <xsl:variable name="img-file-name" select="key($key-name, current()/@r:id)/@Target" as="xs:string" />
      <xsl:if test="matches($img-file-name, '^media/image[0-9]+\.(jpe?g|png|tiff?)$')">
        <xsl:call-template name="signal-error">
          <xsl:with-param name="error-code" select="'W2D_502'"/>
          <xsl:with-param name="exit" select="'no'"/>
          <xsl:with-param name="hash">
            <value key="xpath"><xsl:value-of select="@xpath"/></value>
            <value key="level">WRN</value>
            <value key="comment"/>
            <value key="info-text"><xsl:value-of select="$img-file-name"/></value>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:element>
  </xsl:template>

  <xsl:template match="@o:title[parent::v:imagedata]" mode="vml">
  </xsl:template>

  <xsl:template match="@croptop[parent::v:imagedata]" mode="vml">
  </xsl:template>

  <xsl:template match="@cropleft[parent::v:imagedata]" mode="vml">
  </xsl:template>

  <xsl:template match="@cropright[parent::v:imagedata]" mode="vml">
  </xsl:template>

  <xsl:template match="@cropbottom[parent::v:imagedata]" mode="vml">
  </xsl:template>

  <xsl:template match="v:line" mode="vml">
  </xsl:template>

  <xsl:template match="v:rect" mode="vml">
  </xsl:template>

  <xsl:template match="o:lock" mode="vml">
  </xsl:template>

  <xsl:template match="o:OLEObject[parent::w:object]" mode="vml">
    <xsl:variable name="key-name" as="xs:string"
      select="if (ancestor::w:footnote)
              then 'footnote-rel-by-id'
              else 'doc-rel-by-id'" />
    <inlinemediaobject role="OLEObject">
      <imageobject>
        <xsl:apply-templates select="@* except @r:id" mode="#current"/>
        <imagedata fileref="{key($key-name, current()/@r:id)/@Target}"/>
      </imageobject>
    </inlinemediaobject>        
  </xsl:template>

  <xsl:template match="@ObjectID[parent::o:OLEObject]" mode="vml">
  </xsl:template>

  <xsl:template match="@DrawAspect[parent::o:OLEObject]" mode="vml">
  </xsl:template>

  <xsl:template match="@ProgID[parent::o:OLEObject]" mode="vml">
    <xsl:attribute name="role" select="."/>
  </xsl:template>

  <xsl:template match="@ShapeID[parent::o:OLEObject]" mode="vml">
  </xsl:template>

  <xsl:template match="@Type[parent::o:OLEObject]" mode="vml">
  </xsl:template>

  <xsl:template match="v:shape" mode="vml">
    <!-- process attributes at child -->
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="@css:*" mode="vml">
    <xsl:copy-of select="." />
  </xsl:template>

  <xsl:template match="@alt[parent::v:shape]" mode="vml">
    <!-- doch auswerten? -->
  </xsl:template>

  <xsl:template match="@coordsize[parent::v:shape]" mode="vml">
    <!--
         The physical size of a coordinate unit length is determined by both the size of the coordinate space (coordsize) and the size of the shape (style width and height). The coordsize attribute defines the number of horizontal and vertical subdivisions into which the shape's bounding box is divided. The combination of coordsize and style width/height effective scales the shape anisotropically.
         -->
  </xsl:template>

  <xsl:template match="@fillcolor[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@filled[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@insetpen[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@id[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@path[parent::v:shape]" mode="vml">
    <!-- 
         <xsl:message>Achtung. Strichzeichnung ignoriert (@path).</xsl:message>
         -->
  </xsl:template>

  <xsl:template match="@type[parent::v:fill]" mode="vml">
  </xsl:template>

  <xsl:template match="@color2[parent::v:fill]" mode="vml">
  </xsl:template>

  <xsl:template match="@type[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@stroked[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@strokeweight[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@style[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@o:allowoverlap[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@o:connectortype[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@o:ole[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@o:oleicon[parent::v:shape]" mode="vml">
  </xsl:template>
  
  <xsl:template match="@o:preferrelative[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="@o:spid[parent::v:shape]" mode="vml">
  </xsl:template>

  <xsl:template match="v:shapetype" mode="vml">
  </xsl:template>

  <xsl:template match="v:shadow" mode="vml">
  </xsl:template>

  <xsl:template match="o:callout" mode="vml">
  </xsl:template>

  <xsl:template match="v:textbox" mode="vml">
    <sidebar remap="v:textbox">
      <xsl:apply-templates select="parent::v:shape/@*" mode="#current"/>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates mode="#current"/>
    </sidebar>
  </xsl:template>

  <xsl:template match="@inset" mode="vml">
  </xsl:template>

  <xsl:template match="@style[parent::v:textbox]" mode="vml">
  </xsl:template>

  <xsl:template match="w10:wrap" mode="vml">
  </xsl:template>

  <xsl:template match="w10:anchorlock" mode="vml">
  </xsl:template>

  <xsl:template match="@o:bwmode" mode="vml">
  </xsl:template>

  <xsl:template match="@strokecolor" mode="vml">
  </xsl:template>

  <xsl:template match="@opacity" mode="vml">
  </xsl:template>

  <xsl:template match="@grayscale" mode="vml">
  </xsl:template>

  <xsl:template match="@blacklevel" mode="vml">
  </xsl:template>

  <xsl:template match="@gain" mode="vml">
  </xsl:template>

  <xsl:template match="@o:opacity2" mode="vml">
  </xsl:template>

  <xsl:template match="w:txbxContent" mode="vml">
    <!-- wechsel des Namespace -->
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>

  <xsl:template match="v:path" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_501'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@xpath"/></value>
        <value key="level">WRN</value>
        <value key="comment"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="v:oval" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_501'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@xpath"/></value>
        <value key="level">WRN</value>
        <value key="comment"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="v:stroke" mode="vml">
    <!--    <xsl:message>Achtung. Strichzeichnung ignoriert (stroke).</xsl:message> -->
  </xsl:template>


  <xsl:template match="*" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_020'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@xpath"/></value>
        <value key="level">INT</value>
        <value key="mode">vml</value>
        <value key="info-text"><xsl:value-of select="concat('Element: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="@*" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_021'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="../@xpath"/></value>
        <value key="level">INT</value>
        <value key="mode">vml</value>
        <value key="info-text"><xsl:value-of select="concat('Attribut: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="processing-instruction()" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_023'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@xpath"/></value>
        <value key="level">INT</value>
        <value key="mode">vml</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="comment()" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_022'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@xpath"/></value>
        <value key="level">INT</value>
        <value key="mode">vml</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

</xsl:stylesheet>