<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"

  xmlns:xsl		= "http://www.w3.org/1999/XSL/Transform"
  xmlns:fn              = "http://www.w3.org/2005/xpath-functions"
  xmlns:xs		= "http://www.w3.org/2001/XMLSchema"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
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
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  xmlns="http://docbook.org/ns/docbook"

  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn letex w10 mml"
  >

  <!-- for the time being, encapsulate this markup in a <sidebar role="hub:foreign">.
    If there is content within this markup for which a Hub conversion exists, the usual
    conversion rules will apply. -->
  <xsl:template match="mc:AlternateContent" mode="wml-to-dbk">
    <xsl:variable name="elt-name" as="xs:string">
      <xsl:choose>
        <xsl:when test="parent::w:r | parent::w:p">
          <xsl:sequence select="'phrase'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="'sidebar'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:element name="{$elt-name}">
      <xsl:attribute name="role" select="'hub:foreign'"/>
      <xsl:copy-of select="parent::*/@srcpath"/>
      <xsl:apply-templates select="." mode="foreign"/>
    </xsl:element>
  </xsl:template>

  <xsl:template match="@* | * | w:drawing | w:txbxContent | w:pict" mode="foreign">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="w:*" mode="foreign">
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>

  <!--  OOXML spec 17.3.3.19 object (Embedded Object) -->
  <!--  This element specifies that an embedded object is located at this position in the run’s contents. The layout
        properties of this embedded object, as well as an optional static representation, are specified using the drawing
        element (§17.3.3.9). -->
  <!--  translation: objections can include all or nothing, VML, images, ActiveX, equations -->
  
  <xsl:template match="w:object" mode="wml-to-dbk">
    <xsl:apply-templates mode="vml">
      <xsl:with-param name="inline" select="true()" tunnel="yes"/>
    </xsl:apply-templates>
  </xsl:template>
  
  <!--  M.5.2 Shape Element
        The Shape element is the basic building block of VML. A shape can exist on its own or within a Group
        element. Shape defines many attributes and sub-elements that control the look and behavior of the
        shape. A shape must define at least a Path and size (Width, Height). VML also uses properties of the
        CSS2 style attribute to specify positioning and sizing. -->
  
  <xsl:template match="v:shape" mode="vml">
    <xsl:param name="inline" select="false()" tunnel="yes"/>
    <!--  VML also uses properties of the CSS2 style attribute to specify positioning and sizing. -->
    <xsl:apply-templates mode="vml">
      <xsl:with-param name="inline" select="$inline" tunnel="yes"/>
    </xsl:apply-templates>
  </xsl:template>

  <xsl:template match="w:pict" mode="wml-to-dbk">
    <xsl:apply-templates mode="vml">
      <xsl:with-param name="inline" select="false()" tunnel="yes"/>
    </xsl:apply-templates>
  </xsl:template>

  <xsl:template match="@srcpath" mode="vml"/>

  <xsl:template match="v:fill" mode="vml">
    <xsl:apply-templates select="@* except (@o:detectmouseclick)" mode="#current"/>
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="v:group" mode="vml">
    <xsl:apply-templates mode="#current"/>
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
    <inlinemediaobject role="OLEObject" annotations="{concat('object_',generate-id(parent::w:object))}">
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
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@srcpath"/></value>
        <value key="level">WRN</value>
        <value key="comment"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="v:oval" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_501'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@srcpath"/></value>
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
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">vml</value>
        <value key="info-text"><xsl:value-of select="concat('Element: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="@*[not(starts-with(name(), 'docx2hub:generated'))]" mode="vml" priority="-0.5">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_021'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="../@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">vml</value>
        <value key="info-text"><xsl:value-of select="concat('Attribut: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="processing-instruction()" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_023'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">vml</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="comment()" mode="vml">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_022'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">vml</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

</xsl:stylesheet>