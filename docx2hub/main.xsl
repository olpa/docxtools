<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"
  xmlns:xsl		= "http://www.w3.org/1999/XSL/Transform"
  xmlns:xs		= "http://www.w3.org/2001/XMLSchema"
  xmlns:w		= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:dbk		= "http://docbook.org/ns/docbook"
  xmlns:saxon		= "http://saxon.sf.net/"
  xmlns:letex		= "http://www.le-tex.de/namespace"
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  xmlns:mml             = "http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
  exclude-result-prefixes = "xs dbk saxon letex mml docx2hub"
>

  <!-- ================================================================================ -->
  <!-- IMPORT OF OTHER STYLESHEETS -->
  <!-- ================================================================================ -->

  <xsl:import href="wml2dbk.xsl"/>
  <xsl:import href="insert-xpath.xsl"/>
  <xsl:import href="map-props.xsl"/>
  <xsl:import href="join-runs.xsl"/>

  <!-- ================================================================================ -->
  <!-- OUTPUT FORMAT -->
  <!-- ================================================================================ -->

    <xsl:character-map  name="cleanups">
      <!-- Nobreak Hyphen zu Hyphen -->
      <xsl:output-character character="&#x2011;" string="-"/>
      <!-- Das ist der Dash der unordered lists aus dem Symbol Font -->
      <xsl:output-character character="&#xF0BE;" string="&#8212;"/>
      <!-- Istgleich -->
      <xsl:output-character character="&#xF03D;" string="="/>
      <!-- bullet -->
      <xsl:output-character character="&#xF8B7;" string="&#x2022;"/>
      <xsl:output-character character="&#xF8A7;" string="&#x2022;"/>
      <!-- psi -->
      <xsl:output-character character="&#xF079;" string="&#x03C8;"/>
      <!-- en-dash -->
      <xsl:output-character character="&#x23AF;" string="&#8211;"/>
      <!-- keine Ahnung was das ist -->
      <xsl:output-character character="&#xF02B;" string=""/>
    </xsl:character-map>
  
  <xsl:output
    method="xml"
    encoding="utf-8"
    indent="no"
    use-character-maps="cleanups"
    cdata-section-elements=''
    />

  <xsl:output
    name="debug"
    method="xml"
    encoding="utf-8"
    indent="yes"
    use-character-maps="cleanups"
    saxon:suppress-indentation="entry para title term link"
    cdata-section-elements=''
    />



  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PROCESS CHAIN ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  
  <xsl:variable name="insert-xpath">
    <xsl:apply-templates select="/" mode="insert-xpath"/>    
  </xsl:variable>

  <xsl:variable name="docx2hub:add-props" as="document-node(element(*))">
    <xsl:document>
      <xsl:apply-templates select="$insert-xpath" mode="docx2hub:add-props"/>
    </xsl:document>
  </xsl:variable>

  <!-- make proper attributes out of intermediate ones -->
  <xsl:variable name="docx2hub:props2atts" as="document-node(element(*))">
    <xsl:document>
      <xsl:apply-templates select="$docx2hub:add-props" mode="docx2hub:props2atts"/>
    </xsl:document>
  </xsl:variable>

  <xsl:variable name="docx2hub:remove-redundant-run-atts" as="document-node(element(*))">
    <xsl:document>
      <xsl:apply-templates select="$docx2hub:props2atts" mode="docx2hub:remove-redundant-run-atts"/>
    </xsl:document>
  </xsl:variable>

  <xsl:variable name="docx2hub:separate-field-functions" as="document-node(element(*))">
    <xsl:document>
      <xsl:apply-templates select="$docx2hub:remove-redundant-run-atts" mode="docx2hub:separate-field-functions"/>
    </xsl:document>
  </xsl:variable>

  <xsl:variable name="wml-to-dbk" as="document-node(element(*))">
    <xsl:document>
      <xsl:apply-templates  select="$docx2hub:separate-field-functions" mode="wml-to-dbk"/>
    </xsl:document>
  </xsl:variable>

  <xsl:variable name="docx2hub:join-runs" as="document-node(element(*))">
    <xsl:document>
      <xsl:apply-templates select="$wml-to-dbk" mode="docx2hub:join-runs"/>
    </xsl:document>
  </xsl:variable>


  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ main template ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  
  <xsl:template name="main">
    <xsl:if test="$debug = ('true', 'yes', '1')">
      <xsl:call-template name="debug" />
    </xsl:if>
    <xsl:sequence select="$docx2hub:join-runs"/>
  </xsl:template>

  <xsl:template name="debug">
    <xsl:result-document href="{letex:resolve-uri($debug-dir, '/01_insert-xpath.xml')}" format="debug">
      <xsl:sequence select="$insert-xpath"/>
    </xsl:result-document>
    <xsl:result-document href="{letex:resolve-uri($debug-dir, '/03_add-props.xml')}" format="debug">
      <xsl:sequence select="$docx2hub:add-props"/>
    </xsl:result-document>
    <xsl:result-document href="{letex:resolve-uri($debug-dir, '/04_props2atts.xml')}" format="debug">
      <xsl:sequence select="$docx2hub:props2atts"/>
    </xsl:result-document>
    <xsl:result-document href="{letex:resolve-uri($debug-dir, '/05_remove-redundant-run-atts.xml')}" format="debug">
      <xsl:sequence select="$docx2hub:remove-redundant-run-atts"/>
    </xsl:result-document>
    <xsl:result-document href="{letex:resolve-uri($debug-dir, '/11_separate-field-functions.xml')}" format="debug">
      <xsl:sequence select="$docx2hub:separate-field-functions"/>
    </xsl:result-document>
    <xsl:result-document href="{letex:resolve-uri($debug-dir, '/20_wml-to-dbk.xml')}" format="debug">
      <xsl:sequence select="$wml-to-dbk"/>
    </xsl:result-document>
    <xsl:result-document href="{letex:resolve-uri($debug-dir, '/24_join-runs.xml')}" format="debug">
      <xsl:sequence select="$docx2hub:join-runs"/>
    </xsl:result-document>
  </xsl:template>

</xsl:stylesheet>