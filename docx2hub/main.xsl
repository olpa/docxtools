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
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
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
  <xsl:import href="modules/prop-mapping/map-props.xsl"/>
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
  

</xsl:stylesheet>
