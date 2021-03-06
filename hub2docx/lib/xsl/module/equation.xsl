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
    xmlns:o		= "urn:schemas-microsoft-com:office:office"
    xmlns:w		= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:m="http://www.w3.org/1998/Math/MathML"
    xmlns:omml		= "http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:wp		= "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:r		= "http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:mml="http://www.w3.org/1998/Math/MathML"

    xpath-default-namespace = "http://docbook.org/ns/docbook"

    exclude-result-prefixes = "xsl xs xsldoc saxon letex saxExtFn hub xlink o w m omml wp r mml"
>

  <xsl:import href="../omml/mml2omml.xsl"/>

<!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
<!-- mode="hub:default" -->
<!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <xsl:template  match="inlineequation"  mode="hub:default">
    <xsl:variable name="rPrContent" as="element(w:rStyle)">
      <w:rStyle w:val="InlineEquation" />
    </xsl:variable>
    <xsl:apply-templates select="node()[not(self::text()[matches(., '^\s*$')])]" mode="#current">
      <xsl:with-param name="rPrContent" select="$rPrContent" as="element(*)+" tunnel="yes" />
    </xsl:apply-templates>
  </xsl:template>

  <xsl:template  match="mathphrase"  mode="hub:default">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="m:math" mode="hub:default">
    <oMath xmlns="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <xsl:variable name="mml" as="node()*">
        <xsl:apply-templates select="." mode="m-to-mml"/>
      </xsl:variable>
      <xsl:apply-templates select="$mml" mode="mml" />
    </oMath>
  </xsl:template>
  
  <xsl:template match="m:*" mode="m-to-mml">
    <xsl:element name="mml:{local-name(.)}">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:element>
  </xsl:template>

  <xsl:template match="node()[not(self::m:*)] | @*" mode="m-to-mml">
    <xsl:copy-of select="."/>
  </xsl:template>

  <xsl:template  match="informalequation"  mode="hub:default">
    <w:p>
      <oMathPara xmlns="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <xsl:apply-templates mode="#current" />
      </oMathPara>
    </w:p>
  </xsl:template>

</xsl:stylesheet>
