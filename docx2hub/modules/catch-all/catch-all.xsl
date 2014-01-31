<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  exclude-result-prefixes="docx2hub w">

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- CATCH-ALL (for all modes) -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  
  <xsl:template match="node()|@*" mode="#all" priority="-1">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*"     mode="#current"/>
      <xsl:apply-templates select="node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="*" mode="wml-to-dbk docx2hub:add-props" priority="-0.9">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*, *" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="w:t | *[@xml:space eq 'preserve']" mode="wml-to-dbk docx2hub:add-props" priority="-0.8">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  

</xsl:stylesheet>