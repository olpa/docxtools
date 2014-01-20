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
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"

  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn letex mml"
  >

  <xsl:template name="handle-index">
    <xsl:param name="instr" as="xs:string?"/>
    <xsl:param name="text" as="node()*"/>
    <xsl:param name="nodes" as="node()*"/>

    <!-- §§§ switch to xsl:analyze-string eventually -->
    
    <!-- replaced with variable below. transformation failed with german quotation marks -->
    <!--<xsl:variable name="instru" select="replace(replace(replace($instr,'\\&quot;','_quot_'),'\\([&#8222;&#8220;])','$1'),'\\:', '#_-semi-_-colon-_#')"/>-->
    <xsl:variable name="instru" select="replace(replace(replace($instr,'\\&quot;','_quot_'),'&#8222;|&#8220;','&quot;'),'\\:', '#_-semi-_-colon-_#')"/>
    <xsl:choose>
      <xsl:when test="matches($instru,'&quot;(\s+\\[bi])?\s*$')">
        <xsl:variable name="split-instru" select="tokenize($instru,'(^|&quot;)[\s&#160;]*[Xx][eE]($|[\s&#160;])')[string-length(.) gt 0]" as="xs:string*"/>
        <xsl:for-each select="$split-instru">
          <xsl:variable name="current-instru" select="if (matches(.,'&quot;(\s+\\[bi])?\s*$')) then concat('XE ',.) else concat('XE ',.,'&quot;')"/>
          <xsl:if test="matches($current-instru, '\\[^bfrity]')">
            <xsl:call-template name="signal-error">
              <xsl:with-param name="error-code" select="'W2D_001'"/>
              <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="$nodes[1]/@srcpath"/></value>
                <value key="level">INT</value>
                <value key="info-text"><xsl:value-of select="$instr"/></value>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>
          <xsl:variable name="see" as="xs:string?">
            <xsl:choose>
              <xsl:when test="matches($current-instru, '\\t')">
                <xsl:value-of select="replace($current-instru, '^.*\\t\s*&quot;(.+?)&quot;.*$', '$1')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:sequence select="()"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="type" as="xs:string?">
            <xsl:choose>
              <xsl:when test="matches($current-instru, '\\f')">
                <xsl:value-of select="replace($current-instru, '^.*\\f\s*&quot;(.+?)&quot;.*$', '$1')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:sequence select="()"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="term" select="tokenize(replace($current-instru, '^\s*[Xx][Ee]\s+&quot;(.+?)&quot;.*$', '$1'), ':')"/>
          <indexterm>
            <xsl:if test="not(empty($type))">
              <xsl:attribute name="type" select="$type"/>
            </xsl:if>
            <primary>
              <xsl:value-of select="replace(replace($term[1],'_quot_','&quot;'), '#_-semi-_-colon-_#', ':')"/>
            </primary>
            <xsl:if test="exists($term[2]) and not(matches($term[2], '^[\s&#xa0;]+$'))">
              <secondary>
                <xsl:value-of select="replace(replace($term[2],'_quot_','&quot;'), '#_-semi-_-colon-_#', ':')"/>
              </secondary>
            </xsl:if>
            <xsl:if test="exists($term[3]) and not(matches($term[3], '^[\s&#xa0;]+$'))">
              <tertiary>
                <xsl:value-of select="replace(replace($term[3],'_quot_','&quot;'), '#_-semi-_-colon-_#', ':')"/>
              </tertiary>
            </xsl:if>
            <xsl:if test="not($see = '') and not(empty($see))">
              <see>
                <xsl:value-of select="replace($see,'_quot_','&quot;')"/>
              </see>
            </xsl:if>
          </indexterm>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:message select="$nodes[1]"></xsl:message>
        <xsl:call-template name="signal-error">
          <xsl:with-param name="error-code" select="'W2D_002'"/>
          <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
          <xsl:with-param name="hash">
            <value key="xpath"><xsl:value-of select="$nodes[1]/@srcpath"/></value>
            <value key="level">INT</value>
            <value key="info-text"><xsl:value-of select="$instru"/></value>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>