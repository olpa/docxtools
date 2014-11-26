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
    <xsl:param name="text" as="element(*)*"/>
    <xsl:param name="nodes" as="element(*)*"/>

    <xsl:variable name="instr-from-nodes" as="node()*">
      <xsl:for-each select="$nodes//w:instrText">
        <xsl:choose>
          <xsl:when test="exists(parent::*/@css:*)">
            <phrase>
              <xsl:apply-templates select="parent::*/@css:*" mode="#current"/>
              <xsl:value-of select="replace(replace(replace(.,'\\&quot;','_quot_'),'&#8222;|&#8220;','&quot;'),'\\:', '#_-semi-_-colon-_#')"/>
            </phrase>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="replace(replace(replace(.,'\\&quot;','_quot_'),'&#8222;|&#8220;','&quot;'),'\\:', '#_-semi-_-colon-_#')"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="instr-from-nodes-text" as="xs:string?"
      select="string-join($instr-from-nodes, '')"/>
    <xsl:choose>
      <xsl:when test="matches($instr-from-nodes-text,'&quot;(\s+\\[bi])?\s*$')">
        <xsl:for-each-group select="$instr-from-nodes" group-starting-with="node()[matches(.,'^[\s&#160;]*[Xx][eE][\s&#160;]+')]">
          <xsl:variable name="current-instr-from-nodes-text" select="string-join(current-group(),'')"/>
          <xsl:if test="matches($current-instr-from-nodes-text, '\\[^bfrity]')">
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
              <xsl:when test="matches($current-instr-from-nodes-text, '\\t')">
                <xsl:value-of select="replace($current-instr-from-nodes-text, '^.*\\t\s*&quot;(.+?)&quot;(\\.*$|$)', '$1')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:sequence select="()"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="type" as="xs:string?">
            <xsl:choose>
              <xsl:when test="matches($current-instr-from-nodes-text, '\\f')">
                <xsl:value-of select="replace($current-instr-from-nodes-text, '^.*\\f\s*&quot;?(.+?)&quot;?\s*(\\.*$|$)', '$1')"/>
              </xsl:when>
              <xsl:when test="some $i in tokenize($current-instr-from-nodes-text,':') satisfies matches($i,'Register§§')">
                <xsl:value-of select="replace(tokenize($current-instr-from-nodes-text,':')[matches(.,'.*Register§§')],'.*Register§§(.*)$','$1')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:sequence select="()"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="indexterm-attributes" as="attribute()*">
            <xsl:if test="matches($current-instr-from-nodes-text, '\\i')">
              <xsl:attribute name="pagenum" select="'italic'"/>
            </xsl:if>
            <xsl:if test="matches($current-instr-from-nodes-text, '\\b')">
              <xsl:attribute name="pagenum" select="'bold'"/>
            </xsl:if>
            <xsl:if test="matches($current-instr-from-nodes-text, '\\s')">
              <xsl:attribute name="class" select="'startofrange'"/>
            </xsl:if>
            <xsl:if test="matches($current-instr-from-nodes-text, '\\e')">
              <xsl:attribute name="class" select="'endofrange'"/>
            </xsl:if>
            <xsl:if test="matches($current-instr-from-nodes-text, '\\r')">
              <xsl:attribute name="id" select="letex:rereplace-chars(replace($current-instr-from-nodes-text, '^.*\\r\s*&quot;?\s*(.+?)\s*&quot;?\s*(\\.*$|$)', '$1'))"/>
            </xsl:if>
            <xsl:if test="not(empty($type))">
              <xsl:attribute name="type" select="letex:rereplace-chars($type)"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name="temporary-term" as="node()*">
            <xsl:sequence select="letex:extract-chars(current-group(),'\','\\')"/>
          </xsl:variable>
          <xsl:variable name="real-term" as="node()*">
            <xsl:for-each-group select="$temporary-term" group-starting-with="*:text[matches(.,'^\\')]">
              <xsl:choose>
                <xsl:when test="current-group()[1][self::*:text[matches(.,'^\\')]]"/>
                <xsl:otherwise>
                  <xsl:for-each select="current-group()">
                    <xsl:choose>
                      <xsl:when test="self::*:text">
                        <xsl:apply-templates select=".//text()" mode="#current"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:apply-templates select="." mode="#current"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:for-each>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each-group>
          </xsl:variable>
          <xsl:variable name="real-term-text" select="string-join($real-term,'')"/>
          <xsl:variable name="indexterm">
            <indexterm>
              <xsl:apply-templates select="$indexterm-attributes" mode="#current"/>
              <primary sortas="{normalize-space(replace(letex:rm-last-quot(tokenize($real-term-text,':')[not(matches(.,'Register§§'))][1]),'^[\s&#160;]*[Xx][eE][\s&#160;]+&quot;',''))}">
                <xsl:apply-templates select="$real-term" mode="index-processing"/>
              </primary>
              <xsl:if test="count(tokenize($real-term-text,':')[not(matches(.,'Register§§'))]) gt 1">
                <secondary sortas="{normalize-space(replace(letex:rm-last-quot(tokenize($real-term-text,':')[not(matches(.,'Register§§'))][2]),'^[\s&#160;]*[Xx][eE][\s&#160;]+&quot;',''))}">
                  <xsl:apply-templates select="$real-term" mode="index-processing"/>
                </secondary>
              </xsl:if>
              <xsl:if test="count(tokenize($real-term-text,':')[not(matches(.,'Register§§'))]) gt 2">
                <tertiary sortas="{normalize-space(replace(letex:rm-last-quot(tokenize($real-term-text,':')[not(matches(.,'Register§§'))][3]),'^[\s&#160;]*[Xx][eE][\s&#160;]+&quot;',''))}">
                  <xsl:apply-templates select="$real-term" mode="index-processing"/>
                </tertiary>
              </xsl:if>
              <xsl:if test="not($see = '') and not(empty($see))">
                <see>
                  <xsl:value-of select="letex:rereplace-chars($see)"/>
                </see>
              </xsl:if>
            </indexterm>  
          </xsl:variable>
          <xsl:apply-templates select="$indexterm" mode="index-processing-1"/>
        </xsl:for-each-group>
      </xsl:when>
      <xsl:otherwise>
        <xsl:message select="$nodes[1]"></xsl:message>
        <xsl:call-template name="signal-error">
          <xsl:with-param name="error-code" select="'W2D_002'"/>
          <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
          <xsl:with-param name="hash">
            <value key="xpath"><xsl:value-of select="$nodes[1]/@srcpath"/></value>
            <value key="level">INT</value>
            <value key="info-text"><xsl:value-of select="$instr-from-nodes"/></value>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="text()[matches(., '^\s*XE\s*&quot;.*$')]" mode="index-processing" priority="10">
    <xsl:choose>
      <xsl:when test="matches(., '\s*XE\s*&quot;(.*)&quot;\s*$')">
        <xsl:value-of select="replace(., '^\s*XE\s*&quot;(.*)&quot;\s*$', '$1')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="replace(., '^\s*XE\s*&quot;(.*)$', '$1')"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="text()[matches(., '^\s*XE\s+$')]" mode="index-processing"/>

  <xsl:template match="text()[matches(.,'^\s*&quot;[^\s]+')]" mode="index-processing" priority="+1">
    <xsl:choose>
      <xsl:when test="matches(., '^\s*&quot;(.*)&quot;\s*$')">
        <xsl:value-of select="replace(., '^\s*&quot;(.*)&quot;\s*$', '$1')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="replace(., '^\s*&quot;(.*)$', '$1')"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="text()[matches(.,'&quot;\s*$')]" mode="index-processing">
    <xsl:value-of select="letex:rm-last-quot(.)"/>
  </xsl:template>
  
  <xsl:template match="*:primary" mode="index-processing-1">
    <xsl:variable name="content" as="node()*">
      <xsl:sequence select="letex:extract-chars(node(),':',':')"/>
    </xsl:variable>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="$content" group-starting-with="*:text[matches(.,'^:')]">
        <xsl:variable name="pos" select="position()"/>
        <xsl:choose>
          <xsl:when test="$pos=1 and not(exists($content[matches(.,'Register§§')]))">
            <xsl:value-of select="normalize-space(letex:rereplace-chars(replace(current-group()[1]/descendant-or-self::text(),'^:[\s&#160;]*','')))"/>
            <xsl:apply-templates select="current-group()[position() gt 1]" mode="index-processing-2"/>
          </xsl:when>
          <xsl:when test="$pos=2 and exists($content[matches(.,'Register§§')])">
            <xsl:value-of select="normalize-space(letex:rereplace-chars(replace(current-group()[1]/descendant-or-self::text(),'^:[\s&#160;]*','')))"/>
            <xsl:apply-templates select="current-group()[position() gt 1]" mode="index-processing-2"/>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:secondary" mode="index-processing-1">
    <xsl:variable name="content" as="node()*">
      <xsl:sequence select="letex:extract-chars(node(),':',':')"/>
    </xsl:variable>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="$content" group-starting-with="*:text[matches(.,'^:')]">
        <xsl:variable name="pos" select="position()"/>
        <xsl:choose>
          <xsl:when test="$pos=2 and not(exists($content[matches(.,'Register§§')]))">
            <xsl:value-of select="normalize-space(letex:rereplace-chars(replace(current-group()[1]//text(),'^:[\s&#160;]*','')))"/>
            <xsl:apply-templates select="current-group()[position() gt 1]" mode="index-processing-2"/>
          </xsl:when>
          <xsl:when test="$pos=3 and exists($content[matches(.,'Register§§')])">
            <xsl:value-of select="normalize-space(letex:rereplace-chars(replace(current-group()[1]//text(),'^:[\s&#160;]*','')))"/>
            <xsl:apply-templates select="current-group()[position() gt 1]" mode="index-processing-2"/>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:tertiary" mode="index-processing-1">
    <xsl:variable name="content" as="node()*">
      <xsl:sequence select="letex:extract-chars(node(),':',':')"/>
    </xsl:variable>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="$content" group-starting-with="*:text[matches(.,'^:')]">
        <xsl:variable name="pos" select="position()"/>
        <xsl:choose>
          <xsl:when test="$pos=3 and not(exists($content[matches(.,'Register§§')]))">
            <xsl:value-of select="normalize-space(letex:rereplace-chars(replace(current-group()[1]//text(),'^:[\s&#160;]*','')))"/>
            <xsl:apply-templates select="current-group()[position() gt 1]" mode="index-processing-2"/>
          </xsl:when>
          <xsl:when test="$pos=4 and exists($content[matches(.,'Register§§')])">
            <xsl:value-of select="normalize-space(letex:rereplace-chars(replace(current-group()[1]//text(),'^:[\s&#160;]*','')))"/>
            <xsl:apply-templates select="current-group()[position() gt 1]" mode="index-processing-2"/>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:text | text()" mode="index-processing-2">
    <xsl:value-of select="letex:rereplace-chars(.)"/>
  </xsl:template>
  
  <xsl:function name="letex:rm-last-quot">
    <xsl:param name="context"/>
    <xsl:value-of select="replace($context, '&quot;\s*$', '')"/>
  </xsl:function>
  
  <xsl:function name="letex:rereplace-chars">
    <xsl:param name="context"/>
    <xsl:value-of select="replace(replace($context,'_quot_','&quot;'),'#_-semi-_-colon-_#',':')"/>
  </xsl:function>
  
  <xsl:function name="letex:extract-chars">
    <xsl:param name="context" as="node()*"/>
    <xsl:param name="string-char" as="xs:string"/>
    <xsl:param name="regex-char" as="xs:string"/>
    
    <xsl:for-each select="$context">
      <xsl:choose>
        <xsl:when test="matches(.,$regex-char)">
          <xsl:choose>
            <xsl:when test="self::text()">
              <xsl:if test="not(tokenize(.,$regex-char)[1]='')">
                <text>
                  <xsl:value-of select="tokenize(.,$regex-char)[1]"/>
                </text>
              </xsl:if>
              <xsl:for-each select="tokenize(.,$regex-char)[position() gt 1]">
                <text>
                  <xsl:value-of select="concat($string-char,.)"/>
                </text>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="self::*">
              <xsl:variable name="element-name" select="local-name(.)"/>
              <xsl:variable name="attributes" select="./@*" as="attribute()*"/>
              <xsl:if test="not(tokenize(./text(),$regex-char)[1]='')">
                <xsl:element name="{$element-name}">
                  <xsl:copy-of select="$attributes"/>
                  <xsl:value-of select="tokenize(./text(),$regex-char)[1]"/>
                </xsl:element>
              </xsl:if>
              <xsl:for-each select="tokenize(.,$regex-char)[position() gt 1]">
                <text>
                  <xsl:value-of select="$string-char"/>
                </text>
                <xsl:element name="{$element-name}">
                  <xsl:copy-of select="$attributes"/>
                  <xsl:value-of select="."/>
                </xsl:element>
              </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
              <xsl:sequence select="."/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="."/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:function>
  
</xsl:stylesheet>