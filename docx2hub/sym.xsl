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

  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn letex mml"
  >

  <xsl:template match="w:sym
                       |
                       w:t[string-length(.)=1 and ../w:rPr/w:rFonts/@w:ascii=('Symbol', 'Wingdings')]
                       |
                       w:lvlText[../w:rPr/w:rFonts/@w:ascii=('Symbol', 'Wingdings')]/@w:val
                       |
                       w:t[string-length(.)=1 and ../@css:font-family = ('Symbol', 'Wingdings')]
                       |
                       w:lvlText[../@css:font-family = ('Symbol', 'Wingdings')]/@w:val
                       " mode="wml-to-dbk" priority="1.5">
    <!-- priority = 1.5 because of priority = 1 ("default for attributes") in wml2dbk.xsl -->
    <xsl:variable name="font" select="if (self::w:sym) 
                                      then @w:font
                                      else
                                        if (self::attribute(w:val)) (: in w:lvlText :)
                                        then ../../w:rPr/w:rFonts/@w:ascii
                                        else (../w:rPr/w:rFonts/@w:ascii, ../@css:font-family)[1]" as="xs:string?"/>
    <xsl:if test="empty($font)">
      <xsl:call-template name="signal-error">
        <xsl:with-param name="error-code" select="'W2D_080'"/>
        <xsl:with-param name="exit" select="'yes'"/>
        <xsl:with-param name="hash">
          <value key="xpath"><xsl:value-of select="@xpath"/></value>
          <value key="level">INT</value>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
    <xsl:variable name="number" select="if (self::w:sym) then @w:char else xs:string(.)"/>
    <xsl:variable name="font-map-name" select="concat('fontmaps/', replace($font, ' ', '_'), '.xml')" as="xs:string" />
    <xsl:variable name="font_map" select="if (doc-available($font-map-name)) then document($font-map-name)/symbols else ()" as="element(symbols)?"/>
    <xsl:variable name="text" as="node()">
      <xsl:choose>
        <xsl:when test="if (self::w:sym) then $font_map/symbol[@number = $number] else $font_map/symbol[@entity = $number]">
          <xsl:choose>
            <xsl:when test="if (self::w:sym) then $font_map/symbol[@number = $number]/@char = '&#x000a;' else $font_map/symbol[@entity = $number]/@char = '&#x000a;'">
              <br/>
            </xsl:when>
            <xsl:otherwise>
              <text>
                <xsl:value-of select="if (self::w:sym) then $font_map/symbol[@number = $number]/@char else $font_map/symbol[@entity = $number]/@char"/>
              </text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="self::w:sym">
              <symbol css:font-family="{$font}">
                <xsl:value-of select="$number"/>
              </symbol>
            </xsl:when>
            <xsl:otherwise>
              <text>
                <xsl:value-of select="$number"/>
                <xsl:call-template name="signal-error">
                  <xsl:with-param name="error-code" select="'W2D_601'"/>
                  <xsl:with-param name="exit" select="'no'"/>
                  <xsl:with-param name="hash">
                    <value key="xpath"><xsl:value-of select="@xpath"/></value>
                    <value key="level">WRN</value>
                    <value key="info-text"><xsl:value-of select="$font"/> (0x<xsl:value-of select="letex:dec-to-hex(string-to-codepoints($number))"/>)</value>
                    <value key="pi">Zeichen <xsl:value-of select="string-to-codepoints($number)"/> aus Font <xsl:value-of select="$font"/> konnte nicht gemappt werden</value>
                  </xsl:with-param>
                </xsl:call-template>              
              </text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$text[self::text]">
        <xsl:value-of select="$text"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:message>
          <xsl:copy-of select="$text"/>
        </xsl:message>
        <xsl:call-template name="signal-error">
          <xsl:with-param name="error-code" select="'W2D_601'"/>
          <xsl:with-param name="exit" select="'no'"/>
          <xsl:with-param name="hash">
            <value key="xpath"><xsl:value-of select="@xpath"/></value>
            <value key="level">WRN</value>
            <value key="info-text"><xsl:value-of select="$font"/> (0x<xsl:value-of select="$text"/>)</value>
            <value key="pi">Zeichen <xsl:value-of select="string-to-codepoints($text)"/> aus Font <xsl:value-of select="$font"/> konnte nicht gemappt werden</value>
            <value key="comment"/>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:element name="{$text/name()}">
          <xsl:copy-of select="$text/@*"/>
          <xsl:value-of select="$text"/>
        </xsl:element>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="create-symbol">
    <xsl:param name="tokens" as="xs:string+"/>
    <xsl:param name="context" as="node()"/>
    <xsl:choose>
      <xsl:when test="$tokens[3] = '\f' and count($tokens) = 4">
        <xsl:variable name="font" select="replace($tokens[4], '&quot;', '')"/>
        <xsl:variable name="sym" select="$tokens[2]"/>
        <xsl:choose>
          <xsl:when test="$font = ('Symbol', 'Wingdings')">
            <xsl:variable name="number" select="if (matches($tokens[2], '^[0-9]+$')) then letex:dec-to-hex(xs:integer($tokens[2])) else 'NaN'"/>
            <xsl:choose>
              <xsl:when test="$number = 'NaN'">
                <xsl:call-template name="signal-error">
                  <xsl:with-param name="error-code" select="'W2D_601'"/>
                  <xsl:with-param name="exit" select="'no'"/>
                  <xsl:with-param name="hash">
                    <value key="xpath"><xsl:value-of select="$context/@xpath"/></value>
                    <value key="level">WRN</value>
                    <value key="info-text"><xsl:value-of select="concat($font, ': ', $tokens[2])"/></value>
                    <value key="pi">Zeichen <xsl:value-of select="string-to-codepoints($tokens[2])"/> aus Font <xsl:value-of select="$font"/> konnte nicht gemappt werden</value>
                    <value key="comment"/>
                  </xsl:with-param>
                </xsl:call-template>
                <symbol font-family="{$font}">
                  <xsl:value-of select="$tokens[2]"/>
                </symbol>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="key('symbol-by-number', upper-case($number), $symbol-font-map)/@char"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="signal-error">
              <xsl:with-param name="error-code" select="'W2D_601'"/>
              <xsl:with-param name="exit" select="'no'"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="$context/@xpath"/></value>
                <value key="level">WRN</value>
                <value key="info-text"><xsl:value-of select="concat($font, ': ', $tokens[2])"/></value>
                <value key="pi">Zeichen <xsl:value-of select="string-to-codepoints($tokens[2])"/> aus Font <xsl:value-of select="$font"/> konnte nicht gemappt werden</value>
                <value key="comment"/>
              </xsl:with-param>
            </xsl:call-template>
            <symbol font-family="{$font}">
              <xsl:value-of select="$tokens[2]"/>
            </symbol>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="signal-error">
          <xsl:with-param name="error-code" select="'W2D_602'"/>
          <xsl:with-param name="exit" select="'no'"/>
          <xsl:with-param name="hash">
            <value key="xpath"><xsl:value-of select="$context/@xpath"/></value>
            <value key="level">WRN</value>
            <value key="info-text"><xsl:value-of select="string-join($tokens, ' ')"/></value>
            <value key="comment"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="letex:dec-to-hex" as="xs:string">
    <xsl:param name="in" as="xs:integer?"/>
    <xsl:sequence select="if (not($in) or ($in eq 0)) 
                          then '0' 
                          else concat(
                            if ($in gt 15) 
                            then letex:dec-to-hex($in idiv 16) 
                            else '',
                            substring('0123456789abcdef', ($in mod 16) + 1, 1)
                          )"/>
  </xsl:function>

</xsl:stylesheet>