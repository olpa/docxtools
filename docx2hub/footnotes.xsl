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

  <xsl:template match="w:footnoteReference" mode="wml-to-dbk">
    <footnote>
      <xsl:variable name="id" select="@w:id"/>
      <xsl:apply-templates select="/*/w:footnotes/w:footnote[@w:id = $id]" mode="#current"/>
    </footnote>
  </xsl:template>

  <xsl:template match="w:footnote" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:footnoteRef" mode="wml-to-dbk">
    <!-- setzt die Nummer der Fußnote. Prüfen!! -->
    <!-- GI 2013-05-23: Apparently both Word 2013 and LibreOffice 4.0.3 generate a number even if the 
      footnote doesn’t contain a footnoteRef. See for example DIN EN 419251-1, Sect. 6.1 -->
    <xsl:variable name="footnote-num-format" select="/w:root/w:settings/w:footnotePr/w:numFmt/@w:val" as="xs:string?"/>
    <xsl:variable name="footnote-number">
      <xsl:number value="(count(preceding::w:footnoteRef) + 1)" 
        format="{
        if ($footnote-num-format)
        then letex:get-numbering-format($footnote-num-format, '') 
        else '1'
        }"/>
    </xsl:variable>
    <phrase role="hub:identifier">
      <xsl:choose>
        <xsl:when test="//w:docVar[@w:name='footnote_check']">
          <xsl:choose>
            <xsl:when test="some $i in (tokenize(//w:docVar[@w:name='footnote_check']/@w:val,'&#xD;')) satisfies tokenize($i,',')[1]=$footnote-number">
              <xsl:value-of select="tokenize(tokenize(//w:docVar[@w:name='footnote_check']/@w:val,'&#xD;')[tokenize(.,',')[1]=$footnote-number],',')[2]"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$footnote-number"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$footnote-number"/>
        </xsl:otherwise>
      </xsl:choose>
    </phrase>
  </xsl:template>


</xsl:stylesheet>