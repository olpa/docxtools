<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"

  xmlns:xsl		= "http://www.w3.org/1999/XSL/Transform"
  xmlns:xs		= "http://www.w3.org/2001/XMLSchema"
  xmlns:w		= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:dbk		= "http://docbook.org/ns/docbook"
  xmlns:letex		= "http://www.le-tex.de/namespace"
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"

  exclude-result-prefixes = "w xs dbk letex docx2hub"
  >

  <xsl:function name="letex:insert-numbering" as="item()*">
    <xsl:param name="context" as="node()"/>
    <xsl:if test="not($context/name() eq 'w:p')">
      <xsl:call-template name="signal-error">
        <xsl:with-param name="error-code" select="'W2D_060'"/>
        <xsl:with-param name="exit" select="'yes'"/>
        <xsl:with-param name="hash">
          <value key="xpath"><xsl:value-of select="$context/@srcpath"/></value>
          <value key="level">INT</value>
          <value key="info-text"><xsl:value-of select="$context/name()"/></value>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
    <xsl:variable name="lvl" select="letex:get-lvl-of-numbering($context, false())" as="element(w:lvl)?"/>
    <xsl:choose>
      <xsl:when test="exists($lvl)">
        <xsl:if test="not($lvl/w:lvlText)">
          <xsl:call-template name="signal-error">
            <xsl:with-param name="error-code" select="'W2D_061'"/>
            <xsl:with-param name="exit" select="'yes'"/>
            <xsl:with-param name="hash">
              <value key="xpath">
                <xsl:value-of select="$lvl/@srcpath"/>
              </value>
              <value key="level">INT</value>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
        <!-- should we also propagate $lvl/w:pPr/dbk:tabs? Then we might end up with
          two tabs declarations on a paragraph. Text case: DIN EN 1865-2, tr document,
          the heading_1 elements.
          We are only inserting attributes that are not already present on the context
          paragraph. 
        -->
        <xsl:apply-templates mode="numbering"
          select="$lvl/w:pPr/@*[letex:different-style-att(., $context)]">
          <xsl:with-param name="context" select="$context" tunnel="yes"/>
        </xsl:apply-templates>
        <xsl:apply-templates select="$context/dbk:tabs" mode="#current"/>
        <xsl:element name="phrase">
          <xsl:attribute name="role" select="'hub:identifier'"/>
          <xsl:apply-templates mode="numbering"
            select="$lvl/w:rPr/@*[letex:different-style-att(., $context)], $lvl/w:lvlText" >
            <xsl:with-param name="context" select="$context" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:element>
        <tab/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="$context/dbk:tabs" mode="#current"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:key name="style-by-name" match="css:rule | dbk:style" use="@name | @role"/>

  <xsl:function name="letex:different-style-att" as="xs:boolean">
    <xsl:param name="att" as="attribute(*)"/>
    <xsl:param name="context" as="element(*)"/>
    <xsl:variable name="context-atts" select="key('style-by-name', $context/@role, $att/root())/@* | $context/@*" as="attribute(*)*"/>
    <xsl:sequence select="not(some $a in $context-atts satisfies ($a/name() = $att/name() and string($a) = string($att)))"/>
  </xsl:function>

  <xsl:template match="dbk:tabs" mode="numbering">
    <xsl:copy-of select="." copy-namespaces="no"/>
  </xsl:template>

  <xsl:function name="letex:get-lvl-of-numbering" as="element(w:lvl)?">
    <xsl:param name="context" as="node()"/>
    <xsl:param name="just-count" as="xs:boolean"/>
    <!-- for-each: just to avoid an XTDE1270 which shouldn't happen when the 3-arg form of key() is invoked: -->
    <xsl:for-each select="$context">
    <xsl:variable name="numPr" select="if ($just-count) 
                                       then if ($context/w:numPr[w:numId/@w:val]) 
                                            then $context/w:numPr 
                                            else ()
                                       else if ($context/w:numPr) 
                                            then $context/w:numPr 
                                            else ()"/>
    <xsl:variable name="style" select="key('docx2hub:style-by-role', @role, root($context))/w:numPr" as="element(w:numPr)?"/>
    <xsl:sequence select="if ($numPr)
                          then key(
                                 'abstract-numbering-by-id', 
                                 key('numbering-by-id', $numPr/w:numId/@w:val, root($context))/w:abstractNumId/@w:val,
                                 root($context)
                               )/w:lvl[
                                 @w:ilvl = $numPr/w:ilvl/@w:val
                               ]
                          else if ($style)
                               then if ($style/w:ilvl/@w:val) 
                                    then key(
                                          'abstract-numbering-by-id', 
                                          key('numbering-by-id', $style/w:numId/@w:val, root($context))/w:abstractNumId/@w:val,
                                          root($context)
                                         )/w:lvl[@w:ilvl = $style/w:ilvl/@w:val] 
                                    else if ($context/@role and exists(key(
                                                                        'abstract-numbering-by-id', 
                                                                        key('numbering-by-id', $style/w:numId/@w:val, root($context))/w:abstractNumId/@w:val,
                                                                        root($context)
                                                                       )/w:lvl[w:pStyle[@w:val = $context/@role]]))
                                         then key(
                                                'abstract-numbering-by-id', 
                                                key('numbering-by-id', $style/w:numId/@w:val, root($context))/w:abstractNumId/@w:val,
                                                root($context)
                                              )/w:lvl[w:pStyle[@w:val = $context/@role]]
                                         else key(
                                                'abstract-numbering-by-id', 
                                                key('numbering-by-id', $style/w:numId/@w:val, root($context))/w:abstractNumId/@w:val,
                                                root($context)
                                               )/w:lvl[@w:ilvl = '0']
                               else ()"/>
    </xsl:for-each>
  </xsl:function>

  <xsl:function name="letex:get-lvl-override" as="node()?">
    <xsl:param name="context" as="element(*)"/>
    <xsl:choose>
      <xsl:when test="empty($context)">
        <xsl:sequence select="()"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="numPr" select="if ($context/w:numPr) then $context/w:numPr else ()"/>
        <xsl:variable name="style" select="if ($context/w:pPr/w:pStyle) then key('style-by-id', $context/w:pPr/w:pStyle/@w:val, root($context))/w:numPr else ()"/>
        <xsl:sequence select="if ($numPr)
                              then key('numbering-by-id', $numPr/w:numId/@w:val, root($context))/w:lvlOverride[@w:ilvl = $numPr/w:ilvl/@w:val]
                              else if ($style)
                              then key('numbering-by-id', $style/w:numId/@w:val, root($context))/w:lvlOverride[@w:ilvl = $style/w:ilvl/@w:val]
                              else ()"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="*" mode="numbering">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="error-code" select="'W2D_020'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">numbering</value>
        <value key="info-text"><xsl:value-of select="concat('Element: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="w:start" mode="numbering">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:numFmt" mode="numbering">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:pStyle" mode="numbering">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:function name="letex:get-numbering-format" as="xs:string">
    <xsl:param name="format" as="xs:string"/>
    <xsl:param name="default" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="$format = 'lowerLetter'">a</xsl:when>
      <xsl:when test="$format = 'upperLetter'">A</xsl:when>
      <xsl:when test="$format = 'decimal'">1</xsl:when>
      <xsl:when test="$format = 'lowerRoman'">i</xsl:when>
      <xsl:when test="$format = 'upperRoman'">I</xsl:when>
      <xsl:when test="$format = 'bullet'"><xsl:value-of select="$default"/></xsl:when>
      <xsl:when test="$format = 'none'">none</xsl:when><!--GR-->
      <xsl:otherwise>
        <xsl:call-template name="signal-error">
          <xsl:with-param name="error-code" select="'W2D_062'"/>
          <xsl:with-param name="exit" select="'yes'"/>
          <xsl:with-param name="hash">
            <value key="level">INT</value>
            <value key="info-text"><xsl:value-of select="$format"/>INT</value>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="w:lvlText" mode="numbering">
    <xsl:param name="context" as="node()" tunnel="yes"/>
    <xsl:variable name="start" select="../w:start/@w:val"/>
    <!-- ST_NumberFormat, siehe Seite 1776 von Part 4 -->
    <xsl:variable name="fmt" select="letex:get-numbering-format(../w:numFmt/@w:val, @w:val)"/>
    <xsl:variable name="ilvl" select="../@w:ilvl"/>
    <xsl:variable name="lvl-id" select="generate-id(..)"/>
    <xsl:variable name="current" select="."/>
    <xsl:choose>

      <!-- gibt es Counter? -->
      <xsl:when test="matches(@w:val, '%')">

        <xsl:analyze-string select="@w:val" regex="%([0-9])">

          <!-- Die Counter auswerten -->
          <xsl:matching-substring>

            <xsl:choose>

              <!-- aktueller Counter entspricht der aktuellen Ebene -->
              <xsl:when test="xs:integer(regex-group(1)) - 1 = $ilvl">

                <xsl:variable name="last-step"
                  select="($context/preceding-sibling::w:p[
                          (: der abstractNum ist der selbe wie der von $current :)
                          letex:get-lvl-of-numbering(., true())/../generate-id() = $current/../../generate-id()
                          and
                          (: höhere Ebene oder Zähler zurückgesetzt :)
                          (letex:get-lvl-of-numbering(., true())/@w:ilvl &lt; $ilvl or letex:get-lvl-override(.))
                          and
                          (: listen check :)
                          (if ($current/../../w:name/@w:val = 'numbered list')
                              then (letex:get-lvl-of-numbering(., true())/../w:name/@w:val = 'numbered list'
                                    and letex:get-lvl-of-numbering(., true())/@w:ilvl = $ilvl)
                              else true())
                          ])[last()]"/>

                <xsl:variable name="stop-count" select="if ($last-step) then $last-step
                                                        else $context/preceding::*[last()]"/>

                <xsl:variable name="last-override"
                  select="$context/preceding-sibling::w:p[letex:get-lvl-override(.)
                          and matches(letex:get-lvl-of-numbering(., true())/w:lvlText/@w:val, '%')
                          and (not(letex:get-lvl-of-numbering(., true())/../generate-id() = $current/../../generate-id())
                          or not(w:numPr/w:numId/@w:val = $context/w:numPr/w:numId/@w:val
                              and w:numPr/w:ilvl/@w:val = $context/w:numPr/w:ilvl/@w:val))
                          and not(letex:get-lvl-of-numbering(., true())/@w:ilvl &gt; $ilvl)][1]"/>

                <xsl:variable name="last-override-stop"
                  select="if ($last-override) then $last-override else $context/preceding-sibling::*[last()]"/>

                <xsl:variable name="tmp-lvl" select="$current/../../w:lvl[@w:ilvl = ($ilvl+1)]"/>
                <xsl:variable name="current-lvl" select="$current/../../w:lvl[@w:ilvl = $ilvl]"/>

                <xsl:variable name="value"
                  select="if (letex:get-lvl-override($context) and $context/w:numPr) then
                          count($context/preceding-sibling::w:p[not(ancestor::w:tbl)][w:numPr/w:numId/@w:val = $context/w:numPr/w:numId/@w:val
                                                                and w:numPr/w:ilvl/@w:val = $context/w:numPr/w:ilvl/@w:val
                                                                and . &gt;&gt; $last-override-stop])
                          + $start
                          + (if
                          ($last-override-stop/preceding-sibling::w:p[w:numPr/w:numId/@w:val = $context/w:numPr/w:numId/@w:val
                                                                      and w:numPr/w:ilvl/@w:val = $context/w:numPr/w:ilvl/@w:val])
                          then 1 else 0)
                          + (if (letex:get-lvl-override($context)/w:startOverride/@w:val &gt; 1)
                             then (xs:integer(letex:get-lvl-override($context)/w:startOverride/@w:val) - 1) else 0)
                          else
                          count($context/preceding::w:p[not(ancestor::w:tbl)][letex:get-lvl-of-numbering(., true())/generate-id() = $lvl-id and (. &gt;&gt; $stop-count)])
                          + $start
                          + (if (letex:get-lvl-override($stop-count)
                                 and letex:get-lvl-override($stop-count)/w:startOverride/@w:val = 1) then 1 else 0)
                          (:KW 10.8.12 - wenn es tiefere ebenen gibt, die ebenen ueberspringen, den zaehler der aktuellen ebene korrigieren:)
                          + (if (exists($context/preceding-sibling::w:p[letex:get-lvl-of-numbering(., true())/generate-id() = $tmp-lvl/generate-id() and . &gt;&gt; $stop-count][count(./preceding-sibling::w:p[letex:get-lvl-of-numbering(., true())/generate-id() = $current-lvl/generate-id() and . &gt;&gt; $stop-count]) = 0])) then 1 else 0)
                          "/>
                <xsl:number value="$value" format="{$fmt}"/>
              </xsl:when>

              <!-- eine andere Ebene wird gezählt -->
              <xsl:otherwise>
                <xsl:variable name="tmp-lvl" select="$current/../../w:lvl[@w:ilvl = (xs:integer(regex-group(1)) - 1)]"/>
                <xsl:variable name="current-lvl" select="$current/../../w:lvl[@w:ilvl = xs:integer(regex-group(1))]"/>
                <xsl:variable name="last-step"
                  select="$context/preceding-sibling::w:p[
                          letex:get-lvl-of-numbering(., true())/../generate-id() = $current/../../generate-id()
                          and
                          letex:get-lvl-of-numbering(., true())/@w:ilvl &lt; (xs:integer(regex-group(1)) - 1)][1]"/>
                <xsl:variable name="stop-count" select="if ($last-step) then $last-step else $context/preceding-sibling::*[last()]"/>
                <xsl:variable name="count-preced" select="count($context/preceding-sibling::w:p[letex:get-lvl-of-numbering(., true())/generate-id() = $tmp-lvl/generate-id() and . &gt;&gt; $stop-count])"/>
                
                <xsl:variable name="value"
                  select="$tmp-lvl/w:start/@w:val - 1
                          (:KW 10.8.12 - wenn es keinen vorgaengerknoten in der gezaehlten ebene gibt, darf die numerierung
                          trotzdem nicht null werden; ebenfalls muessen darauf folgende knoten auch inkrementiert werden:)
                          + (if (($count-preced = 0) or (exists($context/preceding-sibling::w:p[letex:get-lvl-of-numbering(., true())/generate-id() = $current-lvl/generate-id() and . &gt;&gt; $stop-count][count(./preceding-sibling::w:p[letex:get-lvl-of-numbering(., true())/generate-id() = $tmp-lvl/generate-id() and . &gt;&gt; $stop-count]) = 0]))) 
                          then $count-preced+1 
                          else ($count-preced + (if ($stop-count/generate-id() = $context/preceding-sibling::w:p[letex:get-lvl-of-numbering(., true())/generate-id() = $tmp-lvl/generate-id()][last()]/generate-id())
                          then 1 else 0)))"/>
                <xsl:number value="$value" format="{letex:get-numbering-format($tmp-lvl/w:numFmt/@w:val, 'this should never happen')}"/>
              </xsl:otherwise>

            </xsl:choose>
          </xsl:matching-substring>

          <!-- Text zum Counter ausgeben -->
          <xsl:non-matching-substring>
            <xsl:value-of select="."/>
          </xsl:non-matching-substring>

        </xsl:analyze-string>
      </xsl:when>
      
      <!-- keine Counter, Text ausgeben -->
      <xsl:otherwise>
        <xsl:variable name="resolve-symbol-encoding" as="node()?">
          <!-- If ../w:rPr/w:rFonts/@w:ascii='Symbol', it will be transformed according to sym.xsl,
               otherwise by default template -->
          <xsl:apply-templates select="@w:val" mode="wml-to-dbk"/>
        </xsl:variable>
        <!-- GR: Text mit lvlOverride?  -->
        <xsl:value-of select="if (letex:get-lvl-override($context)/w:lvl[@w:ilvl = $context/w:numPr/w:ilvl/@w:val]/w:lvlText)
                              then letex:get-lvl-override($context)/w:lvl[@w:ilvl = $context/w:numPr/w:ilvl/@w:val]/w:lvlText/@w:val
                              else $resolve-symbol-encoding"/>
      </xsl:otherwise>

    </xsl:choose>
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:lvlJc" mode="numbering">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:suff" mode="numbering">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:lvlRestart" mode="numbering">
    <!-- check if this element must be checked within w:lvlText -->
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:pPr" mode="numbering">
  </xsl:template>

  <xsl:template match="w:rPr" mode="numbering">
  </xsl:template>

</xsl:stylesheet>