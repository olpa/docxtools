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

  <!-- This mode is called from docx2hub:remove-redundant-run-atts as a collateral -->

  <xsl:template match="w:numId" mode="docx2hub:abstractNum">
    <xsl:param name="ilvl" as="xs:integer"/>
    <xsl:variable name="lvl" as="element(w:lvl)?" select="key(
                                                           'abstract-numbering-by-id', 
                                                           key(
                                                             'numbering-by-id', 
                                                             @w:val 
                                                           )/w:abstractNumId/@w:val
                                                         )/w:lvl[@w:ilvl = $ilvl]"/>
    <xsl:variable name="lvlOverride" as="element(w:lvlOverride)?"
      select="key(
                'numbering-by-id', 
                @w:val 
              )/w:lvlOverride[@w:ilvl = $ilvl]"/>
    <xsl:apply-templates select="$lvl" mode="#current">
      <xsl:with-param name="numId" select="@w:val"/>
      <xsl:with-param name="start-override" select="for $so in $lvlOverride/w:startOverride/@w:val
                                                    return xs:integer($so)"/>
    </xsl:apply-templates>    
  </xsl:template>

  <!-- It seems that numId = '0' is for lists without marker? -->
  <xsl:template match="w:numId[@w:val = '0']" mode="docx2hub:abstractNum"/>

  <xsl:template match="w:lvl" mode="docx2hub:abstractNum">
    <xsl:param name="numId" as="xs:string"/>
    <xsl:param name="start-override" as="xs:integer?"/>
    <xsl:variable name="restart" as="xs:integer?" select="(for $r in w:lvlRestart/@w:val 
                                                           return xs:integer($r),
                                                           0[$start-override])[last()]"/>
    <xsl:attribute name="docx2hub:num-signature" select="string-join((../@w:abstractNumId, @w:ilvl), '_')"/>
    <xsl:attribute name="docx2hub:num-abstract" select="../@w:abstractNumId"/>
    <xsl:attribute name="docx2hub:num-ilvl" select="@w:ilvl"/>
    <xsl:attribute name="docx2hub:num-id" select="$numId"/>
    <xsl:if test="exists($restart)">
      <xsl:attribute name="docx2hub:num-restart-level" select="$restart"/>
    </xsl:if>
    <xsl:attribute name="docx2hub:num-restart-val" 
        select="($start-override, for $s in w:start/@w:val return xs:integer($s), 1)[1]"/>
  </xsl:template>
  
  <!-- collateral (only the first in a row should trigger a reset) -->
  <xsl:template match="@docx2hub:num-signature[../preceding-sibling::*[1]/@docx2hub:num-signature = current()]"
    mode="docx2hub:join-instrText-runs">
    <xsl:attribute name="docx2hub:num-continue" select="."/>
  </xsl:template>

  <xsl:template match="@docx2hub:num-signature[exists(../@docx2hub:num-restart-level)] (: should check whether 0 or a higher number :)
                                              [
                                                .. is (
                                                       key('docx2hub:num-signature', current())
                                                         [@docx2hub:num-id = current()/../@docx2hub:num-id]
                                                     )[1]
                                              ]" mode="docx2hub:join-instrText-runs" priority="2">
    <!-- the first of a numId that defines a start value override for this ilvl -->
    <xsl:copy/>
  </xsl:template>

  <xsl:template match="@docx2hub:num-signature" mode="docx2hub:join-instrText-runs">
    <xsl:variable name="last-same-signature" as="element(w:p)?" 
      select="(key('docx2hub:num-signature', current())[. &lt;&lt; current()/..])[last()]"/>
    <xsl:variable name="in-between" as="element(w:p)*"
      select="//w:p[. &gt;&gt; $last-same-signature][. &lt;&lt; current()/..]"/>
    <xsl:choose>
      <xsl:when test="empty ($last-same-signature)
                      or 
                      (
                        some $p in $in-between satisfies 
                        $p/@docx2hub:ilvl &lt; current()/../@docx2hub:ilvl
                      )">
        <xsl:copy/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="docx2hub:num-continue" select="."/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <xsl:function name="letex:insert-numbering" as="item()*">
    <xsl:param name="context" as="element(w:p)"/>
    
    <xsl:variable name="lvl" select="letex:get-lvl-of-numbering($context)" as="element(w:lvl)?"/>
    <xsl:choose>
      <xsl:when test="exists($lvl)">
        <xsl:if test="not($lvl/w:lvlText)">
          <xsl:call-template name="signal-error">
            <xsl:with-param name="error-code" select="'W2D_061'"/>
            <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
            <xsl:with-param name="hash">
              <value key="xpath">
                <xsl:value-of select="$lvl/@srcpath"/>
              </value>
              <value key="level">INT</value>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
        <xsl:variable name="context-atts" select="key('style-by-name', $context/@role, $context/root())/@* | $context/@*" as="attribute(*)*"/>
        <xsl:variable name="pPr" as="attribute(*)*">
          <xsl:apply-templates mode="numbering" select="$lvl/w:pPr/@*">
            <xsl:with-param name="context" select="$context" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:variable>
        <xsl:variable name="rPr" as="attribute(*)*">
          <xsl:apply-templates mode="numbering" select="$lvl/w:rPr/@*">
            <xsl:with-param name="context" select="$context" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:variable>
        <xsl:sequence select="$pPr, $context-atts[name() = $pPr/name()]"/>
        <xsl:apply-templates select="$context/dbk:tabs" mode="wml-to-dbk"/>
        <phrase role="hub:identifier">
          <xsl:sequence select="$rPr, $context-atts[name() = $rPr/name()]"/>
          <xsl:value-of select="letex:get-identifier($context,$lvl)"/>
        </phrase>
        <tab/>
      </xsl:when>
      <xsl:otherwise>
        <!--KW 11.6.13: mode hart reingeschrieben wegen null pointer exception-->
        <xsl:apply-templates select="$context/dbk:tabs" mode="wml-to-dbk"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="letex:get-lvl-of-numbering" as="element(w:lvl)?">
    <xsl:param name="context" as="node()"/>
    <!-- for-each: just to avoid an XTDE1270 which shouldn't happen when the 3-arg form of key() is invoked: -->
    <xsl:variable name="lvls" as="element(w:lvl)*">
      <xsl:for-each select="$context">
      <xsl:variable name="numPr" select="if ($context/w:numPr) 
                                         then $context/w:numPr 
                                         else ()"/>
      <xsl:variable name="style" select="key('docx2hub:style-by-role', @role, root($context))[last()]/w:numPr" as="element(w:numPr)?"/>
      <xsl:sequence select="if ($numPr)
                            then if (exists(
                                       key(
                                         'abstract-numbering-by-id', 
                                         key(
                                           'numbering-by-id', 
                                           $numPr/w:numId/@w:val, 
                                           root($context)
                                         )/w:abstractNumId/@w:val,
                                         root($context)
                                       )/w:lvl[@w:ilvl = $numPr/w:ilvl/@w:val]
                                     )) 
                                 then key(
                                        'abstract-numbering-by-id', 
                                        key(
                                          'numbering-by-id', 
                                          $numPr/w:numId/@w:val, 
                                          root($context)
                                        )/w:abstractNumId/@w:val,
                                        root($context)
                                      )/w:lvl[@w:ilvl = $numPr/w:ilvl/@w:val]
                                 else root($context)//w:abstractNum[
                                          w:styleLink/@w:val = key(
                                                                 'abstract-numbering-by-id', 
                                                                 key(
                                                                   'numbering-by-id', 
                                                                   $numPr/w:numId/@w:val, 
                                                                   root($context)
                                                                 )/w:abstractNumId/@w:val,
                                                                 root($context)
                                                               )/w:numStyleLink/@w:val
                                        ]/w:lvl[@w:ilvl = $numPr/w:ilvl/@w:val]
                            else if ($style)
                                 then if ($style/w:ilvl/@w:val) 
                                      then if (exists(
                                                 key(
                                                   'abstract-numbering-by-id', 
                                                   key(
                                                     'numbering-by-id', 
                                                     $style/w:numId/@w:val, 
                                                     root($context)
                                                   )/w:abstractNumId/@w:val,
                                                   root($context)
                                                 )/w:lvl[@w:ilvl = $style/w:ilvl/@w:val])
                                               )
                                           then key(
                                                  'abstract-numbering-by-id', 
                                                  key(
                                                    'numbering-by-id', 
                                                    $style/w:numId/@w:val, 
                                                    root($context)
                                                  )/w:abstractNumId/@w:val,
                                                  root($context)
                                                )/w:lvl[@w:ilvl = $style/w:ilvl/@w:val]
                                           else root($context)//w:abstractNum[
                                                    w:styleLink/@w:val = key(
                                                                           'abstract-numbering-by-id', 
                                                                           key(
                                                                             'numbering-by-id', 
                                                                             $style/w:numId/@w:val, 
                                                                             root($context)
                                                                           )/w:abstractNumId/@w:val,
                                                                           root($context)
                                                                         )/w:numStyleLink/@w:val
                                                  ]/w:lvl[@w:ilvl = $style/w:ilvl/@w:val]
                                      else if ($context/@role and exists(key(
                                                                           'abstract-numbering-by-id', 
                                                                           key(
                                                                             'numbering-by-id', 
                                                                             $style/w:numId/@w:val, 
                                                                             root($context)
                                                                           )/w:abstractNumId/@w:val,
                                                                           root($context)
                                                                         )/w:lvl[w:pStyle[@w:val = $context/@role]]))
                                           then key(
                                                  'abstract-numbering-by-id', 
                                                  key(
                                                    'numbering-by-id', 
                                                    $style/w:numId/@w:val, 
                                                    root($context)
                                                  )/w:abstractNumId/@w:val,
                                                  root($context)
                                                )/w:lvl[w:pStyle[@w:val = $context/@role]]
                                                else if ($context/@role and exists(root($context)//w:abstractNum[
                                                                                      w:styleLink/@w:val = key(
                                                                                                             'abstract-numbering-by-id', 
                                                                                                             key(
                                                                                                               'numbering-by-id', 
                                                                                                               $style/w:numId/@w:val, 
                                                                                                               root($context)
                                                                                                             )/w:abstractNumId/@w:val,
                                                                                                             root($context)
                                                                                                           )/w:numStyleLink/@w:val
                                                                                     ]/w:lvl[w:pStyle[@w:val = $context/@role]]))
                                                     then root($context)//w:abstractNum[
                                                              w:styleLink/@w:val = key(
                                                                                     'abstract-numbering-by-id', 
                                                                                     key(
                                                                                       'numbering-by-id', 
                                                                                       $style/w:numId/@w:val, 
                                                                                       root($context)
                                                                                     )/w:abstractNumId/@w:val,
                                                                                     root($context)
                                                                                   )/w:numStyleLink/@w:val
                                                            ]/w:lvl[w:pStyle[@w:val = $context/@role]]
                                                     else key(
                                                            'abstract-numbering-by-id', 
                                                            key(
                                                              'numbering-by-id', 
                                                              $style/w:numId/@w:val, 
                                                              root($context)
                                                            )/w:abstractNumId/@w:val,
                                                            root($context)
                                                          )/w:lvl[@w:ilvl = '0']
                                 else ()"/>
    </xsl:for-each>  
    </xsl:variable>
    <xsl:sequence select="$lvls[last()]"/>
    <!--    Only last lvl chosen, because of errors. Check for multiple lvls has to be implemented   -->
  </xsl:function>
  
  <xsl:function name="letex:get-lvl-override" as="element(*)?">
    <xsl:param name="context" as="element(w:p)?"/>
    <xsl:choose>
      <xsl:when test="empty($context)">
        <xsl:sequence select="()"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="numPr" select="if ($context/w:numPr) 
                                           then $context/w:numPr 
                                           else ()"/>
        <xsl:variable name="style" select="if ($context/w:pPr/w:pStyle) 
                                           then key('style-by-id', $context/w:pPr/w:pStyle/@w:val, root($context))/w:numPr 
                                           else ()"/>
        <xsl:sequence select="if ($numPr)
                              then key('numbering-by-id', $numPr/w:numId/@w:val, root($context))/w:lvlOverride[@w:ilvl = $numPr/w:ilvl/@w:val]
                              else if ($style)
                                   then key('numbering-by-id', $style/w:numId/@w:val, root($context))/w:lvlOverride[@w:ilvl = $style/w:ilvl/@w:val]
                                   else ()"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:key name="docx2hub:num-signature" match="*[@docx2hub:num-signature]" use="@docx2hub:num-signature"/>

  <xsl:function name="letex:get-identifier" as="xs:string">
    <xsl:param name="context" as="element(w:p)"/>
    <xsl:param name="lvl" as="element(w:lvl)"/>
    
    <xsl:variable name="abstract-num-id" select="xs:double($lvl/ancestor::w:abstractNum/@w:abstractNumId)" as="xs:double"/>
    <xsl:variable name="lvl-to-use" select="if (exists(letex:get-lvl-override($context)/w:lvl)) 
                                            then letex:get-lvl-override($context)/w:lvl 
                                            else $lvl"/>
    <xsl:variable name="ilvl" select="xs:double($lvl-to-use/@w:ilvl)"/>

    <xsl:variable name="start-of-relevant" as="element(w:p)?"
      select="if ($context/@docx2hub:num-signature)
              then $context
              else
                (
                  key(
                    'docx2hub:num-signature', 
                    ($context/@docx2hub:num-signature, $context/@docx2hub:num-continue), 
                    root($context)
                  )[. &lt;&lt; $context]
                )[last()]"/>
    
    <xsl:variable name="level-counter" as="xs:integer" 
      select="(for $s in $start-of-relevant/@docx2hub:num-restart-val return xs:integer($s), 1)[1] 
              + count(root($context)//w:p[. &gt;&gt; $start-of-relevant][. &lt;&lt; $context]
                                         [@docx2hub:num-continue = $start-of-relevant/@docx2hub:num-signature])
              + count($context[not(. is $start-of-relevant)])"/>
    
    <xsl:variable name="resolve-symbol-encoding">
      <element>
        <xsl:apply-templates select="$lvl-to-use/w:lvlText/@w:val" mode="wml-to-dbk"/>
      </element>
    </xsl:variable>
    <xsl:variable name="string" as="xs:string*">
      <xsl:choose>
        <xsl:when test="$resolve-symbol-encoding//@w:val">
          <xsl:analyze-string select="$lvl-to-use/w:lvlText/@w:val" regex="%([0-9])">
            <xsl:matching-substring>
              <xsl:number value="if (xs:double(regex-group(1)) gt $ilvl) 
                                 then $level-counter 
                                 else $level-counter - 1"
                          format="{letex:get-numbering-format($lvl/ancestor::w:abstractNum/w:lvl[@w:ilvl=xs:double(regex-group(1))-1]/w:numFmt/@w:val, $lvl-to-use/w:lvlText/@w:val)}"/>
            </xsl:matching-substring>
            <xsl:non-matching-substring>
              <xsl:value-of select="."/>
            </xsl:non-matching-substring>
          </xsl:analyze-string>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="$resolve-symbol-encoding//text()"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:sequence select="string-join($string,'')"/>
  </xsl:function>
  
  <xsl:function name="letex:get-numbering-format" as="xs:string">
    <xsl:param name="format" as="xs:string"/>
    <xsl:param name="default" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="$format = 'lowerLetter'">a</xsl:when>
      <xsl:when test="$format = 'upperLetter'">A</xsl:when>
      <xsl:when test="$format = 'decimal'">1</xsl:when>
      <xsl:when test="$format = 'lowerRoman'">i</xsl:when>
      <xsl:when test="$format = 'upperRoman'">I</xsl:when>
      <xsl:when test="$format = 'bullet'">
        <xsl:value-of select="$default"/>
      </xsl:when>
      <xsl:when test="$format = 'none'">none</xsl:when><!--GR-->
      <xsl:otherwise>
        <xsl:call-template name="signal-error">
          <xsl:with-param name="error-code" select="'W2D_062'"/>
          <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
          <xsl:with-param name="hash">
            <value key="level">INT</value>
            <value key="info-text"><xsl:value-of select="$format"/>INT</value>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:template match="@docx2hub:num-signature | @docx2hub:num-continue | @docx2hub:num-abstract | @docx2hub:num-id 
                       | @docx2hub:num-restart-val | @docx2hub:num-ilvl | @docx2hub:num-restart-level" mode="docx2hub:join-runs"/>
  
</xsl:stylesheet>