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
    <xsl:param name="context" as="element(w:p)"/>
    
    <xsl:variable name="lvl" select="letex:get-lvl-of-numbering($context)" as="element(w:lvl)?"/>
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
        <xsl:apply-templates mode="numbering"
          select="$lvl/w:pPr/@*[letex:different-style-att(., $context)]">
          <xsl:with-param name="context" select="$context" tunnel="yes"/>
        </xsl:apply-templates>
        <xsl:apply-templates select="$context/dbk:tabs" mode="wml-to-dbk"/>
        <phrase role="hub:identifier">
          <xsl:apply-templates mode="numbering"
            select="$lvl/w:rPr/@*[letex:different-style-att(., $context)]" >
            <xsl:with-param name="context" select="$context" tunnel="yes"/>
          </xsl:apply-templates>
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

  <xsl:key name="style-by-name" match="css:rule | dbk:style" use="@name | @role"/>

  <xsl:function name="letex:different-style-att" as="xs:boolean">
    <xsl:param name="att" as="attribute(*)"/>
    <xsl:param name="context" as="element(*)"/>
    <xsl:variable name="context-atts" select="key('style-by-name', $context/@role, $att/root())/@* | $context/@*" as="attribute(*)*"/>
    <xsl:sequence select="not(some $a in $context-atts satisfies ($a/name() = $att/name() and string($a) = string($att)))"/>
  </xsl:function>

  <xsl:function name="letex:get-lvl-of-numbering" as="element(w:lvl)?">
    <xsl:param name="context" as="node()"/>
    <!-- for-each: just to avoid an XTDE1270 which shouldn't happen when the 3-arg form of key() is invoked: -->
    <xsl:for-each select="$context">
    <xsl:variable name="numPr" select="if ($context/w:numPr) 
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
  
  <xsl:function name="letex:get-identifier" as="xs:string">
    <xsl:param name="context" as="element(w:p)"/>
    <xsl:param name="lvl" as="element(w:lvl)"/>
    
    <xsl:variable name="abstract-num-id" select="xs:double($lvl/ancestor::w:abstractNum/@w:abstractNumId)" as="xs:double"/>
    <xsl:variable name="lvl-to-use" select="if (exists(letex:get-lvl-override($context)/w:lvl)) 
                                            then letex:get-lvl-override($context)/w:lvl 
                                            else $lvl"/>
    <xsl:variable name="ilvl" select="xs:double($lvl-to-use/@w:ilvl)"/>
    
    <xsl:variable name="context-relevant" 
                  select="if (matches($context/@role,'List')) 
                          then $context/preceding-sibling::w:p
      [
        (
          exists(letex:get-lvl-of-numbering(.)) 
            and 
          $abstract-num-id = xs:double(letex:get-lvl-of-numbering(.)/ancestor::w:abstractNum/@w:abstractNumId)
        ) 
          or
        (
          not(exists(letex:get-lvl-of-numbering(.))) 
            and 
          not(w:numPr)
        ) 
          or
        (
          exists(letex:get-lvl-of-numbering(.)) 
            and 
          not($abstract-num-id = xs:double(letex:get-lvl-of-numbering(.)/ancestor::w:abstractNum/@w:abstractNumId)) 
            and
          (some $i 
           in ((xs:string(./@role), xs:string(letex:get-lvl-of-numbering(.)/ancestor::w:abstractNum/w:lvl[@w:ilvl=$ilvl]/w:pStyle/@w:val))) 
           satisfies (($i = xs:string($context/@role)) or ($i=$lvl/ancestor::w:abstractNum/w:lvl[@w:ilvl=$ilvl or @w:ilvl=$ilvl -1 or @w:ilvl=$ilvl+1]/w:pStyle/@w:val))) 
            and 
          (letex:get-lvl-of-numbering(.)/@w:ilvl=$ilvl+1 or letex:get-lvl-of-numbering(.)/@w:ilvl=$ilvl -1)
        )
      ]
      [ 
        not(following-sibling::w:p
          [ . &lt;&lt; $context ]
          [ 
            not( 
              (
                exists(letex:get-lvl-of-numbering(.)) 
                  and 
                $abstract-num-id = xs:double(letex:get-lvl-of-numbering(.)/ancestor::w:abstractNum/@w:abstractNumId)
              ) 
                or
              (
                not(exists(letex:get-lvl-of-numbering(.))) 
                  and 
                not(w:numPr)
              ) 
                or
              (
                exists(letex:get-lvl-of-numbering(.)) 
                  and 
                not($abstract-num-id = xs:double(letex:get-lvl-of-numbering(.)/ancestor::w:abstractNum/@w:abstractNumId)) 
                  and
                (some $i 
                 in ((xs:string(./@role), xs:string(letex:get-lvl-of-numbering(.)/ancestor::w:abstractNum/w:lvl[@w:ilvl=$ilvl]/w:pStyle/@w:val))) 
                 satisfies (($i = xs:string($context/@role)) or ($i=$lvl/ancestor::w:abstractNum/w:lvl[@w:ilvl=$ilvl or @w:ilvl=$ilvl -1 or @w:ilvl=$ilvl+1]/w:pStyle/@w:val))) 
                  and 
                (letex:get-lvl-of-numbering(.)/@w:ilvl=$ilvl+1 or letex:get-lvl-of-numbering(.)/@w:ilvl=$ilvl -1)
              )
            ) 
          ] 
        ) 
      ] 
                          else $context/preceding-sibling::w:p[letex:get-lvl-of-numbering(.)/ancestor::w:abstractNum/@w:abstractNumId=$abstract-num-id]" 
                  as="element(w:p)*"/>
    <xsl:variable name="resolve-symbol-encoding">
      <element>
        <xsl:apply-templates select="$lvl-to-use/w:lvlText/@w:val" mode="wml-to-dbk"/>
      </element>
    </xsl:variable>
    <xsl:variable name="string">
      <xsl:choose>
        <xsl:when test="$resolve-symbol-encoding//@w:val">
          <xsl:analyze-string select="$lvl-to-use/w:lvlText/@w:val" regex="%([0-9])">
            <xsl:matching-substring>
              <xsl:variable name="level-counter" select="letex:get-level-counter(
                                                           $context-relevant,
                                                           $lvl-to-use,
                                                           xs:double(regex-group(1))-1,
                                                           if (exists(letex:get-lvl-override($context)/w:startOverride)) 
                                                           then letex:get-lvl-override($context)/w:startOverride/@w:val 
                                                           else if ($lvl-to-use/w:start/@w:val) 
                                                                then $lvl-to-use/w:start/@w:val 
                                                                else 1)"/>
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
          <xsl:value-of select="$resolve-symbol-encoding//text()"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:value-of select="string-join($string,'')"/>
  </xsl:function>
  
  <xsl:function name="letex:get-level-counter" as="xs:double">
    <xsl:param name="context-relevant" as="element(w:p)*"/>
    <xsl:param name="lvl" as="element(w:lvl)"/>
    <xsl:param name="ilvl" as="xs:double"/>
    <xsl:param name="start" as="xs:double"/>
    
    <xsl:variable name="restart-point" select="if (exists($context-relevant[letex:get-lvl-of-numbering(.)/@w:ilvl &lt; $ilvl][last()])) 
                                               then $context-relevant[letex:get-lvl-of-numbering(.)/@w:ilvl &lt; $ilvl][last()] 
                                               else ()"/>
    <xsl:variable name="restart" select="if (exists($lvl/w:lvlRestart)) 
                                         then if ($lvl/w:lvlRestart = '0') 
                                              then false() 
                                              else true() 
                                         else true()"/>
    <xsl:variable name="count-precedings" select="(if (exists($restart-point) and $restart) 
                                                   then count($context-relevant[. &gt;&gt; $restart-point][letex:get-lvl-of-numbering(.)/@w:ilvl = $ilvl]) 
                                                   else count($context-relevant[letex:get-lvl-of-numbering(.)/@w:ilvl = $ilvl])) 
                                                  + $start"/>
    
    <xsl:value-of select="$count-precedings"/>
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
          <xsl:with-param name="exit" select="'yes'"/>
          <xsl:with-param name="hash">
            <value key="level">INT</value>
            <value key="info-text"><xsl:value-of select="$format"/>INT</value>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
</xsl:stylesheet>