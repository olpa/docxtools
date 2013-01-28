<?xml version="1.0" encoding="UTF-8" ?>

<!--~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Stylesheet zur Word 2007 XML nach DocBook (version 4.4) Konvertierung

Entwicklung: le-tex publishing services oHG (2008)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~-->


<xsl:stylesheet version="2.0"

    xmlns:xsl		= "http://www.w3.org/1999/XSL/Transform"
    xmlns:fn            = "http://www.w3.org/2005/xpath-functions"
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
    xmlns:m             = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:mml           = "http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
    xmlns:mc            = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:xlink = "http://www.w3.org/1999/xlink"
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"
    exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn letex m mml mc xlink docx2hub"
>

  <!-- ================================================================================ -->
  <!-- IMPORT OF OTHER STYLESHEETS -->
  <!-- ================================================================================ -->

  <!-- Do not sort catch-all.xsl in alphabetical order. Must be first! -->
  <xsl:import href="modules/catch-all/catch-all.xsl"/>

  <!-- sorted imports -->
  <!-- To avoid conflicts sort in alphabetical order. -->
  <xsl:import href="modules/error-handler/error-handler.xsl"/>

  <xsl:import href="letex-util/resolve-uri/resolve-uri.xsl"/>


  <xsl:param name="debug" select="'yes'" as="xs:string?"/>
  <xsl:param name="srcpath" select="'no'" as="xs:string?"/>
  <xsl:param name="base-dir" select="replace(base-uri(), '[^/]+$', '')"/>
  <xsl:variable name="debug-dir" select="concat(replace($base-dir, '^(.+/)(.+?/)$', '$1'), 'debug')"/>
  <!-- Links that probably have been inserted by Word without user consent: -->
  <xsl:param name="unwrap-tooltip-links" select="'no'" as="xs:string?"/>
  <xsl:param name="hub-version" select="'1.0'" as="xs:string"/>
  
  <xsl:variable name="symbol-font-map" select="document('Symbol.xml')" as="document-node(element(symbols))" />
  <xsl:variable 
    name="footnotes" 
    select="if (doc-available(concat($base-dir, '/footnotes.xml'))) 
            then document(concat($base-dir, '/footnotes.xml'))/w:footnotes
            else ()" 
    as="element(w:footnotes)?" />
  <xsl:variable name="docRels" select="if (doc-available(concat($base-dir, '/_rels/document.xml.rels'))) 
                                            then document(concat($base-dir, '/_rels/document.xml.rels'))/rel:Relationships
                                            else ()" />
  <xsl:variable name="footnoteRels" select="if (doc-available(concat($base-dir, '/_rels/footnotes.xml.rels'))) 
                                            then document(concat($base-dir, '/_rels/footnotes.xml.rels'))/rel:Relationships
                                            else ()" />
  <xsl:variable name="commentRels" select="if (doc-available(concat($base-dir, '/_rels/comments.xml.rels'))) 
                                            then document(concat($base-dir, '/_rels/comments.xml.rels'))/rel:Relationships
                                            else ()" />

  <xsl:key name="style-by-id" match="w:style" use="@w:styleId" />
  <xsl:key name="numbering-by-id" match="w:num" use="@w:numId" />
  <xsl:key name="abstract-numbering-by-id" match="w:abstractNum" use="@w:abstractNumId" />
  <xsl:key name="footnote-by-id" match="w:footnote" use="@w:id" />
  <xsl:key name="endnote-by-id" match="w:endnote" use="@w:id" />
  <xsl:key name="comment-by-id" match="w:comment" use="@w:id" />
  <xsl:key name="doc-rel-by-id" match="w:docRels/rel:Relationships/rel:Relationship" use="@Id" />
  <xsl:key name="footnote-rel-by-id" match="w:footnoteRels/rel:Relationships/rel:Relationship" use="@Id" />
  <xsl:key name="comment-rel-by-id" match="w:commentRels/rel:Relationships/rel:Relationship" use="@Id" />
  <xsl:key name="symbol-by-number" match="symbol" use="upper-case(replace(@number, '^0*(.+?)$', '$1'))" />


  <!-- sorted includes -->
  <xsl:include href="comments.xsl"/>
  <xsl:include href="endnotes.xsl"/>
  <xsl:include href="footnotes.xsl"/>
  <xsl:include href="index.xsl"/>
  <xsl:include href="numbering.xsl"/>
  <xsl:include href="objects.xsl"/>
  <xsl:include href="omml2mml/omml2mml.xsl"/>
<!--   <xsl:include href="para-props.xsl"/> -->
  <xsl:include href="sym.xsl"/>
<!--   <xsl:include href="text-props.xsl"/> -->
  <xsl:include href="tables.xsl"/>


  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- named Templates -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <xsl:template name="handle-field-function">
    <xsl:param name="nodes" as="node()*"/>
    <xsl:param name="is-multi-para" as="xs:boolean" select="false()"/>
    <xsl:choose>
      <xsl:when test="$is-multi-para">
        <xsl:choose>
          <xsl:when test="$nodes[1][self::w:p[count(w:r/w:fldChar[@w:fldCharType = 'begin']) = 2]]">
            <xsl:if test="not($nodes[last()][self::w:p[count(w:r/w:fldChar) = 1 and w:r/w:fldChar[@w:fldCharType = 'end']]])">
              <xsl:call-template name="signal-error" xmlns="">
                <xsl:with-param name="error-code" select="'W2D_010'"/>
                <xsl:with-param name="exit" select="'yes'"/>
                <xsl:with-param name="hash">
                  <value key="xpath"><xsl:value-of select="$nodes[last()]/@xpath"/></value>
                  <value key="level">INT</value>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:if>
            <xsl:variable name="first-cut" select="($nodes[1]/w:r/w:fldChar[@w:fldCharType = 'separate'])[1]"/>
            <xsl:variable name="first-beg" select="($nodes[1]/w:r/w:fldChar[@w:fldCharType = 'begin'])[1]"/>
            <xsl:variable name="split" select="$nodes[w:r/w:fldChar[@w:fldCharType = 'begin'] and w:r/w:fldChar[@w:fldCharType = 'end']
                                               and count(w:r/w:fldChar[@w:fldCharType = 'end']) = count(w:r/w:fldChar[@w:fldCharType = 'begin'])
                                               and w:r[w:fldChar/@w:fldCharType = 'begin'][1]/preceding-sibling::w:r[w:fldChar/@w:fldCharType = 'end']
                                               and not(.//w:t)]"/>
<!--             <xsl:message> -->
<!--               <xsl:copy-of select="$first-cut" copy-namespaces="no"/> -->
<!--             </xsl:message> -->
            <xsl:variable name="first-node">
              <xsl:element name="{$nodes[1]/name()}">
                <xsl:copy-of select="$nodes[1]/@*"/>
                <xsl:copy-of select="$nodes[1]/w:pPr"/>
                <xsl:copy-of select="$nodes[1]/node()[. &gt;&gt; $first-cut]"/>                
              </xsl:element>
            </xsl:variable>
<!--             <xsl:message> -->
<!--               <xsl:copy-of select="$first-node" copy-namespaces="no"/> -->
<!--             </xsl:message> -->
            <xsl:apply-templates select="$first-node" mode="wml-to-dbk"/>
            <xsl:apply-templates select="$nodes[position() &gt; 1 and not(position() = last())] except $split" mode="wml-to-dbk"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="first-node">
              <xsl:element name="{$nodes[1]/name()}">
                <xsl:copy-of select="$nodes[1]/@*"/>
                <xsl:copy-of select="$nodes[1]/node()"/>
                <w:r>
                  <w:fldChar w:fldCharType="end"/>
                </w:r>
              </xsl:element>
            </xsl:variable>
            <xsl:apply-templates select="$first-node" mode="wml-to-dbk"/>
            <xsl:apply-templates select="$nodes[position() &gt; 1 and not(position() = last())]" mode="wml-to-dbk"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="count($nodes/w:fldChar[@w:fldCharType = 'begin']) - count($nodes/w:fldChar[@w:fldCharType = 'end']) gt 0">
          <xsl:call-template name="signal-error" xmlns="">
            <xsl:with-param name="error-code" select="'W2D_011'"/>
            <xsl:with-param name="exit" select="'yes'"/>
            <xsl:with-param name="hash">
              <value key="xpath"><xsl:value-of select="$nodes[1]/@xpath"/></value>
              <value key="level">INT</value>
              <value key="info-text"><xsl:value-of select="$nodes//text()"/></value>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
        <xsl:choose>
          <xsl:when test="count($nodes/w:fldChar[@w:fldCharType = 'begin']) eq 1">
            <xsl:choose>
              <xsl:when test="$nodes[1][self::w:r[w:fldChar[@w:fldCharType = 'begin']]]">
                <xsl:choose>
                  <xsl:when test="not($nodes[w:fldChar[@w:fldCharType = 'separate']])">
                    <xsl:variable name="end" select="$nodes[w:fldChar[@w:fldCharType = 'end']]"/>
                    <xsl:apply-templates select="($nodes[position() &gt; 1 and . &lt;&lt; $end])[1]" mode="wml-to-dbk">
                      <xsl:with-param name="instrText" select="string-join($nodes[position() &gt; 1 and . &lt;&lt; $end]//text(), '')" tunnel="yes"/>
                      <xsl:with-param name="nodes" select="$nodes[position() &gt; 1 and . &lt;&lt; $end]" tunnel="yes"/>
                      <xsl:with-param name="text" select="()" tunnel="yes"/>
                    </xsl:apply-templates>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:variable name="sep" select="$nodes[w:fldChar[@w:fldCharType = 'separate']]"/>
                    <xsl:variable name="end" select="$nodes[w:fldChar[@w:fldCharType = 'end']]"/>
                    <xsl:apply-templates select="($nodes[position() &gt; 1 and . &lt;&lt; $sep])[1]" mode="wml-to-dbk">
                      <xsl:with-param name="instrText" select="string-join($nodes[position() &gt; 1 and . &lt;&lt; $sep]//text(), '')" tunnel="yes"/>
                      <xsl:with-param name="nodes" select="$nodes[position() &gt; 1 and . &lt;&lt; $sep]" tunnel="yes"/>
                      <xsl:with-param name="text" select="$nodes[. &gt;&gt; $sep and . &lt;&lt; $end]" tunnel="yes"/>
                    </xsl:apply-templates>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="signal-error" xmlns="">
                  <xsl:with-param name="error-code" select="'W2D_012'"/>
                  <xsl:with-param name="exit" select="'yes'"/>
                  <xsl:with-param name="hash">
                    <value key="xpath"><xsl:value-of select="$nodes[1]/@xpath"/></value>
                    <value key="level">INT</value>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="count($nodes/w:fldChar[@w:fldCharType = 'begin']) gt 1">
            <xsl:choose>
              <xsl:when test="$nodes[1][self::w:r[w:fldChar[@w:fldCharType = 'begin']]]">
<!--                 <xsl:message>=====================================================</xsl:message> -->
<!--                 <xsl:message select="string-join($nodes//text()[parent::w:instrText], '')"/> -->
<!--                 <xsl:message select="string-join($nodes[.//text()[parent::w:t]], '')"/> -->
<!--                 <xsl:message select="($nodes[w:instrText])[1]"/> -->
                <xsl:apply-templates select="($nodes[w:instrText])[1]" mode="wml-to-dbk">
                  <xsl:with-param name="instrText" select="string-join($nodes//text()[parent::w:instrText], '')" tunnel="yes"/>
                  <xsl:with-param name="nodes" select="$nodes[descendant::w:instrText]" tunnel="yes"/>
                  <xsl:with-param name="text" select="$nodes[.//text()[parent::w:t] or .//w:tab or .//w:br]" tunnel="yes"/>
                </xsl:apply-templates>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="signal-error" xmlns="">
                  <xsl:with-param name="error-code" select="'W2D_012'"/>
                  <xsl:with-param name="exit" select="'yes'"/>
                  <xsl:with-param name="hash">
                    <value key="xpath"><xsl:value-of select="$nodes[1]/@xpath"/></value>
                    <value key="level">INT</value>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="signal-error" xmlns="">
              <xsl:with-param name="error-code" select="'W2D_013'"/>
              <xsl:with-param name="exit" select="'yes'"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="$nodes[1]/@xpath"/></value>
                <value key="level">INT</value>
                <value key="info-text"><xsl:value-of select="count($nodes/w:fldChar[@w:fldCharType = 'begin'])"/></value>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="check-field-functions">
    <xsl:param name="nodes" as="node()*"/>
    <xsl:for-each-group select="node()"
      group-adjacent="count(preceding::w:fldChar[@w:fldCharType = 'begin'])
                      - count(preceding::w:fldChar[@w:fldCharType = 'end'])
                      + (if (w:r/w:fldChar[@w:fldCharType = 'begin'])
                      then (count(w:r/w:fldChar[@w:fldCharType = 'begin']) - count(w:r/w:fldChar[@w:fldCharType = 'end']))
                      else 0)">
      <xsl:choose>
        <xsl:when test="current-grouping-key() &gt; 0">
          <xsl:for-each-group select="current-group()" group-starting-with="*[count(preceding::w:fldChar[@w:fldCharType = 'begin'])
                                                                            - count(preceding::w:fldChar[@w:fldCharType = 'end']) = 0]">
            <xsl:call-template name="handle-field-function">
              <xsl:with-param name="nodes" select="current-group()"/>
              <xsl:with-param name="is-multi-para" select="true()"/>
            </xsl:call-template>
            <xsl:if test="current-group()[last()]/w:r[w:fldChar[@w:fldCharType = 'end']][last()]/following-sibling::node()">
              <!-- verlorengegangenen Knoten ohne @w:fldChar reproduzieren -->
              <xsl:variable name="saved-last-node">
                <xsl:apply-templates select="current-group()[position() = last()]" mode="rescue-node"/>
              </xsl:variable>
              <xsl:apply-templates select="$saved-last-node" mode="wml-to-dbk"/>
            </xsl:if>
          </xsl:for-each-group>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="current-group()" mode="#current"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each-group>
  </xsl:template>

  <xsl:template match="w:r[w:fldChar[@w:fldCharType = 'end'] and not(following-sibling::w:r[w:fldChar])]" mode="rescue-node">
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- Mode: pre-process -->
  <!-- ================================================================================ -->

  <!-- Verlauf -->
  <xsl:template match="w:del" mode="docx2hub:separate-field-functions">
    <!-- gelöschten Text wegwerfen -->
  </xsl:template>

  <xsl:template match="w:ins" mode="docx2hub:separate-field-functions">
    <xsl:apply-templates select="node()" mode="#current"/>
  </xsl:template>

  <!-- Ende von Feldfunktionen ueber mehrere Absaetze in einzelnen Absatz packen -->
  <!-- Grund: wenn in dem gleichen Absatz eine neue Feldfunktion beginnt, liefert check-field-functions falsche Gruppen -->
  <xsl:template match="w:p[
                         w:r[w:fldChar][1][count(w:fldChar) = 1]
                         /w:fldChar[@w:fldCharType='end']
                       ][count(w:r[w:fldChar]) gt 1]" mode="docx2hub:separate-field-functions">
    <xsl:variable name="attribute-names" as="xs:string *">
      <xsl:for-each select="@*">
        <xsl:sequence select="name(.)"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="attribute-values" as="xs:string *">
      <xsl:for-each select="@*">
        <xsl:sequence select="."/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="pPr" as="node() *">
      <xsl:apply-templates select="w:pPr" mode="#current"/>
    </xsl:variable>
    <xsl:for-each-group select="node()" group-ending-with="w:r[w:fldChar][1]">
      <w:p>
        <xsl:for-each select="$attribute-names">
          <xsl:variable name="pos" select="position()"/>
          <xsl:attribute name="{.}" select="$attribute-values[position() eq $pos]"/>
        </xsl:for-each>
        <xsl:copy-of select="$pPr"/>
        <xsl:apply-templates select="current-group()[not(self::w:pPr)]"/>
      </w:p>
    </xsl:for-each-group>
  </xsl:template>

  <xsl:template match="w:p[
                         w:r[last()][w:fldChar][count(w:fldChar) = 1]
                         /w:fldChar[@w:fldCharType='end']
                       ][
                         count(w:r[w:fldChar[@w:fldCharType='end']])
                         gt
                         count(w:r[w:fldChar[@w:fldCharType='begin']])
                       ][
                         count(w:r[w:fldChar[@w:fldCharType='end']]) gt 1
                       ]" mode="docx2hub:separate-field-functions">
    <xsl:variable name="attribute-names" as="xs:string *">
      <xsl:for-each select="@*">
        <xsl:sequence select="name(.)"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="attribute-values" as="xs:string *">
      <xsl:for-each select="@*">
        <xsl:sequence select="."/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="pPr" as="node() *">
      <xsl:apply-templates select="w:pPr" mode="#current"/>
    </xsl:variable>
    <xsl:for-each-group select="node()" group-ending-with="w:r[position() lt last()][w:fldChar][last()]">
      <w:p>
        <xsl:for-each select="$attribute-names">
          <xsl:variable name="pos" select="position()"/>
          <xsl:attribute name="{.}" select="$attribute-values[position() eq $pos]"/>
        </xsl:for-each>
        <xsl:copy-of select="$pPr"/>
        <xsl:apply-templates select="current-group()[not(self::w:pPr)]"/>
      </w:p>
    </xsl:for-each-group>
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- Mode: wml-to-dbk -->
  <!-- ================================================================================ -->

  <!-- default for elements -->
  <xsl:template match="*" mode="wml-to-dbk">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_020'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@xpath"/></value>
        <value key="level">INT</value>
        <value key="mode">wml-to-dbk</value>
        <value key="info-text"><xsl:value-of select="concat('Element: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- GI 2012-10-08 §§§
       Want to get rid of the warnings. Does that hurt? Not tested.
       -->
  <xsl:template match="w:p/w:numPr | css:rule/w:numPr | style/w:numPr |
                       /*/w:numbering | /*/w:docRels | /*/w:fonts | /*/w:comments" mode="wml-to-dbk" priority="-0.25"/>

  <xsl:template match="dbk:* | css:*" mode="wml-to-dbk">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current" />
      <xsl:apply-templates select="node()" mode="#current" />
    </xsl:copy>
  </xsl:template>

  <!-- default for attributes -->
  <xsl:template match="@*" mode="wml-to-dbk" priority="1">
    <xsl:copy/>
  </xsl:template>

  <xsl:template match="@*" mode="wml-to-dbk">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_021'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="../@xpath"/></value>
        <value key="level">INT</value>
        <value key="mode">wml-to-dbk</value>
        <value key="info-text"><xsl:value-of select="concat('Attribut: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- default for comments -->
  <xsl:template match="comment()" mode="wml-to-dbk">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_022'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@xpath"/></value>
        <value key="level">INT</value>
        <value key="mode">wml-to-dbk</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- standard Kommentar von Aspose nicht warnen -->
  <xsl:template match="comment()[matches(., '^ Generated by Aspose')]" mode="wml-to-dbk"/>

  <!-- default for PIs -->
  <xsl:template match="processing-instruction()" mode="wml-to-dbk">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_023'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@xpath"/></value>
        <value key="level">INT</value>
        <value key="mode">wml-to-dbk</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="w:commentRels" mode="wml-to-dbk"/>


  <!-- element section -->


  <xsl:template match="w:document" mode="wml-to-dbk">
    <xsl:apply-templates select="@* except @xpath | node()" mode="wml-to-dbk"/>
  </xsl:template>

  <xsl:template match="@mc:Ignorable" mode="wml-to-dbk"/>

  <xsl:template match="/dbk:*" mode="wml-to-dbk">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="wml-to-dbk"/>
      <xsl:call-template name="check-field-functions">
        <xsl:with-param name="nodes" select="node()"/>
      </xsl:call-template>
    </xsl:copy>
  </xsl:template>

  <!-- paragraphs (w:p) -->

  <xsl:variable name="docx2hub:allowed-para-element-names" as="xs:string+"
    select="('w:r', 'w:pPr', 'w:bookmarkStart', 'w:bookmarkEnd', 'w:smartTag', 'w:commentRangeStart', 'w:commentRangeEnd', 'w:proofErr', 'w:hyperlink', 'w:del', 'w:ins', 'w:fldSimple', 'm:oMathPara', 'm:oMath')" />

  <xsl:template match="w:p" mode="wml-to-dbk">
    <xsl:element name="para">
      <xsl:apply-templates select="@* except @*[matches(name(),'^w:rsid')]" mode="#current"/>
      <xsl:if test="false() (: §§§ :)
                    and
                    (some $x in * satisfies not($x/name() = $docx2hub:allowed-para-element-names))">
        <xsl:call-template name="signal-error" xmlns="">
          <xsl:with-param name="error-code" select="'W2D_030'"/>
          <xsl:with-param name="exit" select="'yes'"/>
          <xsl:with-param name="hash">
            <value key="xpath"><xsl:value-of select="@xpath"/></value>
            <value key="level">INT</value>
            <value key="mode">wml-to-dbk</value>
            <value key="info-text"><xsl:value-of select="string-join(*[not(name() = $docx2hub:allowed-para-element-names)]/name(), ' ')"/></value>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:if>
      <xsl:if test=".//w:r">
        <xsl:sequence select="letex:insert-numbering(.)"/>
      </xsl:if>
      <xsl:choose>
        <xsl:when test="w:r[w:fldChar]">
          <xsl:call-template name="inline-field-function"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates mode="#current"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:element>
  </xsl:template>

  <xsl:template name="inline-field-function">
    <xsl:variable name="starts" select="count(w:r[w:fldChar/@w:fldCharType = 'begin'])"/>
    <xsl:variable name="ends" select="count(w:r[w:fldChar/@w:fldCharType = 'end'])"/>
    <xsl:variable name="seps" select="count(w:r[w:fldChar/@w:fldCharType = 'separate'])"/>
    <xsl:if test="$starts lt $ends">
      <xsl:call-template name="signal-error" xmlns="">
        <xsl:with-param name="error-code" select="'W2D_014'"/>
        <xsl:with-param name="exit" select="'yes'"/>
        <xsl:with-param name="hash">
          <value key="xpath"><xsl:value-of select="@xpath"/></value>
          <value key="level">INT</value>
          <value key="info-text"><xsl:value-of select="."/></value>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
    <xsl:for-each-group select="node()" 
      group-adjacent="if (count(self::w:r[w:fldChar/@w:fldCharType='begin'])
                      + count(preceding-sibling::w:r[w:fldChar/@w:fldCharType='begin'])
                      (: - count(self::w:r[w:fldChar/@w:fldCharType='end']) :)
                      - count(preceding-sibling::w:r[w:fldChar/@w:fldCharType='end']) &gt; 0) then true() else false()">
      <xsl:choose>
        <xsl:when test="current-grouping-key()">
          <xsl:call-template name="handle-field-function">
            <xsl:with-param name="nodes" select="current-group()"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="current-group()" mode="#current"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each-group>
  </xsl:template>

  <!-- Verlauf -->
  <xsl:template match="w:del" mode="wml-to-dbk">
    <!-- gelöschten Text wegwerfen -->
  </xsl:template>

  <xsl:template match="w:ins" mode="wml-to-dbk">
    <xsl:apply-templates select="node()" mode="#current"/>
  </xsl:template>


  <!-- bookmarks -->

  <xsl:template match="w:bookmarkStart" mode="wml-to-dbk">
    <anchor role="start" xml:id="{@w:name}" xreflabel="{@w:id}"/>
  </xsl:template>

  <xsl:template match="w:bookmarkEnd" mode="wml-to-dbk">
    <anchor role="end" xml:id="{concat(preceding::w:bookmarkStart[@w:id = current()/@w:id]/@w:name, '_end')}"/>
  </xsl:template>

  <!-- comments -->
  <xsl:template match="w:commentRangeStart" mode="wml-to-dbk"/>
  <xsl:template match="w:commentRangeEnd" mode="wml-to-dbk"/>
  <xsl:template match="w:proofErr" mode="wml-to-dbk"/>



  <!-- paragraph properties (w:pPr) -->

  <xsl:template match="w:pPr" mode="wml-to-dbk">
    <!-- para properties are collected by para-props.xsl -->
  </xsl:template>

  <!-- run properties (w:rPr) -->

  <xsl:template match="w:rPr" mode="wml-to-dbk">
    <!-- run properties are collected when text nodes are handled -->
  </xsl:template>

  <!-- smartTags (w:smartTag) -->
  <xsl:template match="w:smartTag" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <!-- Hyperlinks -->
  <xsl:template match="w:hyperlink" mode="wml-to-dbk">
    <link>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:choose>
        <xsl:when test="w:r[w:fldChar]">
          <xsl:call-template name="inline-field-function"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="node()" mode="#current"/>
        </xsl:otherwise>
      </xsl:choose>
    </link>
  </xsl:template>

  <xsl:template match="w:hyperlink[@w:tooltip][$unwrap-tooltip-links = 'yes']" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="@w:anchor[parent::w:hyperlink]" mode="wml-to-dbk" priority="1.5">
    <xsl:choose>
      <xsl:when test="exists(parent::w:hyperlink/@r:id)"/>
      <xsl:otherwise>
        <xsl:attribute name="linkend" select="."/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="@w:tooltip[parent::w:hyperlink]" mode="wml-to-dbk" priority="1.5">
    <!-- p1609: The method by which this string is surfaced by an application is outside the scope of this Office Open XML Standard. -->
  </xsl:template>

  <xsl:template match="@r:id[parent::w:hyperlink]" mode="wml-to-dbk" priority="1.5">
    <xsl:variable name="key-name" as="node()*"
      select="if (ancestor::w:footnote)
              then $footnoteRels
              else if (ancestor::w:comment) 
                then $commentRels
                else $docRels" />
    <xsl:variable name="value" select="."/>
    <xsl:variable name="rel-item" select="$key-name/rel:Relationship[@Id=$value]"/>
    <xsl:choose>
      <xsl:when test="exists(parent::w:hyperlink/@w:anchor)">
        <xsl:attribute name="linkend" select="concat(
                                                $rel-item/@Target,
                                                '#',
                                                parent::w:hyperlink/@w:anchor
                                              )"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$rel-item/@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'">
            <xsl:attribute name="xlink:href" select="$rel-item/@Target" />
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="linkend" select="$rel-item/@Target"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="@w:history[parent::w:hyperlink]" mode="wml-to-dbk" priority="1.5"/>


  <xsl:template match="w:smartTagPr" mode="wml-to-dbk"/>


  <!-- textbox -->
  <xsl:template match="w:txbxContent" mode="wml-to-dbk">
    <xsl:apply-templates select="@* except @xpath" mode="#current"/>
    <xsl:apply-templates mode="#current"/>
  </xsl:template>


  <!-- runs (w:r) -->
  <xsl:template match="w:r[@*]" mode="wml-to-dbk">
    <xsl:element name="phrase">
      <xsl:apply-templates select="@* except @*[matches(name(),'^w:rsid')]" mode="#current"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:element>
  </xsl:template>

  <xsl:template match="w:r" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <!-- text (w:t) -->
  <xsl:template match="w:t" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <!-- instrText (w:instrText)  [not(../preceding-sibling::*[w:instrText])] -->
  <xsl:template match="w:instrText" mode="wml-to-dbk" priority="20">
    <xsl:param name="instrText" as="xs:string?" tunnel="yes"/>
    <xsl:param name="text" as="node()*" tunnel="yes"/>
    <xsl:param name="nodes" as="node()*" tunnel="yes"/>
    <xsl:variable name="tokens" select="tokenize(normalize-space($instrText), ' ')"/>
    <xsl:variable name="func" select="doc('')//letex:field-functions/letex:field-function[@name = $tokens[1]]"/>
    <xsl:choose>
      <xsl:when test="not($func)">
        <xsl:choose>
          <xsl:when test="$tokens[1] = 'SYMBOL'">
            <!-- Template in sym.xsl -->
            <xsl:call-template name="create-symbol">
              <xsl:with-param name="tokens" select="$tokens"/>
              <xsl:with-param name="context" select=".."/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$tokens[1] = ('XE', 'xe')">
            <xsl:call-template name="handle-index">
              <xsl:with-param name="instr" select="$instrText"/>
              <xsl:with-param name="text" select="$text"/>
              <xsl:with-param name="nodes" select="$nodes"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$tokens[1] = ('EQ','eq','FORMCHECKBOX')">
            <xsl:call-template name="signal-error" xmlns="">
              <xsl:with-param name="error-code" select="'W2D_045'"/>
              <xsl:with-param name="exit"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="@xpath"/></value>
                <value key="level">WRN</value>
                <value key="info-text"><xsl:value-of select="$instrText"/></value>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$tokens[1] = 'INCLUDEPICTURE'">
            <xsl:call-template name="handle-figures">
              <xsl:with-param name="instr" select="$instrText"/>
              <xsl:with-param name="text" select="$text"/>
              <xsl:with-param name="nodes" select="$nodes"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="matches($instrText,'^[\s&#160;]*$')"/>
          <xsl:otherwise>
            <xsl:call-template name="signal-error" xmlns="">
              <xsl:with-param name="error-code" select="'W2D_040'"/>
              <xsl:with-param name="exit" select="'yes'"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="@xpath"/></value>
                <value key="level">INT</value>
                <value key="info-text"><xsl:value-of select="$instrText"/></value>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$func/@element">
            <xsl:element name="{$func/@element}">
              <xsl:if test="$func/@attrib">
                <xsl:attribute name="{$func/@attrib}" select="replace($tokens[position() = $func/@value], '&quot;', '')"/>
                <xsl:apply-templates select="$text" mode="#current"/>
              </xsl:if>
            </xsl:element>
          </xsl:when>
          <xsl:when test="$func/@destroy = 'yes'"/>
          <xsl:otherwise>
            <xsl:apply-templates select="$text" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <letex:field-functions>
    <letex:field-function name="HYPERLINK" element="link" attrib="xlink:href" value="2"/>
    <letex:field-function name="INDEX" destroy="yes"/>
    <letex:field-function name="NOTEREF" element="link" attrib="linkend" value="2"/>
    <letex:field-function name="PAGE"/>
    <letex:field-function name="PAGEREF" element="link" attrib="linkend" value="2"/>
    <letex:field-function name="PRINT" destroy="yes"/>
    <letex:field-function name="RD"/>
    <letex:field-function name="REF"/>
    <letex:field-function name="ADVANCE"/>
    <letex:field-function name="QUOTE"/>
    <letex:field-function name="SEQ"/>
    <letex:field-function name="SET" destroy="yes"/>
    <letex:field-function name="STYLEREF"/>
    <letex:field-function name="TOC" destroy="yes"/>
    <letex:field-function name="\IF"/>
  </letex:field-functions>

  <xsl:template name="handle-figures">
    <xsl:param name="inline" select="false()" as="xs:boolean" tunnel="yes"/>
    <xsl:param name="instr" as="xs:string?"/>
    <xsl:param name="text" as="node()*"/>
    <xsl:param name="nodes" as="node()*"/>
    <xsl:variable name="text-tokens" select="for $x in $nodes//text() return $x"/>
    <xsl:element name="mediaobject">
      <xsl:attribute name="style" select="'imagedata'"/>
      <imageobject>
        <xsl:attribute name="style" select="'width:3in;height:3in'"/>
        <imagedata fileref="{replace(tokenize($instr, ' ')[matches(.,'^&#x22;.*&#x22;$')][1],'&#x22;','')}"/>
      </imageobject>
    </xsl:element>
  </xsl:template>

  <!-- w:sectPr ignorieren -->
  <xsl:template match="w:sectPr" mode="wml-to-dbk"/>

  <xsl:template match="w:tcPr" mode="wml-to-dbk"/>

  <!-- Umbruchshilfe zur exakten Reproduktion des Umbruchs -->
  <xsl:template match="w:lastRenderedPageBreak" mode="wml-to-dbk">
  </xsl:template>


  <!-- Background -->
  <xsl:template match="w:background[parent::w:document]" mode="wml-to-dbk">
  </xsl:template>

  <!-- fldSimple -->
  <xsl:template match="w:fldSimple" mode="wml-to-dbk">
    <!-- p1592 gehört zu Feldfunktionen. Wenn w:t darunter, muss der geschrieben werden -->
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <!-- drawing -->
  <xsl:template match="w:drawing" mode="wml-to-dbk">
    <!-- weg damit, wenn dann in extra include -->
  </xsl:template>

  <!-- whitespace elements, etc. -->
  <xsl:template match="w:tab" mode="wml-to-dbk">
    <tab xml:space="preserve">&#9;</tab>
  </xsl:template>

  <xsl:template match="w:br" mode="wml-to-dbk">
    <br/>
  </xsl:template>

  <xsl:template match="w:cr" mode="wml-to-dbk">
    <!-- carriage return -->
    <!-- ggf. als echte Absatzmarke behandeln. Dazu muss ein nachgelagerter neuer mode eingefuehrt werden. -->
    <phrase role="cr"/>
  </xsl:template>

  <xsl:template match="w:softHyphen" mode="wml-to-dbk">
  </xsl:template>

  <xsl:template match="w:noBreakHyphen" mode="wml-to-dbk">
      <xsl:value-of select="'&#x2011;'"/>
  </xsl:template>

  <xsl:template match="w:pageBreakBefore" mode="wml-to-dbk">
    <xsl:if test="@w:val = ('true', '1', 'on') or not(@w:val)">
      <phrase role="pageBreakBefore"/>
    </xsl:if>
  </xsl:template>

  <!-- math section -->
  <xsl:template match="m:oMathPara" mode="wml-to-dbk">
    <equation role="omml">
      <xsl:apply-templates select="@* except @xpath" mode="#current"/>
      <xsl:apply-templates select="node()" mode="omml2mml"/>
    </equation>
  </xsl:template>

  <xsl:template match="m:oMath" mode="wml-to-dbk">
    <inlineequation role="omml">
      <xsl:apply-templates select="@* except @xpath" mode="#current"/>
      <xsl:apply-templates select="." mode="omml2mml"/>
    </inlineequation>
  </xsl:template>


  <!-- attribute section -->

  <xsl:template match="@*[(parent::w:p or parent::w:r) and matches(name(), '^w:rsid')]" mode="wml-to-dbk">
    <!-- IDs zur Kennzeichnung des Verlaufs im Word-Dokument ignorieren -->
  </xsl:template>

  <xsl:template match="@w:val[parent::*/name() = (
                      'w:pStyle'
                      )]" mode="wml-to-dbk">
    <!-- Attributswerte, die in anderem Kontext bereits ausgegeben werden -->
  </xsl:template>

  <xsl:template match="@*[parent::w:smartTag]" mode="wml-to-dbk">
    <!-- Attribute von smartTag vorerst ignoriert. -->
  </xsl:template>



  <xsl:template match="@role" mode="wml-to-dbk" priority="2">
    <xsl:attribute name="role" select="replace(., ' ', '_')" />
  </xsl:template>

  <xsl:template match="@srcpath[$srcpath != 'yes']" mode="wml-to-dbk" priority="2" />


  <xsl:function name="docx2hub:twips2mm" as="xs:string">
    <xsl:param name="val" as="xs:integer"/>
    <xsl:sequence select="concat(xs:string($val * 0.01763889), 'mm')" />
  </xsl:function>



</xsl:stylesheet>

