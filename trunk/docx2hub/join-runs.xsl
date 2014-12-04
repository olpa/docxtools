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
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  xmlns:mml             = "http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"

  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn letex mml"
>

  <!-- w:r is here for historic reasons. We used to group the text runs
       prematurely until we found out that it is better to group when
       there's docbook markup. So we implemented the special case of
       dbk:anchors (corresponds to w:bookmarkStart/End) only for dbk:anchor. 
       dbk:anchors between identically formatted phrases will be merged into
       with the phrases' content into a consolidated phrase. -->
  <xsl:template match="*[w:r or dbk:phrase]" mode="docx2hub:join-runs" priority="3">
    <!-- move sidebars to para level --><xsl:variable name="context" select="."/>
    <xsl:if test="self::dbk:para and .//dbk:sidebar">
      <xsl:call-template name="docx2hub_move-invalid-sidebar-elements"/>
    </xsl:if>
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:call-template name="docx2hub_pagebreak-elements-to-attributes"/>
      <xsl:for-each-group select="node()" group-adjacent="letex:signature(.)">
        <xsl:choose>
          <xsl:when test="current-grouping-key() eq ''">
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy copy-namespaces="no">
              <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
              <xsl:copy-of select="current-group()/@srcpath"/>
              <xsl:apply-templates select="current-group()[not(self::dbk:anchor)]/node() 
                                           union current-group()[self::dbk:anchor]" mode="#current" />
            </xsl:copy>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>

  <!-- collateral: replace name of mapped symbols with default Unicode font name -->
  <xsl:template match="@css:font-family[. = $docx2hub:symbol-font-names][.. = docx2hub:font-map(.)/symbol/@char]"
    mode="docx2hub:join-runs">
    <xsl:attribute name="{name()}" select="$docx2hub:symbol-replacement-rfonts/@w:ascii"/>
  </xsl:template>

  <xsl:template match="dbk:para" mode="docx2hub:join-runs">
    <xsl:call-template name="docx2hub_move-invalid-sidebar-elements"/>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:call-template name="docx2hub_pagebreak-elements-to-attributes"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template name="docx2hub_move-invalid-sidebar-elements">
    <xsl:for-each select=".//dbk:sidebar">
      <xsl:copy>
        <xsl:apply-templates select="@*" mode="#current"/>
        <xsl:attribute name="linkend" select="concat('id_', generate-id(.))"/>
        <xsl:apply-templates select="node()" mode="#current"/>
      </xsl:copy>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="docx2hub_pagebreak-elements-to-attributes">
    <xsl:apply-templates select=".//dbk:br[@role[not(. eq 'textWrapping')]]
                                          [dbk:same-scope(., current())]" 
      mode="docx2hub:join-runs-br-attr"/>
  </xsl:template>

  <xsl:function name="letex:signature" as="xs:string?">
    <xsl:param name="node" as="node()?" />
    <xsl:variable name="result-strings" as="xs:string*">
      <xsl:apply-templates select="$node" mode="docx2hub:join-runs-signature" />
    </xsl:variable>
    <xsl:value-of select="string-join($result-strings,'')"/>
  </xsl:function>

  <xsl:template match="dbk:phrase" mode="docx2hub:join-runs-signature">
    <xsl:sequence select="string-join(
                            (: don't join runs that contain field chars or instrText :)
                            (name(), w:fldChar/@w:fldCharType, w:instrText/name(), letex:attr-hashes(.)), 
                            '___'
                          )" />
  </xsl:template>

  <xsl:template match="dbk:anchor[
                         letex:signature(following-sibling::node()[not(self::dbk:anchor)][1]/self::element())
                         =
                         letex:signature(preceding-sibling::node()[not(self::dbk:anchor)][1]/self::element())
                       ]" mode="docx2hub:join-runs-signature">
    <xsl:apply-templates select="preceding-sibling::node()[not(self::dbk:anchor)][1]" mode="docx2hub:join-runs-signature" />
  </xsl:template>

  <xsl:template match="node()" mode="docx2hub:join-runs-signature">
    <xsl:sequence select="''" />
  </xsl:template>

  <xsl:function name="letex:attr-hashes" as="xs:string*">
    <xsl:param name="elt" as="node()*" />
    <xsl:perform-sort>
      <xsl:sort/>
      <xsl:sequence select="for $a in ($elt/@* except ($elt/@letex:processed, $elt/@srcpath)) return letex:attr-hash($a)" />
    </xsl:perform-sort>
  </xsl:function>

  <xsl:function name="letex:attr-hash" as="xs:string">
    <xsl:param name="att" as="attribute(*)" />
    <xsl:sequence select="concat(name($att), '__=__', $att)" />
  </xsl:function>

  <xsl:function name="letex:attname" as="xs:string">
    <xsl:param name="hash" as="xs:string" />
    <xsl:value-of select="replace($hash, '__=__.+$', '')" />
  </xsl:function>

  <xsl:function name="letex:attval" as="xs:string">
    <xsl:param name="hash" as="xs:string" />
    <xsl:value-of select="replace($hash, '^.+__=__', '')" />
  </xsl:function>
  
  <!-- @type = ('column', 'page') --> 
  <xsl:template match="dbk:br[@role[not(. eq 'textWrapping')]]" mode="docx2hub:join-runs-br-attr">
    <xsl:choose>
      <xsl:when test="dbk:before-text-in-para(., ancestor::dbk:para[1])">
        <xsl:attribute name="css:page-break-before" select="'always'"/>
      </xsl:when>
      <xsl:when test="dbk:after-text-in-para(., ancestor::dbk:para[1])">
        <xsl:attribute name="css:page-break-after" select="'always'"/>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="dbk:same-scope" as="xs:boolean">
    <xsl:param name="node" as="node()" />
    <xsl:param name="ancestor-elt" as="element(*)*" />
    <xsl:sequence select="not($node/ancestor::*[self::entry 
                                                or self::footnote
                                                or self::annotation
                                                or self::figure
                                                or self::listitem]
                                               [some $a in ancestor::* satisfies (some $b in $ancestor-elt satisfies ($a is $b))])" />
  </xsl:function>
  
  <xsl:function name="dbk:before-text-in-para" as="xs:boolean">
    <xsl:param name="elt" as="element(*)"/><!-- typically dbk:br[@role = 'page'] -->
    <xsl:param name="para" as="element(dbk:para)"/>
    <xsl:sequence select="not($para//text())
                          or
                          (
                            dbk:same-scope($elt, $para)
                            and
                            not( some $text in $para//text()[dbk:same-scope(., $para)] 
                                 satisfies ($text &lt;&lt; $elt) 
                            )
                          )"/>
  </xsl:function>

  <xsl:function name="dbk:after-text-in-para" as="xs:boolean">
    <xsl:param name="elt" as="element(*)"/><!-- typically dbk:br[@role = 'page'] -->
    <xsl:param name="para" as="element(dbk:para)"/>
    <xsl:sequence select="dbk:same-scope($elt, $para)
                          and
                          not( some $text in $para//text()[dbk:same-scope(., $para)] 
                               satisfies ($text &gt;&gt; $elt) 
                          )"/>
  </xsl:function>

  <xsl:template match="dbk:br[@role[not(. eq 'textWrapping')]]
                             [
                               dbk:before-text-in-para(., ancestor::dbk:para[1])
                               or dbk:after-text-in-para(., ancestor::dbk:para[1])
                             ]" mode="docx2hub:join-runs"/>

  <!-- sidebar -->
  <xsl:template match="dbk:sidebar" mode="docx2hub:join-runs">
    <anchor xml:id="side_{generate-id(.)}"/>
  </xsl:template>

  <!-- mode hub:fix-libre-office-issues -->
  <!-- style names from Libre Office templates are coded in @native-name only, the @name attributes are automatically numbered 'style65' etc.
        therefore the @names are replace by normalized @native-names in paragraphs and css:rules -->
  
  <xsl:function name="docx2hub:normalize-to-css-name" as="xs:string">
    <xsl:param name="style-name" as="xs:string"/>
    <xsl:sequence select="replace(replace(replace($style-name, '[^-_~a-z0-9]', '_', 'i'), '~', '_-_'), '^(\I)', '_$1')"/>
<!--    <xsl:sequence select="replace($style-name, '~', '_-_')"/>-->
  </xsl:function>
  
  <xsl:key name="natives" match="css:rule" use="@name"/> 
  
  <xsl:variable name="is-libre-office-document"
              select="if (/dbk:hub/dbk:info/dbk:keywordset/dbk:keyword[@role = 'source-application'][matches(., '^LibreOffice', 'i')]) 
                      then true() 
                      else false()" as="xs:boolean"/> 
  
  <xsl:template match="css:rule[$is-libre-office-document][matches(@native-name, '^p$')]/@native-name" mode="docx2hub:join-runs">
    <xsl:attribute name="{name()}">
      <xsl:sequence select="'para'"/>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template match="css:rule[$is-libre-office-document]/@name" mode="docx2hub:join-runs">
    <xsl:attribute name="{name()}">
      <xsl:choose>
        <xsl:when test="matches(../@native-name, '(Kein Absatzformat|^\s*$)')">
          <xsl:sequence select="'None'"/>
        </xsl:when>
        <xsl:when test="matches(../@native-name, '^p$')">
          <xsl:sequence select="'para'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="docx2hub:normalize-to-css-name(../@native-name)"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute> 
  </xsl:template>
  
  <xsl:template match="*[not((local-name(.) = ('keyword', 'keywordset', 'anchor')))][$is-libre-office-document]/@role" mode="docx2hub:join-runs">
    <xsl:attribute name="{name()}">
      <xsl:choose>
        <xsl:when test="key('natives', .)[matches(@native-name, 'Kein Absatzformat')]">
          <xsl:sequence select="'None'"/>
        </xsl:when>
        <xsl:when test="key('natives', .)[matches(@native-name, '(Einfaches Absatzformat|^p$)', 'i')]">
          <xsl:sequence select="'para'"/>
        </xsl:when>
        <xsl:when test="matches(., 'hub:identifier')">
          <xsl:sequence select="."/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="key('natives', .)/docx2hub:normalize-to-css-name(@native-name)"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute> 
  </xsl:template>
  
  <xsl:template match="*:keywordset[@role='fieldVars']" mode="docx2hub:join-runs">
    <xsl:if test="exists(//*:keyword[matches(@role,'^fieldVar_')])">
      <xsl:copy>
        <xsl:apply-templates select="@*" mode="#current"/>
        <xsl:apply-templates select="//*:keyword[matches(@role,'^fieldVar_')]" mode="field-var"/>
      </xsl:copy>  
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="*:keyword[matches(@role,'^fieldVar_')]" mode="field-var">
    <xsl:copy>
      <xsl:attribute name="role" select="replace(@role,'^fieldVar_','')"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*:keyword[matches(@role,'^fieldVar_')]" mode="docx2hub:join-runs"/>
  
</xsl:stylesheet>