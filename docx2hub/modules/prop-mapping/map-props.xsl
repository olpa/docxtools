<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"
  xmlns:css = "http://www.w3.org/1996/css"
  xmlns:dbk	= "http://docbook.org/ns/docbook"
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  xmlns:letex = "http://www.le-tex.de/namespace"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:mml = "http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
  xmlns:o = "urn:schemas-microsoft-com:office:office"
  xmlns:pkg	= "http://schemas.microsoft.com/office/2006/xmlPackage"
  xmlns:r	= "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel	= "http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:v	= "urn:schemas-microsoft-com:vml" 
  xmlns:w	= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:word200x = "http://schemas.microsoft.com/office/word/2003/wordml"
  xmlns:wx = "http://schemas.microsoft.com/office/word/2003/auxHint"
  xmlns:w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:extendedProps = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:customProps = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:xs = "http://www.w3.org/2001/XMLSchema"
  xmlns:xsl	= "http://www.w3.org/1999/XSL/Transform"
  xmlns="http://docbook.org/ns/docbook"
  exclude-result-prefixes = "#all">

  <xsl:import href="propmap.xsl"/>
  <xsl:import href="http://transpect.le-tex.de/xslt-util/colors/colors.xsl"/>

  <xsl:key name="docx2hub:style" match="w:style" use="@w:styleId" />
  <xsl:key name="docx2hub:style-by-role" match="css:rule | dbk:style" use="if ($hub-version eq '1.0') then @role else @name" />

  <xsl:template match="/" mode="docx2hub:add-props">
    <xsl:apply-templates select="w:root/w:document/w:body" mode="#current" />
  </xsl:template>
  
  <xsl:template match="@srcpath" mode="docx2hub:add-props">
    <xsl:attribute name="srcpath" select="substring-after(., escape-html-uri(/w:root/@xml:base))" />
  </xsl:template>

  <xsl:template match="w:body" mode="docx2hub:add-props">
    <xsl:element name="{if ($hub-version eq '1.0') then 'Body' else 'hub'}">
      <xsl:attribute name="version" select="concat('5.1-variant le-tex_Hub-', $hub-version)"/>
      <xsl:attribute name="css:version" select="concat('3.0-variant le-tex_Hub-', $hub-version)" />
      <xsl:if test="not($hub-version eq '1.0')">
        <xsl:attribute name="css:rule-selection-attribute" select="'role'" />
      </xsl:if>
      <info>
        <keywordset role="hub">
          <keyword role="formatting-deviations-only">true</keyword>
          <keyword role="source-type">docx</keyword>
          <xsl:if test="/w:root/@xml:base != ''">
            <keyword role="source-dir-uri">
              <xsl:value-of select="/w:root/@xml:base" />
            </keyword>
            <keyword role="archive-dir-uri">
              <xsl:value-of select="concat(replace(/w:root/@xml:base, '^(.+)/[^/]+/?', '$1'), '/')"/>
            </keyword>
          </xsl:if>
          <keyword role="source-basename">
            <!-- /w:root/@xml:base example: file:/data/docx/M_001.docx.tmp/ -->
            <xsl:value-of select="replace(/w:root/@xml:base, '^.*/(.+)\.docx(\.tmp/?)?$', '$1')"/>
          </keyword>
          <xsl:if test="/w:root/w:containerProps/*:Properties/*:Application">
            <keyword role="source-application">
              <xsl:value-of select="/w:root/w:containerProps/*:Properties/*:Application"/>
            </keyword>
          </xsl:if>
        </keywordset>
        <xsl:if test="exists(../../w:settings/w:docVars/w:docVar)">
          <keywordset role="docVars">
            <xsl:for-each select="../../w:settings/w:docVars/w:docVar">
              <keyword role="{@w:name}">
                <xsl:value-of select="@w:val"/>
              </keyword>
            </xsl:for-each>
          </keywordset>
        </xsl:if>
        <xsl:if test="exists(/w:root/w:containerProps/customProps:Properties/customProps:property)">
          <keywordset role="custom-meta">
            <xsl:apply-templates mode="#current" 
              select="/w:root/w:containerProps/customProps:Properties/customProps:property"/>
          </keywordset>
        </xsl:if>
        <xsl:if test="$field-vars='yes'">
          <keywordset role="fieldVars"/>
        </xsl:if>
        <xsl:choose>
          <xsl:when test="$hub-version eq '1.0'">
            <xsl:call-template name="docx2hub:hub-1.0-styles">
              <xsl:with-param name="version" select="$hub-version" tunnel="yes"/>
              <xsl:with-param name="contexts" select="., /w:root/w:footnotes, /w:root/w:endnotes"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$hub-version eq '1.1'">
            <xsl:call-template name="docx2hub:hub-1.1-styles">
              <xsl:with-param name="version" select="$hub-version" tunnel="yes"/>
              <xsl:with-param name="contexts" select="., /w:root/w:footnotes, /w:root/w:endnotes"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </info>
      <xsl:apply-templates select="../../w:numbering" mode="#current"/>
      <xsl:copy-of select="../../w:docRels, ../../w:footnoteRels, ../../w:commentRels, ../../w:fonts"/>
      <xsl:apply-templates select="../../w:comments, ../../w:footnotes, ../../w:endnotes" mode="#current"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="customProps:property" mode="docx2hub:add-props">
    <keyword role="{@name}">
      <xsl:value-of select="*"/>
    </keyword>
  </xsl:template>
  
  <xsl:template match="*[ancestor-or-self::w:numbering]" mode="docx2hub:add-props" priority="-1">
    <xsl:copy>
      <xsl:copy-of select="@*"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template name="docx2hub:hub-1.0-styles">
    <xsl:param name="contexts" as="element(*)+"/>
    <styles>
      <parastyles>
        <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:pStyle/@w:val))
          union
          (: Resolve linked char styles as their respective para styles :)
          key(
           'docx2hub:style',
            key('docx2hub:style', distinct-values($contexts//w:rStyle/@w:val))/w:link/@w:val
          )" mode="#current">
          <xsl:sort select="@w:styleId" />
        </xsl:apply-templates>
      </parastyles>
      <inlinestyles>
        <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:rStyle/@w:val))[not(w:link)]" mode="#current">
          <xsl:sort select="@w:styleId" />
        </xsl:apply-templates>
      </inlinestyles>
      <tablestyles>
        <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:tblStyle/@w:val))" mode="#current">
          <xsl:sort select="@w:styleId" />
        </xsl:apply-templates>
      </tablestyles>
      <cellstyles/>
    </styles>
  </xsl:template>    

  <xsl:template name="docx2hub:hub-1.1-styles">
    <xsl:param name="contexts" as="element(*)+"/>
    <css:rules>
      <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:pStyle/@w:val))
        union
        key('docx2hub:style', key('docx2hub:style', distinct-values($contexts//w:rStyle/@w:val))/w:link/@w:val)" 
        mode="#current">
        <xsl:sort select="@w:styleId" />
      </xsl:apply-templates>
      <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:rStyle/@w:val))" mode="#current">
        <xsl:sort select="@w:styleId" />
      </xsl:apply-templates>
      <xsl:apply-templates select="key('docx2hub:style', distinct-values($contexts//w:tblStyle/@w:val))" mode="#current">
        <xsl:sort select="@w:styleId" />
      </xsl:apply-templates>
    </css:rules>
  </xsl:template>
  
    
  <xsl:template match="w:style" mode="docx2hub:add-props">
    <xsl:param name="wrap-in-style-element" select="true()" as="xs:boolean"/>
    <xsl:param name="version" as="xs:string" tunnel="yes"/>
    <xsl:variable name="atts" as="element(*)*"> <!-- docx2hub:attribute, ... -->
      <xsl:apply-templates select="if (w:basedOn/@w:val) 
                                   then key('docx2hub:style', w:basedOn/@w:val) 
                                   else ()" mode="#current">
        <xsl:with-param name="wrap-in-style-element" select="false()"/>
      </xsl:apply-templates>
      <xsl:variable name="mergeable-atts" as="element(*)*"> <!-- docx2hub:attribute, ... -->
        <xsl:apply-templates select="* except w:basedOn" mode="#current" />
      </xsl:variable>
      <xsl:for-each-group select="$mergeable-atts[self::docx2hub:attribute]" group-by="@name">
        <docx2hub:attribute name="{current-grouping-key()}"><xsl:value-of select="current-group()" /></docx2hub:attribute>
      </xsl:for-each-group>
      <xsl:sequence select="$mergeable-atts[self::*][not(self::docx2hub:attribute)]"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$wrap-in-style-element">
        <xsl:element name="{if ($hub-version = '1.0') then 'style' else 'css:rule'}">
          <xsl:apply-templates select="." mode="docx2hub:XML-Hubformat-add-properties_layout-type"/>
          <xsl:sequence select="$atts"/>
          <xsl:choose>
            <xsl:when test="$hub-version = '1.0'">
              <docx2hub:attribute name="role">
                <xsl:value-of select="docx2hub:css-compatible-name(@w:styleId)"/>
              </docx2hub:attribute>  
            </xsl:when>
            <xsl:otherwise>
              <docx2hub:attribute name="name">
                <xsl:value-of select="docx2hub:css-compatible-name(@w:styleId)"/>
              </docx2hub:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:element>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$atts"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:lvl" mode="docx2hub:add-props">
    <xsl:copy>
      <xsl:apply-templates mode="#current" select="@*"/>
      <xsl:apply-templates mode="#current" select="*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="w:basedOn/@w:val" mode="docx2hub:add-props" />

  <xsl:template match="w:style"
    mode="docx2hub:XML-Hubformat-add-properties_layout-type">
    <xsl:param name="version" tunnel="yes" as="xs:string"/>
    <xsl:if test="not($version eq '1.0')">
      <xsl:attribute name="layout-type" select="if (@w:type = 'paragraph')
        then 'para'
        else if (@w:type = 'character')
        then 'inline'
        else if (@w:type = 'table')
        then 'table'        
        else 'undefined'"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:rPr[not(ancestor::m:oMath)] | w:pPr" mode="docx2hub:add-props" priority="2">
    <xsl:apply-templates select="*" mode="#current" />
  </xsl:template>

  <xsl:template match="w:lvl/w:rPr | w:lvl/w:pPr" mode="docx2hub:add-props" priority="3">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="*" mode="#current" />
      <!-- in order for subsequent (numbering.xsl) symbol mappings, the original rFonts must also be preserved -->
      <xsl:copy-of select="*"/>
    </xsl:copy>
  </xsl:template>
  
  <!-- § 17.3.1.29: this is only for the paragraph mark's formatting: -->
  <xsl:template match="w:pPr/w:rPr" mode="docx2hub:add-props" priority="2.5" />

  <xsl:template match="w:tblPr | w:tblPrEx" mode="docx2hub:add-props" priority="2">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="*" mode="#current" />
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:tblPrEx" mode="docx2hub:add-props">
    <xsl:apply-templates select="* except w:tblW" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:trPr" mode="docx2hub:add-props">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

  <xsl:template match="w:rubyPr" mode="docx2hub:add-props">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

  <xsl:template match="w:tcPr" mode="docx2hub:add-props" priority="2">
    <xsl:apply-templates select="*" mode="#current" />
    <!-- for cellspan etc. processing as defined in tables.xsl: -->
    <xsl:copy-of select="." />
  </xsl:template>

  <xsl:template match="w:u" mode="docx2hub:add-props" priority="2">
    <xsl:next-match/>
    <xsl:apply-templates select="@w:color" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:ind" mode="docx2hub:add-props" priority="2">
    <!-- Precedence as per § 17.3.1.12.
         Ignore the ..Chars variants and start/end for the time being. -->
    <xsl:apply-templates select="@w:left | @w:right | @w:firstLine, @w:hanging" mode="#current" />
  </xsl:template>

  <xsl:template match="w:pBdr | w:tcBorders | w:tblBorders | w:tblCellMar | w:tcMar" mode="docx2hub:add-props" priority="2">
    <xsl:apply-templates select="w:left | w:right | w:top | w:bottom | w:insideH | w:insideV" mode="#current" />
  </xsl:template>

  <xsl:template match="w:spacing" mode="docx2hub:add-props" priority="2">
    <!-- Precedence as per § 17.3.1.33.
         Ignore autosacing and lineRule.
         Handles both w:pPr/w:spacing and w:rPr/w:spacing in this template rule. -->
    <xsl:apply-templates select="@w:val | @w:after | @w:before | @w:line, @w:afterLines | @w:beforeLines" mode="#current" />
  </xsl:template>

  <xsl:template match="w:tcW" mode="docx2hub:add-props" priority="2">
    <xsl:apply-templates select="@w:w" mode="#current" />
  </xsl:template>

  <xsl:template match="w:sectPr[parent::w:pPr]" mode="docx2hub:add-props" priority="+2">
    <xsl:apply-templates select="w:pgSz" mode="#current"/>
  </xsl:template>

  <xsl:template match="  w:style/*
                       | w:rPr[not(ancestor::m:oMath)]/* 
                       | w:pBdr/* 
                       | w:pPr/* 
                       | w:tblPr/*
                       | w:tcBorders/* 
                       | w:tblBorders/*
                       | w:tcPr/*
                       | w:trPr/*[local-name() ne 'tblHeader']
                       | w:rubyPr/*
                       | w:tblPrEx/*
                       | w:tblCellMar/*
                       | w:tcMar/*
                       | w:pgSz/@*
                       | w:ind/@* 
                       | w:tab/@*[local-name() ne 'srcpath']
                       | w:tcW/@* 
                       | w:u/@w:color
                       | w:spacing/@* 
                       | v:shape/@* 
                       " 
    mode="docx2hub:add-props">
    <xsl:variable name="prop" select="key('docx2hub:prop', docx2hub:propkey(.), $docx2hub:propmap)" />
    <xsl:variable name="raw-output" as="element(*)*">
      <xsl:apply-templates select="$prop" mode="#current">
        <xsl:with-param name="val" select="." tunnel="yes" as="item()"/>
      </xsl:apply-templates>
      <xsl:if test="empty($prop)">
        <!-- Fallback (no mapping in propmap): -->
        <docx2hub:attribute name="docx2hub:generated-{local-name()}"><xsl:value-of select="docx2hub:serialize-prop(.)" /></docx2hub:attribute>
      </xsl:if>
    </xsl:variable>
    <xsl:sequence select="$raw-output" />
  </xsl:template>

  <xsl:template match="@w:rsid
                       | @w:rsidDel
                       | @w:rsidR
                       | @w:rsidRPr
                       | @w:rsidRDefault
                       | @w:rsidP
                       | @w:rsidTr
                       | @w14:paraId
                       | @w14:textId"
    mode="docx2hub:add-props" />

  <xsl:function name="docx2hub:propkey" as="xs:string">
    <xsl:param name="prop" as="node()" /> <!-- w:sz, ... -->
    <xsl:choose>
      <xsl:when test="$prop/self::attribute()">
        <xsl:sequence select="string-join((name($prop/..), name($prop)), '/@')" />
      </xsl:when>
      <xsl:when test="$prop/(parent::w:pBdr, parent::w:tcBorders, parent::w:tblBorders, parent::w:tblCellMar, parent::w:tcMar)">
        <xsl:sequence select="string-join((name($prop/..), name($prop)), '/')" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="name($prop)" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="docx2hub:serialize-prop" as="xs:string">
    <xsl:param name="prop" as="item()" /> <!-- w:sz, @fillcolor, ... -->
    <xsl:sequence select="string-join(
                            for $a in 
                              if ($prop instance of element()) 
                              then $prop/@* 
                              else $prop (: attribute() :)
                            return concat(name($a), ':', $a),
                            '; '
                          )" />
  </xsl:function>

  <xsl:template match="prop" mode="docx2hub:add-props" as="node()*">
    <xsl:variable name="atts" as="element(*)*">
      <!-- in the following line, val is a potential child of prop (do not cofuse with $val)! -->
      <xsl:apply-templates select="@type, val, @target-value[not(../(@type, val))]" mode="#current" />
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="empty($atts) and @default">
        <docx2hub:attribute name="{@target-name}"><xsl:value-of select="@default" /></docx2hub:attribute>
      </xsl:when>
      <xsl:when test="empty($atts)" />
      <xsl:otherwise>
        <xsl:sequence select="$atts" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="prop[not(@type | val)]/@target-value" mode="docx2hub:add-props">
    <xsl:call-template name="docx2hub:XML-Hubformat-atts"/>
  </xsl:template>

  <xsl:template match="val" mode="docx2hub:add-props" as="element(*)?">
    <xsl:apply-templates select="@eq, @match" mode="#current" />
  </xsl:template>

  <xsl:template match="prop/@type" mode="docx2hub:add-props" as="node()*">
    <xsl:param name="val" as="item()" tunnel="yes" /><!-- element or attribute -->
    <xsl:choose>

      <xsl:when test=". eq 'percentage'">
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="if (xs:integer($val) eq -1) then 1 else xs:double($val) * 0.01" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'lang'">
        <docx2hub:attribute name="{../@target-name}">
          <!-- provisional -->
          <xsl:value-of select="if (matches($val, 'German') or matches($val, '\Wde\W'))
                                then 'de'
                                else 
                                  if (matches($val, 'English'))
                                  then 'en'
                                  else $val" />
        </docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-boolean-prop'">
        <xsl:choose>
          <xsl:when test="$val/@w:val = ('0','false') and exists(../@default)">
            <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="../@default" /></docx2hub:attribute>
          </xsl:when>
          <xsl:otherwise>
            <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="../@active" /></docx2hub:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test=". eq 'docx-bdr'">
        <xsl:variable name="borders" as="element(w:pBdr)">
          <w:pBdr>
            <xsl:for-each select="('left', 'right', 'bottom', 'top')">
              <xsl:element name="w:{.}">
                <xsl:copy-of select="$val/@*"/>
              </xsl:element>
            </xsl:for-each>
          </w:pBdr>
        </xsl:variable>
        <xsl:apply-templates select="$borders/*" mode="#current"/>
      </xsl:when>
      
      <xsl:when test=". eq 'docx-border'">
        <!-- According to § 17.3.1.5 and other sections, the top/bottom borders don't apply
             if a set of paras has identical border settings. The between setting should be used instead.
             TODO -->
        <xsl:variable name="orientation" select="replace(../@name, '^.+:', '')" as="xs:string"/>
        <docx2hub:attribute name="css:border-{$orientation}-style">
          <xsl:value-of select="docx2hub:border-style($val/@w:val)"/>
        </docx2hub:attribute>
        <xsl:if test="not($val/@w:val = ('nil','none'))">
          <docx2hub:attribute name="css:border-{$orientation}-width">
            <xsl:value-of select="docx2hub:pt-border-size($val/@w:sz)"/>
          </docx2hub:attribute>
          <xsl:if test="$val/@w:color ne 'auto'">
            <docx2hub:attribute name="css:border-{$orientation}-color">
              <xsl:value-of select="docx2hub:color($val/@w:color)"/>
            </docx2hub:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      
      <xsl:when test=". eq 'docx-padding'">
        <xsl:variable name="orientation" select="replace(../@name, '^.+:', '')" as="xs:string" />
        <docx2hub:attribute name="css:padding-{$orientation}"><xsl:value-of select="docx2hub:pt-length($val/@w:w)" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-charstyle'">
        <xsl:variable name="linked" as=" xs:string?" select="key('docx2hub:style', $val/@w:val, root($val))/w:link/@w:val"/>
        <xsl:call-template name="docx2hub:style-name">
          <xsl:with-param name="val" select="$val"/>
          <xsl:with-param name="linked" select="$linked"/>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test=". eq 'docx-color'">
        <xsl:variable name="colorval" as="xs:string?" select="docx2hub:color( ($val/@w:val, $val)[1] )" />
        <xsl:if test="$colorval">
          <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="$colorval" /></docx2hub:attribute>
        </xsl:if>
      </xsl:when>

      <xsl:when test=". eq 'docx-font-family'">
        <xsl:if test="$val/@w:ascii or $val/@w:cs">
          <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="$val/(@w:ascii, @w:cs)[1]" /></docx2hub:attribute>
        </xsl:if>
      </xsl:when>

      <xsl:when test=". eq 'docx-font-size'">
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="concat(number($val/@w:val) * 0.5, 'pt')" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-font-stretch'">
        <xsl:variable name="result" as="xs:string">
          <xsl:choose>
            <xsl:when test="$val/@w:val &lt; 40">
              <xsl:sequence select="'ultra-condensed'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 60">
              <xsl:sequence select="'extra-condensed'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 80">
              <xsl:sequence select="'condensed'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 96">
              <xsl:sequence select="'semi-condensed'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 104">
              <xsl:sequence select="'normal'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 120">
              <xsl:sequence select="'semi-expanded'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 140">
              <xsl:sequence select="'expanded'" />
            </xsl:when>
            <xsl:when test="$val/@w:val &lt; 160">
              <xsl:sequence select="'extra-expanded'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:sequence select="'ultra-expanded'" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="$result" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-hierarchy-level'">
        <xsl:if test="not($val/@w:val eq '9')"><!-- § 17.3.1.20 -->
          <docx2hub:attribute name="remap"><xsl:value-of select="concat('h', xs:integer($val/@w:val) + 1)" /></docx2hub:attribute>
        </xsl:if>
      </xsl:when>

      <xsl:when test=". eq 'docx-length-attr'">
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="docx2hub:pt-length(($val, $val/@w:val, $val/@w:w)[normalize-space()][1])" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-length-attr-negated'">
        <xsl:variable name="string-val" select="string(($val/@w:val, $val)[1])" as="xs:string"/>
        <docx2hub:attribute name="{../@target-name}">
          <xsl:value-of select="if (matches($string-val, '^-'))
                                then docx2hub:pt-length(replace($string-val, '^-', ''))
                                else docx2hub:pt-length(concat('-', $string-val))" />
        </docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'docx-parastyle'">
        <xsl:call-template name="docx2hub:style-name">
          <xsl:with-param name="val" select="$val"/>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test=". eq 'docx-shd'">
        <xsl:choose>
          <xsl:when test="$val/@w:val = ('clear','nil')">
            <xsl:if test="$val/@w:fill ne 'auto'">
              <docx2hub:attribute name="css:background-color"><xsl:value-of select="concat('#', $val/@w:fill)" /></docx2hub:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$val/@w:val eq 'solid'">
            <docx2hub:attribute name="css:background-color"><xsl:value-of select="concat('#', $val/@w:color)" /></docx2hub:attribute>
          </xsl:when>
          <xsl:when test="matches($val/@w:val, '^pct')">
            <xsl:choose>
              <xsl:when test="exists($val/@w:fill) and exists($val/@w:color)">
                <xsl:if test="not(matches($val/@w:color,'auto'))">
                  <docx2hub:attribute name="css:color"><xsl:value-of select="concat('#', $val/@w:color)"/></docx2hub:attribute>
                </xsl:if>
                <docx2hub:color-percentage target="css:background-color" use="css:color" fill="{if ($val/@w:fill='auto') then '#FFFFFF' else concat('#',$val/@w:fill)}"><xsl:value-of select="replace($val/@w:val, '^pct', '')" /></docx2hub:color-percentage>
              </xsl:when>
              <xsl:otherwise>
                <xsl:message>map-props.xsl: w:shd/@w:val='pct*' only implemented for existing @w:fill and @w:color
                <xsl:copy-of select="$val" />
              </xsl:message>
            </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:message>map-props.xsl: w:shd/@w:val other than 'clear', 'nil', 'pct*', and 'solid' not implemented.
            <xsl:copy-of select="$val" />
            </xsl:message>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test=". eq 'docx-table-row-height'">
        <xsl:variable name="pt-length" select="docx2hub:pt-length($val/@w:val)" as="xs:string"/>
        <xsl:choose>
          <xsl:when test="$val/@w:hRule = 'auto'"/>
          <xsl:when test="$val/@w:hRule = 'atLeast' or (not($val/@w:hRule))">
            <docx2hub:attribute name="css:min-height"><xsl:value-of select="$pt-length" /></docx2hub:attribute>    
          </xsl:when>
          <xsl:when test="$val/@w:hRule = 'exact' 
                          (:or 
                          (
                            not($val/@w:hRule)
                            and
                            $val/ancestor::w:tbl[1]/w:tblPr/w:tblLayout/@w:type = 'fixed'
                          ) :)">
            <!-- not sure about the last condition. § 17.4.81 says:
              “If [@w:hRule] is omitted, then its value shall be assumed to be auto.”
              But there were table rows with a trHeight that lacked @w:hRule, and their height was fixed.
              They were in a fixed-layout table, so we assume that row heights should be respected (at least) in
              fixed tables even if their @w:hRule is missing. -->
            <docx2hub:attribute name="css:height"><xsl:value-of select="$pt-length" /></docx2hub:attribute>    
          </xsl:when>
        </xsl:choose>
        
      </xsl:when>
      
      <xsl:when test=". eq 'docx-underline'">
        <!-- §§§ TODO -->
        <docx2hub:attribute name="css:text-decoration-line"><xsl:value-of select="'underline'" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'tablist'">
        <tabs>
          <xsl:apply-templates select="$val/*" mode="#current"/>
        </tabs>
      </xsl:when>

      <xsl:when test=". eq 'docx-text-direction'">
        <!-- I find 17.18.93 ST_TextDirection remarkably unclear about this. What also bothers me is that
             the value 'btLr' doesn’t appear in the table in that section. In Annex N.1 on p. 5563, they mention
             that btLr et al. have been dropped in Wordprocessing ML. -->
        <xsl:choose>
          <xsl:when test="$val/@w:val = ('tbLr', 'btLr')">
            <!-- preliminary value – only works in IE, while the CSS3 writing mode prop values don’t work -->
            <docx2hub:attribute name="css:writing-mode">bt-lr</docx2hub:attribute>
          </xsl:when>
          <xsl:when test="matches($val/@w:val, 'tb', 'i')">
            <!-- looks funny -->
            <docx2hub:attribute name="css:transform">rotate(90deg)</docx2hub:attribute>
          </xsl:when>
          <xsl:when test="matches($val/@w:val, 'bt', 'i')">
            <!-- looks funny -->
            <docx2hub:attribute name="css:transform">rotate(-90deg)</docx2hub:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:message>Unsupported text direction property <xsl:sequence select="$val"/></xsl:message>
          </xsl:otherwise>
        </xsl:choose>
        <!-- no effect: -->
        <!--<docx2hub:attribute name="css:width">fit-content</docx2hub:attribute>-->
      </xsl:when>
      
      <xsl:when test=". eq 'linear'">
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="if ($val/self::attribute())
                                                                           then $val
                                                                           else $val/@w:val" /></docx2hub:attribute>
      </xsl:when>

      <xsl:when test=". eq 'passthru'">
        <xsl:copy-of select="$val" copy-namespaces="no"/>
      </xsl:when>

      <xsl:when test=". eq 'docx-position'">
        <xsl:choose>
          <xsl:when test="$val/@w:val eq 'baseline'" />
          <xsl:otherwise>
            <docx2hub:wrap element="{$val/@w:val}" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test=". eq 'style-link'">
        <docx2hub:style-link type="{../@name}" target="{$val}"/>
      </xsl:when>

      <xsl:when test=". eq 'style-name'">
        <docx2hub:attribute name="{../@target-name}">
          <xsl:value-of select="docx2hub:css-compatible-name($val/@w:val)"/>
        </docx2hub:attribute>
      </xsl:when>

      <xsl:otherwise>
        <docx2hub:attribute name="{../@target-name}"><xsl:value-of select="$val" /></docx2hub:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- ISO/IEC 29500-1, 17.3.2.40: 
    This color therefore can be automatically be modified by a consumer as appropriate, 
    for example, in order to ensure that the underline can be distinguished against the 
    page's background color. --> 
  <xsl:variable name="docx2hub:auto-color" select="'black'" as="xs:string"/>

  <xsl:function name="docx2hub:color" as="xs:string?" >
    <xsl:param name="val" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="$val = 'none'" />
      <xsl:when test="$val = 'auto'" >
        <!-- not tested yet whether this interferes with <w:shd w:fill="…"/> -->
        <xsl:sequence select="$docx2hub:auto-color" />
      </xsl:when>
      <xsl:when test="matches($val, '^#[0-9a-f]{6}$')">
        <!-- e.g., v:shape/@fillcolor -->
        <xsl:sequence select="upper-case($val)" />
      </xsl:when>
      <xsl:when test="matches($val, '[0-9A-F]{6}')">
        <xsl:sequence select="concat('#', $val)" />
      </xsl:when>
      <xsl:when test="$val = ('black', 'blue', 'red', 'white', 'yellow')">
        <xsl:sequence select="$val" />
      </xsl:when>
      <xsl:when test="$val eq 'cyan'">
        <xsl:sequence select="'#00FFFF'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkBlue'">
        <xsl:sequence select="'#00008B'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkCyan'">
        <xsl:sequence select="'#008B8B'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkGray'">
        <xsl:sequence select="'#A9A9A9'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkGreen'">
        <xsl:sequence select="'#006400'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkMagenta'">
        <xsl:sequence select="'#800080'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkRed'">
        <xsl:sequence select="'#8B0000'" />
      </xsl:when>
      <xsl:when test="$val eq 'darkYellow'">
        <xsl:sequence select="'#808000'" />
      </xsl:when>
      <xsl:when test="$val eq 'green'">
        <xsl:sequence select="'#00FF00'" />
      </xsl:when>
      <xsl:when test="$val eq 'lightGray'">
        <xsl:sequence select="'#D3D3D3'" />
      </xsl:when>
      <xsl:when test="$val eq 'magenta'">
        <xsl:sequence select="'#FF00FF'" />
      </xsl:when>
      <xsl:otherwise><!-- shouldn't happen -->
        <xsl:sequence select="''" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template name="docx2hub:style-name" as="element(docx2hub:attribute)">
    <xsl:param name="val" as="element(*)"/><!-- w:pStyle, w:cStyle -->
    <xsl:param name="linked" as="xs:string?"/>
    <!-- we choose not to use the linked (paragraph) style here because we’d have to
      carefully select only the css:rule’s inline properties when recreating docx run properties
      from hub. -->
    <xsl:variable name="looked-up" as="xs:string" select="$val/@w:val" />
    <!--<xsl:variable name="looked-up" as="xs:string" select="if ($linked) then $linked else $val/@w:val" />-->
    <docx2hub:attribute name="role">
      <xsl:value-of select="if ($hub-version eq '1.0')
                              then $looked-up
                              else docx2hub:css-compatible-name($looked-up)" />
    </docx2hub:attribute>
  </xsl:template>

  <xsl:function name="docx2hub:pt-length" as="xs:string" >
    <xsl:param name="val" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="not($val)">
        <xsl:message>empty argument for docx2hub:pt-length, defaulting to zero. </xsl:message>
        <xsl:sequence select="'0'"/>
      </xsl:when>
      <xsl:when test="not($val castable as xs:integer)">
        <xsl:message>argument '<xsl:value-of select="$val"/>' for docx2hub:pt-length not castable as xs:integer, defaulting to zero. </xsl:message>
        <xsl:sequence select="'0'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="if (matches($val, '%$'))
          then $val
          else concat(xs:string(xs:integer($val) * 0.05), 'pt')" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="docx2hub:pt-border-size" as="xs:string" >
    <xsl:param name="val" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="not($val)">
        <xsl:message>empty argument for docx2hub:pt-border-size, defaulting to zero. </xsl:message>
        <xsl:sequence select="'0'"/>
      </xsl:when>
      <xsl:when test="not($val castable as xs:integer)">
        <xsl:message>argument '<xsl:value-of select="$val"/>' for docx2hub:border-size not castable as xs:integer, defaulting to zero. </xsl:message>
        <xsl:sequence select="'0'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="if (matches($val, '%$'))
          then $val
          else concat(xs:string(xs:integer($val) * 0.125), 'pt')" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:function name="docx2hub:border-style" as="xs:string" >
    <xsl:param name="val" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="$val eq 'single'">
        <xsl:sequence select="'solid'" />
      </xsl:when>
      <xsl:when test="$val eq 'nil'">
        <xsl:sequence select="'none'" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$val" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="val/@match" mode="docx2hub:add-props" as="element(*)?">
    <xsl:param name="val" as="item()" tunnel="yes" />
    <xsl:if test="matches($val/@w:val, .)">
      <xsl:call-template name="docx2hub:XML-Hubformat-atts" />
    </xsl:if>
  </xsl:template>

  <xsl:template match="val/@eq" mode="docx2hub:add-props" as="element(*)?">
    <xsl:param name="val" as="item()" tunnel="yes" />
    <xsl:if test="string($val) = string(.) or (string(.) = 'true' and empty($val/@*))">
      <xsl:call-template name="docx2hub:XML-Hubformat-atts" />
    </xsl:if>
  </xsl:template>

  <xsl:template name="docx2hub:XML-Hubformat-atts" as="element(*)?">
    <xsl:variable name="target-val" select="(../@target-value, ../../@target-value)[last()]" as="xs:string?" />
    <xsl:if test="exists($target-val)">
      <docx2hub:attribute name="{(../@target-name, ../../@target-name)[last()]}"><xsl:value-of select="$target-val" /></docx2hub:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template match="Color" mode="docx2hub:add-props" as="xs:string">
    <xsl:param name="multiplier" as="xs:double" select="1.0" />
    <xsl:choose>
      <xsl:when test="@Space eq 'CMYK'">
        <xsl:sequence select="concat(
                                'device-cmyk(', 
                                string-join(
                                  for $v in tokenize(@ColorValue, '\s') return xs:string(xs:integer(xs:double($v) * 10000 * $multiplier) * 0.000001)
                                  , ','
                                ),
                                ')'
                              )" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:message>Unknown colorspace <xsl:value-of select="@Space"/>
        </xsl:message>
        <xsl:sequence select="@ColorValue" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:tabs/w:tab" mode="docx2hub:add-props" as="element(dbk:tab)" priority="2">
    <tab>
      <xsl:apply-templates select="@*" mode="#current" />
    </tab>
  </xsl:template>


  <xsl:key name="docx2hub:style" 
    match="CellStyle | CharacterStyle | ObjectStyle | ParagraphStyle | TableStyle" 
    use="@Self" />

  <xsl:function name="docx2hub:css-compatible-name" as="xs:string">
    <xsl:param name="input" as="xs:string"/>
    <xsl:sequence select="replace(replace(normalize-unicode($input, 'NFKD'), '\p{Mn}', ''), '[^-_a-z0-9]', '_', 'i')"/>
  </xsl:function>

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- mode: docx2hub:props2atts -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <xsl:template match="* | @*" mode="docx2hub:props2atts">
    <xsl:variable name="content" as="node()*">
      <xsl:apply-templates select="docx2hub:style-link" mode="#current" />
      <xsl:apply-templates select="docx2hub:attribute[not(@name = following-sibling::docx2hub:remove-attribute/@name)]" mode="#current" />
      <xsl:apply-templates select="docx2hub:color-percentage[not(@name = following-sibling::docx2hub:remove-attribute/@name)]" mode="#current" />
      <xsl:variable name="remaining-tabs" as="element(dbk:tab)*">
        <xsl:for-each-group select="dbk:tabs/dbk:tab" group-by="docx2hub:attribute[@name eq 'horizontal-position']">
          <xsl:if test="not(current-group()[last()]/docx2hub:attribute[@name eq 'clear'])">
            <xsl:apply-templates select="current-group()[last()]" mode="#current"/>
          </xsl:if>
        </xsl:for-each-group>
      </xsl:variable>
      <xsl:if test="exists($remaining-tabs)">
        <tabs>
          <xsl:sequence select="$remaining-tabs" />
        </tabs>
      </xsl:if>
      <xsl:apply-templates select="node() except (docx2hub:attribute | docx2hub:color-percentage | docx2hub:wrap | docx2hub:style-link | dbk:tabs)" mode="#current" />
    </xsl:variable>
    <xsl:choose>
      <!-- do not wrap whitespace only subscript or superscript -->
      <xsl:when test="w:t and docx2hub:wrap/@element = ('superscript', 'subscript') 
                      and exists(docx2hub:wrap/@element[ . ne 'superscript' and . ne 'subscript'])
                      and (
                        every $el in $content[self::*] 
                        satisfies $el[self::w:t[@xml:space eq 'preserve'][matches(., '^\p{Zs}*$')]]
                      )">
        <xsl:sequence select="docx2hub:wrap($content, (docx2hub:wrap[not(@element = ('superscript', 'subscript'))]))" />
      </xsl:when>
      <!-- do not wrap whitespace only subscript or superscript -->
      <xsl:when test="w:t and docx2hub:wrap/@element = ('superscript', 'subscript') 
                      and not(exists(docx2hub:wrap/@element[ . ne 'superscript' and . ne 'subscript']))
                      and (
                        every $el in $content[self::*] 
                        satisfies $el[self::w:t[@xml:space eq 'preserve'][matches(., '^\p{Zs}*$')]]
                      )">
        <xsl:copy>
          <xsl:sequence select="docx2hub:wrap($content, (docx2hub:wrap[not(@element = ('superscript', 'subscript'))]))" />
        </xsl:copy>
      </xsl:when>
      <xsl:when test="exists(docx2hub:wrap) and exists(self::css:rule | self::dbk:style)">
        <xsl:copy>
          <xsl:attribute name="remap" select="docx2hub:wrap/@element" />
          <xsl:sequence select="@*, $content" />
        </xsl:copy>
      </xsl:when>
      <xsl:when test="exists(docx2hub:wrap) and not(self::css:rule or self::dbk:style)">
        <xsl:sequence select="docx2hub:wrap($content, (docx2hub:wrap))" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:copy>
          <xsl:sequence select="@*, $content" />
        </xsl:copy>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="docx2hub:wrap" as="node()*">
    <xsl:param name="content" as="node()*" />
    <xsl:param name="wrappers" as="element(docx2hub:wrap)*" />
    <xsl:choose>
      <xsl:when test="exists($wrappers)">
        <xsl:element name="{$wrappers[1]/@element}">
          <xsl:sequence select="docx2hub:wrap($content, $wrappers[position() gt 1])" />
        </xsl:element>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="$content" mode="docx2hub:props2atts"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="docx2hub:attribute" mode="docx2hub:props2atts">
    <xsl:attribute name="{@name}" select="." />
  </xsl:template>

  <xsl:template match="docx2hub:attribute[@name = ('fill-tint')]" mode="docx2hub:props2atts"/>
  
  <xsl:template match="docx2hub:attribute[@name = 'css:text-decoration-line']" mode="docx2hub:props2atts">
    <xsl:variable name="all-atts" select="preceding-sibling::docx2hub:attribute[@name = current()/@name], ."
      as="element(docx2hub:attribute)+"/>
    <xsl:variable name="tokenized" select="for $a in $all-atts return tokenize($a, '\s+')" as="xs:string+"/>
    <xsl:variable name="line-through" select="$tokenized[starts-with(., 'line-through')][last()]"/>
    <xsl:variable name="underline" select="$tokenized[starts-with(., 'underline')][last()]"/>
    <xsl:choose>
      <xsl:when test="every $t in ($line-through, $underline) satisfies (ends-with($t, 'none'))">
        <xsl:attribute name="{@name}" select="'none'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="{@name}" select="($line-through, $underline)[not(ends-with(., 'none'))]"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="docx2hub:color-percentage" mode="docx2hub:props2atts">
    <xsl:variable name="color" select="(../docx2hub:attribute[@name eq 'css:color'][last()], '#000000')[1]" as="xs:string" />
    <xsl:attribute name="{@target}" select="letex:tint-hex-color-filled($color, number(.) * 0.01, @fill)" />
  </xsl:template>

  <!-- Let only the last setting prevail: -->
  <xsl:template match="w:numPr[following-sibling::w:numPr]" mode="docx2hub:props2atts" />
  <xsl:template match="w:tblPr[following-sibling::w:tblPr]" mode="docx2hub:props2atts" />
  <xsl:template match="w:tcPr[following-sibling::w:tcPr]" mode="docx2hub:props2atts" />

  <xsl:template match="w:numPr[not(following-sibling::w:numPr)][not(w:numId)]" mode="docx2hub:props2atts">
    <w:numPr>
      <xsl:apply-templates select="(preceding-sibling::w:numPr/w:numId)[1]" mode="#current"/>
      <xsl:apply-templates select="node()" mode="#current"/>
    </w:numPr>
  </xsl:template>
   
  <xsl:template match="docx2hub:remove-attribute" mode="docx2hub:props2atts" />

  <xsl:template match="docx2hub:style-link" mode="docx2hub:props2atts">
    <xsl:attribute name="{if (@type eq 'AppliedParagraphStyle')
                          then 'parastyle'
                          else @type}" 
      select="@target" />
  </xsl:template>

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- mode: docx2hub:remove-redundant-run-atts -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <xsl:template match="w:r/@*[some $att in ../ancestor::w:p[1]/@* 
                       satisfies (
                         name($att) eq name(current())
                         and 
                         xs:string($att) eq xs:string(current())
                       )]" mode="docx2hub:remove-redundant-run-atts" />

  <xsl:template match="*[w:p[w:pgSz]]" mode="docx2hub:remove-redundant-run-atts">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="node()" group-ending-with="w:p[w:pgSz[@css:orientation='landscape']]">
        <xsl:choose>
          <xsl:when test="current-group()[last()][self::w:p[w:pgSz[@css:orientation='landscape']]]">
            <xsl:for-each-group select="current-group()" group-starting-with="w:p[w:pgSz[not(@css:orientation='landscape')]]">
              <xsl:choose>
                <xsl:when test="current-group()[last()][self::w:p[w:pgSz[@css:orientation='landscape']]]">
                  <xsl:apply-templates select="current-group()[1]" mode="#current"/>
                  <xsl:apply-templates select="current-group()[position() gt 1]" mode="add-attribute"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:apply-templates select="current-group()" mode="#current"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each-group>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="*" mode="add-attribute">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates mode="docx2hub:remove-redundant-run-atts"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:tbl | w:p[descendant::v:imagedata]" mode="add-attribute">
    <xsl:copy>
      <xsl:attribute name="css:orientation" select="'landscape'"/>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:apply-templates mode="docx2hub:remove-redundant-run-atts"/>
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>