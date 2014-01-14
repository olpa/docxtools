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
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns="http://docbook.org/ns/docbook"

  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn letex mml docx2hub"
  >

  <xsl:template match="w:tbl" mode="wml-to-dbk">
    <informaltable>
      <xsl:apply-templates select="w:tblPr/@role|w:tblPr/@css:*[matches(name(.),'border(\-.+)?\-style')]" mode="#current" />
<!--      <xsl:attribute name="frame" select="if (some $i in (w:tblPr/@css:*[matches(name(.),'border\-.+\-style')]) satisfies (exists($i) and not($i=('nil','none')))) then 'all' else 'none'"/>-->
      <tgroup>
        <xsl:attribute name="cols">
          <xsl:value-of select="count(w:tblGrid/w:gridCol)"/>
        </xsl:attribute>
        <xsl:if test="w:tblPr/w:tblW/@w:w">
<!--           <xsl:attribute name="width" select="w:tblPr/w:tblW/@w:w"/> -->
        </xsl:if>
        <xsl:apply-templates select="w:tblGrid" mode="colspec">
          <xsl:with-param name="width" 
            select="if (not(w:tblPr/w:tblW) or w:tblPr/w:tblW/@w:type = 'pct') then 0 else w:tblPr/w:tblW/@w:w" tunnel="yes"/>
        </xsl:apply-templates>
        <xsl:if test="w:tr[w:trPr/w:tblHeader]">
          <thead>
            <xsl:apply-templates select="w:tr[w:trPr/w:tblHeader]" mode="tables">
              <xsl:with-param name="cols" select="count(w:tblGrid/w:gridCol)" tunnel="yes"/>
              <xsl:with-param name="width" select="w:tblPr/w:tblW/@w:w" tunnel="yes"/>
              <xsl:with-param name="col-widths" select="(for $x in w:tblGrid/w:gridCol return $x/@w:w)" tunnel="yes"/>
            </xsl:apply-templates>
          </thead>
        </xsl:if>
        <tbody>
          <xsl:apply-templates select="* except w:tr[w:trPr/w:tblHeader]" mode="tables">
            <xsl:with-param name="cols" select="count(w:tblGrid/w:gridCol)" tunnel="yes"/>
            <xsl:with-param name="width" select="w:tblPr/w:tblW/@w:w" tunnel="yes"/>
            <xsl:with-param name="col-widths" select="(for $x in w:tblGrid/w:gridCol return $x/@w:w)" tunnel="yes"/>
          </xsl:apply-templates>
        </tbody>
      </tgroup>
    </informaltable>
  </xsl:template>

  <xsl:template match="w:tblGridChange" mode="colspec"/>

  <xsl:template match="w:tblGrid" mode="colspec">
    <xsl:param name="width" tunnel="yes" as="xs:integer"/>
    <xsl:if test="not(sum(w:gridCol/@w:w) = $width) and not($width = 0)">
      <xsl:call-template name="signal-error" xmlns="">
        <xsl:with-param name="error-code" select="'W2D_052'"/>
        <xsl:with-param name="exit" select="'no'"/>
        <xsl:with-param name="hash">
          <value key="xpath"><xsl:value-of select="@srcpath"/></value>
          <value key="level">WRN</value>
          <value key="comment"/>
          <value key="info-text"><xsl:value-of select="concat('Table-width: ', $width, '    Sum of col-widths: ', sum(w:gridCol/@w:w))"/></value>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
    <xsl:apply-templates mode="colspec"/>
  </xsl:template>
     
  <xsl:template match="w:tblGrid/w:gridCol" mode="colspec">
    <xsl:variable name="pos" select="index-of(for $elt in ../* return generate-id($elt), generate-id())" as="xs:integer"/>
    <colspec>
      <xsl:attribute name="colnum">
        <xsl:value-of select="$pos"/>
      </xsl:attribute>
      <xsl:attribute name="colname">
        <xsl:value-of select="concat('col', $pos)"/>
      </xsl:attribute>
      <xsl:if test="@w:w != ''">
        <xsl:attribute name="colwidth" select="docx2hub:twips2mm(@w:w)"/>
      </xsl:if>
    </colspec>
  </xsl:template>

  <xsl:template match="w:tblGrid" mode="tables"/>

  <xsl:template match="w:tblPr" mode="tables"/>

  <xsl:template match="w:bookmarkStart | w:bookmarkEnd" mode="tables">
    <!-- das sollte noch mal gecheckt werden -->
  </xsl:template>

  <xsl:template match="*" mode="tables">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_020'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">tables</value>
        <value key="info-text"><xsl:value-of select="concat('Element: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="@*" mode="tables">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_021'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="../@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">tables</value>
        <value key="info-text"><xsl:value-of select="concat('Element: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="processing-instruction()" mode="tables">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_023'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">tables</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="comment()" mode="tables">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_022'"/>
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">tables</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="w:tr" mode="tables">
    <row>
      <xsl:apply-templates select="@css:*, w:tblPrEx/@css:background-color, @xml:lang, @srcpath" mode="wml-to-dbk"/>
      <xsl:apply-templates select="@w:fill-cells-before" mode="wml-to-dbk"/>
      <xsl:apply-templates select="w:tc" mode="#current">
        <xsl:with-param name="row-overrides" as="attribute(*)*" select="w:tblPrEx/(@* except @css:background-color)"/>
        <xsl:with-param name="is-first-row-in-group" select="letex:node-index-of(../w:tr, .) = 1"  tunnel="yes"/>
        <xsl:with-param name="is-last-row-in-group" select="letex:node-index-of(../w:tr, .) = count(../w:tr)"  tunnel="yes"/>
      </xsl:apply-templates>
      <xsl:apply-templates select="@w:fill-cells-after" mode="wml-to-dbk"/>
    </row>
  </xsl:template>
  
  <xsl:template match="@w:fill-width-before | @w:fill-width-after" mode="wml-to-dbk" priority="10"/>
  
  <xsl:template match="@w:fill-cells-before" mode="wml-to-dbk" priority="10">
    <xsl:param name="col-widths" as="xs:integer+" tunnel="yes"/>
    <xsl:for-each select="1 to xs:integer(.)">
      <entry role="hub:fill-grid" css:width="{docx2hub:twips2mm($col-widths[current()])}">
        <para/>
      </entry>
    </xsl:for-each>  
  </xsl:template>
  
  <xsl:template match="@w:fill-cells-after" mode="wml-to-dbk" priority="10">
    <xsl:param name="col-widths" as="xs:integer+" tunnel="yes"/>
    <xsl:param name="cols" as="xs:integer" tunnel="yes"/>
    <xsl:for-each select="($cols - xs:integer(.) + 1) to $cols">
      <entry role="hub:fill-grid" css:width="{docx2hub:twips2mm($col-widths[current()])}">
        <para/>
      </entry>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="w:trPr" mode="tables"/>

  <xsl:template match="w:tcPr" mode="tables"/>

  <xsl:template match="w:tblPrEx" mode="tables"/>

  <xsl:template match="w:p" mode="tables">
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>

  <xsl:template match="w:tc[not(docx2hub:is-blind-vmerged-cell(.))]" mode="tables">
    <xsl:param name="is-first-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="col-widths" tunnel="yes" as="xs:integer*"/>
    <xsl:param name="row-overrides" as="attribute(*)*"/>
    <xsl:element name="entry">
      <xsl:copy-of select="ancestor::w:tbl/w:tblPr/@css:* except ancestor::w:tbl/w:tblPr/@css:width"/>
      <xsl:copy-of select="ancestor::w:tr/@css:* except ancestor::w:tr/@css:width"/>
      <xsl:apply-templates select="$row-overrides" mode="wml-to-dbk">
        <xsl:with-param name="is-first-cell" select="letex:node-index-of(../w:tc, .) = 1"/>
        <xsl:with-param name="is-last-cell" select="letex:node-index-of(../w:tc, .) = count(../w:tc)"/>
      </xsl:apply-templates>
      <xsl:copy-of select="@*" />
      <!-- Process any spans -->
      <xsl:call-template name="cell.span"/>
      <xsl:call-template name="cell.morerows"/>
      <!-- will be treated by map-props.xsl:
      <xsl:call-template name="cell.style"/>
      <xsl:call-template name="cell.align"/>
      -->
      <xsl:call-template name="check-field-functions">
        <xsl:with-param name="nodes" select="node()"/>
      </xsl:call-template>
    </xsl:element>
  </xsl:template>

  <xsl:template match="@css:border-insideV-style | @css:border-insideV-width | @css:border-insideV-color" mode="wml-to-dbk" priority="10">
    <xsl:param name="is-first-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="is-last-row-in-group" as="xs:boolean?" tunnel="yes"/>
    <xsl:if test="not($is-first-row-in-group)">
      <xsl:attribute name="{replace(name(), 'insideV', 'top')}" select="."/>
    </xsl:if>
    <xsl:if test="not($is-last-row-in-group)">
      <xsl:attribute name="{replace(name(), 'insideV', 'bottom')}" select="."/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="@css:border-insideH-style | @css:border-insideH-width | @css:border-insideH-color" mode="wml-to-dbk" priority="10">
    <xsl:param name="is-first-cell" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="is-last-cell" as="xs:boolean?" tunnel="yes"/>
    <xsl:if test="not($is-first-cell)">
      <xsl:attribute name="{replace(name(), 'insideH', 'left')}" select="."/>
    </xsl:if>
    <xsl:if test="not($is-last-cell)">
      <xsl:attribute name="{replace(name(), 'insideH', 'right')}" select="."/>
    </xsl:if>
  </xsl:template>

  <xsl:template name="cell.align">
    <!-- horizontale Ausrichtung des ersten Absatzes -->
    <xsl:variable name="align" select="if (exists(w:p/w:pPr/w:jc/@w:val)) then (w:p/w:pPr/w:jc/@w:val)[1] else 'left'"/>
    <!-- vertikale Ausrichtung -->
    <xsl:variable name="v-align" select="w:tcPr/w:vAlign/@w:val"/>
    <xsl:choose>
      <xsl:when test="$align = 'left'">
        <xsl:attribute name="css:text-align" select="'left'"/>
      </xsl:when>
      <xsl:when test="$align = 'right'">
        <xsl:attribute name="css:text-align" select="'right'"/>
      </xsl:when>
      <xsl:when test="$align = 'both'">
        <xsl:attribute name="css:text-align" select="'justify'"/>
      </xsl:when>
      <xsl:when test="$align = 'center'">
        <xsl:attribute name="css:text-align" select="'center'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="css:text-align" select="'left'"/>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="$v-align = ('both', 'center')">
        <xsl:attribute name="css:vertical-align" select="'middle'"/>        
      </xsl:when>
      <xsl:when test="$v-align = 'top'">
        <xsl:attribute name="css:vertical-align" select="'top'"/>        
      </xsl:when>
      <xsl:when test="$v-align = 'bottom'">
        <xsl:attribute name="css:vertical-align" select="'bottom'"/>        
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="css:vertical-align" select="'top'"/>        
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="cell.style">
    <!-- §§§
    <xsl:attribute name="width" select="w:tcPr/w:tcW/@w:w"/>
    <xsl:if test="exists(w:tcPr/w:shd/@w:fill)">
      <xsl:attribute name="background-color" select="w:tcPr/w:shd/@w:fill"/>
    </xsl:if>

    <xsl:variable name="cell-borders" select="w:tcPr/w:tcBorders"/>
    <!- - Does ../w:tblPrEx/w:tblBorders/w:tblBorders select anything?
         Shouldn't it be ../w:tblPrEx/w:tblBorders instead? - ->
    <xsl:variable name="row-borders" select="../w:tblPrEx/w:tblBorders/w:tblBorders"/>
    <xsl:variable name="table-borders" select="../../w:tblPr/w:tblBorders"/>

    <xsl:variable name="is-most-left" select="if (not(preceding-sibling::w:tc)) then true() else false()"/>
    <xsl:variable name="is-most-right" select="if (not(following-sibling::w:tc)) then true() else false()"/>
    <xsl:variable name="is-most-bottom" select="if (../following-sibling::w:tr) then false() else true()"/>
    <xsl:variable name="is-most-top" select="if (../preceding-sibling::w:tr) then false() else true()"/>

    <xsl:variable name="border-types" select="('top', 'bottom', 'left', 'right')"/>

    <xsl:variable name="unified-borders" select="letex:unify-properties(($row-borders, $table-borders))"/>

    <xsl:if test="$cell-borders/w:insideH or $cell-borders/w:insideV">
      <xsl:call-template name="signal-error" xmlns="">
        <xsl:with-param name="error-code" select="'W2D_050'"/>
        <xsl:with-param name="exit" select="'no'"/>
        <xsl:with-param name="hash">
          <value key="xpath"><xsl:value-of select="$cell-borders[w:insideH or w:insideV][1]/@srcpath"/></value>
          <value key="level">INT</value>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>

    <!- - Morerows können zu Fehlern führen. Die border-bottom einer morerow Zelle wird durch die letzte Zelle der morerow (ausgeblendet durch w:vMerge) definiert und nicht durch die erste - ->

    <xsl:variable name="borders">
      <w:tblBorders>
        <xsl:for-each select="$cell-borders">
          <xsl:copy-of select="node()[local-name() = $border-types]"/>
        </xsl:for-each>
        <xsl:for-each select="$border-types[not(. = $cell-borders/node()/local-name())]">
          <xsl:choose>
            <xsl:when test=". = 'top'">
              <xsl:sequence select="letex:check-border-hierarchy($unified-borders, 'w:top', $is-most-top, 'w:insideH')"/>
            </xsl:when>
            <xsl:when test=". = 'bottom'">
              <xsl:sequence select="letex:check-border-hierarchy($unified-borders, 'w:bottom', $is-most-bottom, 'w:insideH')"/>
            </xsl:when>
            <xsl:when test=". = 'left'">
              <xsl:sequence select="letex:check-border-hierarchy($unified-borders, 'w:left', $is-most-left, 'w:insideV')"/>
            </xsl:when>
            <xsl:when test=". = 'right'">
              <xsl:sequence select="letex:check-border-hierarchy($unified-borders, 'w:right', $is-most-right, 'w:insideV')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="signal-error" xmlns="">
                <xsl:with-param name="error-code" select="'W2D_051'"/>
                <xsl:with-param name="exit" select="'no'"/>
                <xsl:with-param name="hash">
                  <value key="xpath"><xsl:value-of select="$cell-borders[w:insideH or w:insideV][1]/@srcpath"/></value>
                  <value key="level">INT</value>
                  <value key="info-text"><xsl:value-of select="."/></value>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </w:tblBorders>
    </xsl:variable>

    <xsl:if test="not(empty($borders/node()/node()))">
      <xsl:attribute name="style" select="letex:write-table-style($borders/node())"/>
    </xsl:if>

-->
  </xsl:template>

  <xsl:function name="letex:check-border-hierarchy" as="node()?">
    <xsl:param name="borders" as="node()?"/>
    <xsl:param name="name" as="xs:string"/>
    <xsl:param name="flag" as="xs:boolean"/>
    <xsl:param name="alternative" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="$flag">
        <xsl:copy-of select="$borders/*[name() = $name]"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$borders/*[name() = $alternative]">
            <xsl:element name="{$name}">
              <xsl:copy-of select="$borders/*[name() = $alternative]/@*"/>
            </xsl:element>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy-of select="$borders/*[name() = $name]"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>


  <xsl:function name="letex:write-table-style" as="xs:string">
    <xsl:param name="border" as="node()"/>
    <xsl:sequence select="string-join(for $x in $border/node() return
                          (for $y in $x/@*[name() = ('w:sz', 'w:val', 'w:color')]
                          return concat('border-', local-name($x), letex:translate-border-attribs($y))), '; ')"/>
  </xsl:function>

  <xsl:function name="letex:translate-border-attribs" as="xs:string">
    <xsl:param name="attrib" as="attribute()"/>
    <xsl:choose>
      <xsl:when test="name($attrib) = 'w:sz'">
        <xsl:value-of select="concat('-width: ', $attrib div 8, 'pt')"/>
      </xsl:when>
      <xsl:when test="name($attrib) = 'w:val'">
        <xsl:value-of select="concat('-style: ', letex:translate-border-style($attrib))"/>
      </xsl:when>
      <xsl:when test="name($attrib) = 'w:color'">
        <xsl:value-of select="concat('-color: ', $attrib)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="''"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="letex:translate-border-style" as="xs:string">
    <xsl:param name="text" as="xs:string"/>
    <xsl:choose>
      <xsl:when test="$text = 'nil'">none</xsl:when>
      <xsl:when test="$text = 'single'">solid</xsl:when>
      <xsl:otherwise>
        <!--
        <xsl:message>ACHTUNG: border-style kann nicht bestimmt werden: <xsl:value-of select="$text"/></xsl:message>
        -->
        <xsl:value-of select="'solid'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>


  <xsl:template match="w:tc[docx2hub:is-blind-vmerged-cell(.)]" mode="tables"/>

  <xsl:template name="cell.span">
    <xsl:variable name="span" select="0 + w:tcPr/w:gridSpan/@w:val
                                      + (if (not(following-sibling::w:tc) and ../w:trPr/w:gridAfter/@w:val) then ../w:trPr/w:gridAfter/@w:val else 0)
                                      + (if (not(preceding-sibling::w:tc) and ../w:trPr/w:gridBefore/@w:val) then ../w:trPr/w:gridBefore/@w:val else 0)"/>
    <xsl:variable name="colstart" select="letex:colcount(1, .) + (if (preceding-sibling::w:tc and ../w:trPr/w:gridBefore/@w:val) then ../w:trPr/w:gridBefore/@w:val else 0)" />
    <xsl:choose>
      <xsl:when test="$span &gt; 0">
        <!-- Get the current cell number -->
        <xsl:attribute name="namest">
          <xsl:value-of  select="concat('col', $colstart)"/>
        </xsl:attribute>
        <xsl:attribute name="nameend">
          <xsl:value-of  select="concat('col', $colstart + $span - 1)"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="colname">
          <xsl:value-of  select="concat('col', $colstart)"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:function name="docx2hub:is-blind-vmerged-cell" as="xs:boolean">
    <xsl:param name="cell" as="element(w:tc)"/>
    <xsl:sequence select="exists($cell/w:tcPr/w:vMerge)
                          and
                          (every $att in $cell/w:tcPr/w:vMerge/@* satisfies ($att/self::attribute(srcpath) or $att/self::attribute(w:val)[. = 'continue']))"/>
  </xsl:function>

  <xsl:template name="cell.morerows">
    <xsl:if test="w:tcPr/w:vMerge/@w:val = 'restart'">
      <xsl:variable name="counts" as="xs:integer*">
        <xsl:choose>
          <xsl:when test="../following-sibling::w:tr[1]/w:tc[letex:colcount(1, .) = letex:colcount(1, current())][docx2hub:is-blind-vmerged-cell(.)]">
            <xsl:for-each-group select="../following-sibling::w:tr/w:tc[letex:colcount(1, .) = letex:colcount(1, current())]" 
              group-adjacent="docx2hub:is-blind-vmerged-cell(.)">
              <xsl:sequence select="count(current-group())"/>
            </xsl:for-each-group>
          </xsl:when>
          <xsl:otherwise>
            <xsl:sequence select="24"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="morerows" select="$counts[1]" />
      <!-- uncommented, mk: do not invent information where none exist -->
      <!--<xsl:attribute name="css:border-bottom-style" select="'none'" />-->
    </xsl:if>
  </xsl:template>

  <!-- recursively count preceding columns, including spans -->
  <xsl:function name="letex:colcount" as="xs:double">
    <xsl:param name="count" as="xs:double"/>
    <xsl:param name="cell"  as="element(*)"/>
    
    <xsl:choose>
      <xsl:when test="$cell/preceding-sibling::w:tc">
        <xsl:variable name="span" select="if ($cell/preceding-sibling::w:tc[1]/w:tcPr/w:gridSpan/@w:val) 
                                          then $cell/preceding-sibling::w:tc[1]/w:tcPr/w:gridSpan/@w:val
                                          else 1"/>
        <xsl:sequence select="letex:colcount($count + $span, $cell/preceding-sibling::w:tc[1])" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$count"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <!-- Tabellen innerhalb einer Tabellenzelle -->
  <xsl:template match="w:tbl" mode="tables">
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>


</xsl:stylesheet>
