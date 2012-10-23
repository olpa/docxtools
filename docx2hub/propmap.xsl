<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="2.0"
    xmlns:xsl   = "http://www.w3.org/1999/XSL/Transform"
    xmlns:xs    = "http://www.w3.org/2001/XMLSchema"
    xmlns:docx2hub  = "http://www.le-tex.de/namespace/docx2hub"
    exclude-result-prefixes = "xs"
>

  <xsl:key name="docx2hub:prop" match="prop" use="@name" />

  <xsl:variable name="docx2hub:propmap" as="document-node(element(propmap))">
    <xsl:document xmlns="">
      <propmap>
        <prop name="v:shape/@fillcolor" type="docx-color" target-name="css:background-color"/>
        <prop name="v:shape/@id" />
        <prop name="v:shape/@o:allowoverlap"/>
        <prop name="v:shape/@stroked" type="linear" target-name="stroked"/><!-- no passthru b/c no atts may be produced in mode add-atts -->
        <prop name="v:shape/@strokeweight" type="linear" target-name="strokeweight"/>
        <prop name="v:shape/@style" type="linear" target-name="style"/>
        <prop name="v:shape/@type" />
        <prop name="w:adjustRightInd" />
        <prop name="w:autoSpaceDE" />
        <prop name="w:autoSpaceDN" />
        <prop name="w:b" type="docx-boolean-prop" target-name="css:font-weight" default="normal" active="bold"/>
        <prop name="w:bCs" />
        <prop name="w:color" type="docx-color" target-name="css:color"/>
        <prop name="w:gridSpan" /><!-- will be calculated by tables.xsl -->
        <prop name="w:highlight" type="docx-color" target-name="css:background-color"/>
        <prop name="w:i" type="docx-boolean-prop" target-name="css:font-style" default="normal" active="italic"/>
        <prop name="w:iCs" />
        <prop name="w:ind/@w:left" type="docx-length-attr" target-name="css:margin-left" />
        <prop name="w:ind/@w:right" type="docx-length-attr" target-name="css:margin-right" />
        <prop name="w:ind/@w:firstLine" type="docx-length-attr" target-name="css:text-indent" />
        <prop name="w:ind/@w:hanging" type="docx-length-attr-negated" target-name="css:text-indent" />
        <prop name="w:jc">
          <val match="left" target-name="css:text-align" target-value="left" />
          <val match="start" target-name="css:text-align" target-value="right" />
          <val match="right" target-name="css:text-align" target-value="right" />
          <val match="end" target-name="css:text-align" target-value="right" />
          <val match="both" target-name="css:text-align" target-value="justify" />
          <val match="center" target-name="css:text-align" target-value="center" />
          <val match="center" target-name="css:text-align-last" target-value="center" />
        </prop>
        <prop name="w:keepNext" />
        <prop name="w:kern" />
        <prop name="w:lang" type="linear" target-name="xml:lang" />
        <prop name="w:link" />
        <prop name="w:locked" />
        <prop name="w:name" type="linear" target-name="role"/>
        <prop name="w:next" />
        <prop name="w:noProof" />
        <prop name="w:numPr" type="passthru" />
        <prop name="w:outlineLvl" type="docx-hierarchy-level"/>
        <prop name="w:pageBreakBefore" type="docx-boolean-prop" target-name="css:page-break-before" default="auto" active="always"/>
        <prop name="w:pBdr/w:bottom" type="docx-border" />
        <prop name="w:pBdr/w:left" type="docx-border" />
        <prop name="w:pBdr/w:right" type="docx-border" />
        <prop name="w:pBdr/w:top" type="docx-border"  />
        <prop name="w:pStyle" type="docx-parastyle" />
        <prop name="w:qFormat" />
        <prop name="w:rFonts" type="docx-font-family" target-name="css:font-family" />
        <prop name="w:rsid" />
        <prop name="w:rsidR" />
        <prop name="w:rsidRDefault" />
        <prop name="w:rsidP" />
        <prop name="w:rsidRPr" />
        <prop name="w:rStyle" type="docx-charstyle" />
        <prop name="w:semiHidden" />
        <prop name="w:shadow" type="docx-boolean-prop" target-name="css:text-shadow" default="none" active="1pt 1pt"/>
        <prop name="w:shd" type="docx-shd" />
        <prop name="w:spacing/@w:after" type="docx-length-attr" target-name="css:margin-bottom" />
        <prop name="w:spacing/@w:before" type="docx-length-attr" target-name="css:margin-top" />
        <prop name="w:spacing/@w:afterLines" implement="maybe later" />
        <prop name="w:spacing/@w:beforeLines" implement="maybe later" />
        <prop name="w:spacing/@w:line" type="docx-length-attr" target-name="css:line-height" />
        <prop name="w:spacing/@w:val" implement="maybe later (letter spacing, a run property)" />
        <prop name="w:suppressAutoHyphens" type="docx-boolean-prop" target-name="css:hyphens" default="auto" active="manual"/>
        <prop name="w:sz" type="docx-font-size" target-name="css:font-size" />
        <prop name="w:szCs" />
        <prop name="w:tab/@w:leader" type="linear" target-name="leader" />
        <prop name="w:tab/@w:pos" type="docx-length-attr" target-name="horizontal-position" />
        <prop name="w:tab/@w:val">
          <!-- TODO: implement the remainder of the ST_TabJc values in § 17.18.84 -->
          <!-- Interestingly, 'left' and 'right' don't appear in that table. -->
          <val match="left" target-name="align" target-value="left" />
          <val match="center" target-name="align" target-value="center" />
          <val match="right" target-name="align" target-value="right" />
          <val match="num" />
          <val match="clear" target-name="clear" target-value="yes" />
        </prop>
        <prop name="w:tabs" type="tablist" />
        <prop name="w:tblBorders" type="passthru" />
        <prop name="w:tblCellMar" type="passthru" />
        <prop name="w:tblGrid" type="passthru" />
        <prop name="w:tblInd" implement="maybe later" />
        <prop name="w:tblLook" />
        <prop name="w:tblStyle" type="docx-parastyle"/>
        <prop name="w:tblW" type="passthru" />
        <prop name="w:tcBorders/w:bottom" type="docx-border" />
        <prop name="w:tcBorders/w:left" type="docx-border" />
        <prop name="w:tcBorders/w:right" type="docx-border" />
        <prop name="w:tcBorders/w:top" type="docx-border"  />
        <prop name="w:tcW/@w:w" type="docx-length-attr" target-name="css:width"/>
        <prop name="w:u" type="docx-underline" />
        <prop name="w:uiPriority" />
        <prop name="w:unhideWhenUsed" />
        <prop name="w:w" type="docx-font-stretch" target-name="css:font-stretch"/>
        <prop name="w:widowControl" type="docx-boolean-prop" target-name="css:orphans" default="1" active="2"/>
        <prop name="w:widowControl" type="docx-boolean-prop" target-name="css:widows" default="1" active="2"/>
      </propmap>
    </xsl:document>
  </xsl:variable>

</xsl:stylesheet>
