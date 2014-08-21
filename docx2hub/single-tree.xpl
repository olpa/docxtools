<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions" xmlns:letex="http://www.le-tex.de/namespace"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:docx2hub="http://www.le-tex.de/namespace/docx2hub"
  xmlns:transpect="http://www.le-tex.de/namespace/transpect"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" version="1.0" name="docx-single-tree"
  type="docx2hub:single-tree">

  <p:input port="source" primary="true">
    <p:documentation>This is to prevent any other default readable port to be connected with the xslt port.</p:documentation>
    <p:empty/>
  </p:input>
  <p:input port="xslt" primary="false">
    <p:document href="main.xsl"/>
  </p:input>
  <p:output port="result" primary="true"/>
  <p:output port="params">
    <p:pipe port="result" step="params"/>
  </p:output>
  <p:output port="zip-manifest">
    <p:pipe port="result" step="unzip"/>
  </p:output>

  <p:option name="docx" required="true">
    <p:documentation>OS name (preferably with full path, may not resolve if only a relative path is given), file:, http:, or
      https: URL. The file will be fetched first if it is an HTTP URL.</p:documentation>
  </p:option>
  <p:option name="debug" required="false" select="'no'"/>
  <p:option name="debug-dir-uri" required="false" select="'debug'"/>
  <p:option name="srcpaths" required="false" select="'no'"/>
  <p:option name="unwrap-tooltip-links" required="false" select="'no'"/>
  <p:option name="hub-version" required="false" select="'1.1'"/>
  <p:option name="fail-on-error" select="'no'"/>
  <p:option name="extract-dir" required="false" select="''">
    <p:documentation>Directory (OS path, not file: URL) to which the file will be unzipped. If option is empty string, will be
      '.tmp' appended to OS file path.</p:documentation>
  </p:option>
  <p:serialization port="result" omit-xml-declaration="false"/>

  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>
  <p:import href="http://transpect.le-tex.de/calabash-extensions/ltx-lib.xpl"/>
  <p:import href="http://transpect.le-tex.de/xproc-util/file-uri/file-uri.xpl"/>
  <p:import href="http://transpect.le-tex.de/xproc-util/xslt-mode/xslt-mode.xpl"/>

  <p:variable name="basename" select="replace($docx, '^(.+?)([^/\\]+)\.do[ct][mx]$', '$2')"/>

  <transpect:file-uri name="locate-docx">
    <p:with-option name="filename" select="$docx"/>
  </transpect:file-uri>

  <!-- unzip or error message -->
  <letex:unzip name="unzip">
    <p:with-option name="zip" select="/c:result/@os-path">
      <p:pipe step="locate-docx" port="result"/>
    </p:with-option>
    <p:with-option name="dest-dir"
      select="if ($extract-dir = '') 
              then concat(/c:result/@os-path, '.tmp')
              else $extract-dir">
      <p:pipe step="locate-docx" port="result"/>
    </p:with-option>
    <p:with-option name="overwrite" select="'yes'"/>
  </letex:unzip>

  <p:choose>
    <p:when test="name(/*) eq 'c:error'">
      <cx:message>
        <p:with-option name="message" select="'docx2hub error on unzipping.&#xa;', //text(), '&#xa;'"/>
      </cx:message>
    </p:when>
    <p:otherwise>
      <p:identity/>
    </p:otherwise>
  </p:choose>

  <p:load name="document">
    <p:with-option name="href" select="concat(/c:files/@xml:base, 'word/document.xml')"/>
  </p:load>

  <p:sink/>

  <p:add-attribute attribute-name="value" match="/c:param-set/c:param[@name = 'error-msg-file-path']">
    <p:with-option name="attribute-value" select="replace(static-base-uri(), '/[^/]+$', '')"/>
    <p:input port="source">
      <p:inline>
        <c:param-set>
          <c:param name="error-msg-file-path"/>
          <c:param name="hub-version"/>
          <c:param name="local-href"/>
          <c:param name="unwrap-tooltip-links"/>
          <c:param name="extract-dir-uri"/>
        </c:param-set>
      </p:inline>
    </p:input>
  </p:add-attribute>

  <p:add-attribute attribute-name="value" match="/c:param-set/c:param[@name = 'hub-version']">
    <p:with-option name="attribute-value" select="$hub-version"/>
  </p:add-attribute>

  <p:add-attribute attribute-name="value" match="/c:param-set/c:param[@name = 'basename']">
    <p:with-option name="attribute-value" select="$basename"/>
  </p:add-attribute>

  <p:add-attribute attribute-name="value" match="/c:param-set/c:param[@name = 'extract-dir-uri']">
    <p:with-option name="attribute-value" select="/c:files/@xml:base">
      <p:pipe port="result" step="unzip"/>
    </p:with-option>
  </p:add-attribute>

  <p:add-attribute attribute-name="value" match="/c:param-set/c:param[@name = 'local-href']">
    <p:with-option name="attribute-value" select="/c:result/@local-href">
      <p:pipe port="result" step="locate-docx"/>
    </p:with-option>
  </p:add-attribute>

  <p:add-attribute name="params" attribute-name="value" match="/c:param-set/c:param[@name = 'unwrap-tooltip-links']">
    <p:with-option name="attribute-value" select="$unwrap-tooltip-links"/>
  </p:add-attribute>

  <p:sink/>
  
  <letex:xslt-mode msg="yes" mode="insert-xpath">
    <p:input port="source">
      <p:pipe step="document" port="result"/>
    </p:input>
    <p:input port="parameters">
      <p:pipe step="params" port="result"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx-single-tree" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/01')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-param name="srcpaths" select="$srcpaths"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </letex:xslt-mode>

  <p:add-attribute attribute-name="xml:base" match="/*">
    <p:with-option name="attribute-value" select="/c:files/@xml:base">
      <p:pipe port="result" step="unzip"/>
    </p:with-option>
  </p:add-attribute>

</p:declare-step>