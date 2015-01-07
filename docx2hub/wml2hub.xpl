<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions" xmlns:letex="http://www.le-tex.de/namespace"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:docx2hub="http://www.le-tex.de/namespace/docx2hub"
  xmlns:transpect="http://www.le-tex.de/namespace/transpect"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" version="1.0" name="docx2hub" type="docx2hub:convert">

  <p:documentation xmlns="http://www.w3.org/1999/xhtml">
    <p>This script is used to convert docx to Hub XML. By default, the output is stored in the same directory as the input docx
      file, with the same basename. It is a library that needs other externals. For standalone operation, please check out the
        <a href="https://subversion.le-tex.de/docxtools/trunk/docx2hub_frontend/">front-end project</a>.</p>
    <p>From the directory wherey you checked out the front-end project, you invoke it with:</p>
    <p><code>calabash/calabash.sh docx2hub/wml2hub.xpl docx=PATH-TO-MY-DOCX-FILE.docx</code></p>
    <p>where docx may be a an OS path or a file:, http:, or https: URL.</p>
    <p>Import it with</p>
    <p><code>&lt;p:import href="http://transpect.le-tex.de/docx2hub/wml2hub.xpl" /></code></p>
    <p>if you use it from transpect or if you imported this project as svn:external.</p>
    <p>In the latter case, include the following line in you project's xmlcatalog/catalog.xml:</p>
    <p><code>&lt;nextCatalog catalog="../docx2hub/xmlcatalog/catalog.xml"/></code></p>
    <p>Experts may override the default conversion rules by supplying custom XSLT (that imports main.xsl) on the 'xslt'
      port.</p>
  </p:documentation>

  <p:input port="xslt">
    <p:document href="main.xsl"/>
  </p:input>
  <p:input port="source" primary="true">
    <p:documentation>This is to prevent a default readable port connecting to this step’s xslt port.</p:documentation>
    <p:empty/>
  </p:input>
  <p:output port="result" primary="true"/>
  <p:output port="insert-xpath">
    <p:pipe step="single-tree" port="result"/>
  </p:output>

  <p:option name="docx" required="true">
    <p:documentation>OS name (preferably with full path, may not resolve if only a relative path is given), file:, http:, or
      https: URL. The file will be fetched first if it is an HTTP URL.</p:documentation>
  </p:option>
  <p:option name="debug" select="'no'"/>
  <p:option name="debug-dir-uri" select="'debug'"/>
  <p:option name="status-dir-uri" select="'status'"/>
  <p:option name="srcpaths" select="'no'"/>
  <p:option name="unwrap-tooltip-links" select="'no'"/>
  <p:option name="hub-version" select="'1.1'"/>
  <p:option name="fail-on-error" select="'no'"/>
  <p:option name="field-vars" select="'no'"/>
  <p:option name="extract-dir" select="''">
    <p:documentation>Directory (OS path, not file: URL) to which the file will be unzipped. If option is empty string, will be
      '.tmp' appended to OS file path.</p:documentation>
  </p:option>
  <p:option name="create-svg" required="false" select="'no'">
    <p:documentation>Whether Office Open Drawing ML should be mapped to SVG</p:documentation>
  </p:option>
  <p:serialization port="result" omit-xml-declaration="false"/>

  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>
  <p:import href="http://transpect.le-tex.de/calabash-extensions/ltx-lib.xpl"/>
  <p:import href="http://transpect.le-tex.de/xproc-util/file-uri/file-uri.xpl"/>
  <p:import href="http://transpect.le-tex.de/xproc-util/xml-model/prepend-hub-xml-model.xpl"/>
  <p:import href="http://transpect.le-tex.de/xproc-util/xslt-mode/xslt-mode.xpl"/>
  <p:import href="http://transpect.le-tex.de/book-conversion/converter/xpl/simple-progress-msg.xpl"/>
  <p:import href="single-tree.xpl"/>

  <letex:simple-progress-msg name="start-msg" file="docx2hub-start.txt">
    <p:input port="msgs">
      <p:inline>
        <c:messages>
          <c:message xml:lang="en">Starting DOCX to flat Hub XML conversion</c:message>
          <c:message xml:lang="de">Beginne Konvertierung von DOCX zu flachem Hub XML</c:message>
        </c:messages>
      </p:inline>
    </p:input>
    <p:with-option name="status-dir-uri" select="$status-dir-uri"/>
  </letex:simple-progress-msg>

  <docx2hub:single-tree name="single-tree">
    <p:with-option name="docx" select="$docx"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-option name="unwrap-tooltip-links" select="$unwrap-tooltip-links"/>
    <p:with-option name="fail-on-error" select="$fail-on-error"/>
    <p:with-option name="field-vars" select="$field-vars"/>
    <p:with-option name="srcpaths" select="$srcpaths"/>
    <p:with-option name="extract-dir" select="$extract-dir"/>
    <p:input port="xslt">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
  </docx2hub:single-tree>

  <letex:xslt-mode msg="yes" mode="docx2hub:add-props">
    <p:input port="parameters">
      <p:pipe step="single-tree" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[name='basename']/@value, '/03')">
      <p:pipe port="params" step="single-tree"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="field-vars" select="$field-vars"/>
  </letex:xslt-mode>

  <letex:xslt-mode msg="yes" mode="docx2hub:props2atts">
    <p:input port="parameters">
      <p:pipe step="single-tree" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[name='basename']/@value, '/05')">
      <p:pipe port="params" step="single-tree"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </letex:xslt-mode>

  <letex:xslt-mode msg="yes" mode="docx2hub:remove-redundant-run-atts">
    <p:input port="parameters">
      <p:pipe step="single-tree" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[name='basename']/@value, '/07')">
      <p:pipe port="params" step="single-tree"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </letex:xslt-mode>

  <letex:xslt-mode msg="yes" mode="docx2hub:separate-field-functions">
    <p:input port="parameters">
      <p:pipe step="single-tree" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[name='basename']/@value, '/11')">
      <p:pipe port="params" step="single-tree"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </letex:xslt-mode>

  <letex:xslt-mode msg="yes" mode="wml-to-dbk">
    <p:input port="parameters">
      <p:pipe step="single-tree" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[name='basename']/@value, '/20')">
      <p:pipe port="params" step="single-tree"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-param name="srcpaths" select="$srcpaths"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
    <p:with-param name="field-vars" select="$field-vars"/>
  </letex:xslt-mode>

  <letex:xslt-mode msg="yes" mode="docx2hub:join-runs">
    <p:input port="parameters">
      <p:pipe step="single-tree" port="params"/>
    </p:input>
    <p:input port="stylesheet">
      <p:pipe step="docx2hub" port="xslt"/>
    </p:input>
    <p:input port="models">
      <p:empty/>
    </p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', /c:param-set/c:param[name='basename']/@value, '/24')">
      <p:pipe port="params" step="single-tree"/> 
    </p:with-option>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-param name="fail-on-error" select="$fail-on-error"/>
  </letex:xslt-mode>

  <p:add-attribute match="/*" attribute-name="xml:base" name="rebase">
    <p:with-option name="attribute-value" 
      select="replace(/c:param-set/c:param[name='local-href']/@value, '\.do[ct][xm]$', '.hub.xml')">
      <p:pipe step="single-tree" port="result"/>
    </p:with-option>
  </p:add-attribute>

  <letex:prepend-hub-xml-model name="pi">
    <p:with-option name="hub-version" select="$hub-version"/>
  </letex:prepend-hub-xml-model>
  
  <letex:simple-progress-msg name="success-msg" file="docx2hub-success.txt">
    <p:input port="msgs">
      <p:inline>
        <c:messages>
          <c:message xml:lang="en">Successfully finished DOCX to flat Hub XML conversion</c:message>
          <c:message xml:lang="de">Konvertierung von DOCX zu flachem Hub XML erfolgreich abgeschlossen</c:message>
        </c:messages>
      </p:inline>
    </p:input>
    <p:with-option name="status-dir-uri" select="$status-dir-uri"/>
  </letex:simple-progress-msg>

</p:declare-step>