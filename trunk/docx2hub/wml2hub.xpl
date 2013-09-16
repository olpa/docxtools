<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step 
  xmlns:p="http://www.w3.org/ns/xproc" 
  xmlns:c="http://www.w3.org/ns/xproc-step" 
  xmlns:cx="http://xmlcalabash.com/ns/extensions"
  xmlns:letex="http://www.le-tex.de/namespace" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  version="1.0"
  name="docx2hub"
  type="docx2hub:convert"
  >

  <!-- * This script is used to convert docx to hub format. The output is stored in the directory of the input docx file
       * invoke with "sh calabash/calabash.sh wml2hub.xpl docx=PATH-TO-MY-DOCX-FILE.docx" 
       *
       * Import it with
       * <p:import href="http://transpect.le-tex.de/docx2hub/wml2hub.xpl" />
       * if you use it from transpect or if you imported this project as svn:external.
       * In the latter case, include the following line in you project's xmlcatalog/catalog.xml:
       * <nextCatalog catalog="../docx2hub/xmlcatalog/catalog.xml"/>
       * -->

  <p:input port="source">
    <p:empty/>
  </p:input>
  <p:output port="result" primary="true"/>
    
  <p:option name="docx" required="true">
    <p:documentation>A file name as recognized by your system's JVM. Not a file: URI.</p:documentation>
  </p:option>
  <p:option name="debug" required="false" select="'no'"/>
  <p:option name="debug-dir-uri" required="false" select="'debug'"/>
  <p:option name="srcpaths" required="false" select="'no'"/>
  <p:option name="unwrap-tooltip-links" required="false" select="'no'"/>
  <p:option name="hub-version" required="false" select="'1.1'"/>

  <p:serialization port="result" omit-xml-declaration="false"/>

  <!-- import libs with extension steps -->
  
  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl" />
  <p:import href="http://transpect.le-tex.de/calabash-extensions/ltx-lib.xpl" />
  <p:import href="http://transpect.le-tex.de/xproc-util/xml-model/prepend-hub-xml-model.xpl" />
  <p:import href="http://transpect.le-tex.de/xproc-util/xslt-mode/xslt-mode.xpl"/>
  <p:import href="wml2hub.lib.xpl" />


  <p:variable name="basename" select="replace($docx, '^(.+?)([^/\\]+)\.docx$', '$2')"/>

  <!-- unzip or print out error message -->
    
  <letex:unzip name="unzip">
    <p:with-option name="zip" select="$docx" />
    <p:with-option name="dest-dir" select="concat($docx, '.tmp')">
      <p:pipe step="docx2hub" port="source"/>
    </p:with-option>
    <p:with-option name="overwrite" select="'yes'" />
  </letex:unzip>

  <p:load name="document">
    <p:with-option name="href" select="concat(/c:files/@xml:base, 'word/document.xml')" />
  </p:load>
  
  <p:sink/>
  
  <p:load name="stylesheet" href="main.xsl"/>
  
  <p:sink/>

  <p:add-attribute attribute-name="value" match="/c:param-set/c:param[@name = 'error-msg-file-path']">
    <p:with-option name="attribute-value" select="replace( static-base-uri(), '/wml2hub.xpl', '' )"/>
    <p:input port="source">
      <p:inline>
        <c:param-set>
          <c:param name="error-msg-file-path"/>
          <c:param name="hub-version"/>
          <c:param name="unwrap-tooltip-links"/>
        </c:param-set>
      </p:inline>
    </p:input>
  </p:add-attribute>
  
  <p:add-attribute attribute-name="value" match="/c:param-set/c:param[@name = 'hub-version']">
    <p:with-option name="attribute-value" select="$hub-version"/>
  </p:add-attribute>

  <p:add-attribute name="params" attribute-name="value" match="/c:param-set/c:param[@name = 'unwrap-tooltip-links']">
    <p:with-option name="attribute-value" select="$unwrap-tooltip-links"/>
  </p:add-attribute>

  <letex:xslt-mode msg="yes" mode="insert-xpath">
    <p:input port="source"><p:pipe step="document" port="result" /></p:input>
    <p:input port="parameters"><p:pipe step="params" port="result" /></p:input>
    <p:input port="stylesheet"><p:pipe step="stylesheet" port="result" /></p:input>
    <p:input port="models"><p:empty /></p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/01')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-param name="srcpaths" select="$srcpaths"/>
  </letex:xslt-mode>
  
  <letex:xslt-mode msg="yes" mode="docx2hub:add-props">
    <p:input port="parameters"><p:pipe step="params" port="result" /></p:input>
    <p:input port="stylesheet"><p:pipe step="stylesheet" port="result" /></p:input>
    <p:input port="models"><p:empty /></p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/03')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
  </letex:xslt-mode>
  
  <letex:xslt-mode msg="yes" mode="docx2hub:props2atts">
    <p:input port="parameters"><p:pipe step="params" port="result" /></p:input>
    <p:input port="stylesheet"><p:pipe step="stylesheet" port="result" /></p:input>
    <p:input port="models"><p:empty /></p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/04')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
  </letex:xslt-mode>
  
  <letex:xslt-mode msg="yes" mode="docx2hub:remove-redundant-run-atts">
    <p:input port="parameters"><p:pipe step="params" port="result" /></p:input>
    <p:input port="stylesheet"><p:pipe step="stylesheet" port="result" /></p:input>
    <p:input port="models"><p:empty /></p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/05')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
  </letex:xslt-mode>
  
  <letex:xslt-mode msg="yes" mode="docx2hub:separate-field-functions">
    <p:input port="parameters"><p:pipe step="params" port="result" /></p:input>
    <p:input port="stylesheet"><p:pipe step="stylesheet" port="result" /></p:input>
    <p:input port="models"><p:empty /></p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/11')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
  </letex:xslt-mode>
  
  <letex:xslt-mode msg="yes" mode="wml-to-dbk">
    <p:input port="parameters"><p:pipe step="params" port="result" /></p:input>
    <p:input port="stylesheet"><p:pipe step="stylesheet" port="result" /></p:input>
    <p:input port="models"><p:empty /></p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/20')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="hub-version" select="$hub-version"/>
    <p:with-param name="srcpaths" select="$srcpaths"/>
  </letex:xslt-mode>
  
  <letex:xslt-mode msg="yes" mode="docx2hub:join-runs">
    <p:input port="parameters"><p:pipe step="params" port="result" /></p:input>
    <p:input port="stylesheet"><p:pipe step="stylesheet" port="result" /></p:input>
    <p:input port="models"><p:empty /></p:input>
    <p:with-option name="prefix" select="concat('docx2hub/', $basename, '/24')"/>
    <p:with-option name="debug" select="$debug"/>
    <p:with-option name="debug-dir-uri" select="$debug-dir-uri"/>
    <p:with-option name="hub-version" select="$hub-version"/>
  </letex:xslt-mode>

  <p:add-attribute match="/*" attribute-name="xml:base" name="rebase">
    <p:with-option name="attribute-value" select="replace(/c:files/@xml:base, '\.\w+\.tmp/?$', '.hub.xml')" >
      <p:pipe step="unzip" port="result"/>
    </p:with-option>
  </p:add-attribute>

  <letex:prepend-hub-xml-model name="pi">
    <p:with-option name="hub-version" select="$hub-version"/>
  </letex:prepend-hub-xml-model>
  
</p:declare-step>
