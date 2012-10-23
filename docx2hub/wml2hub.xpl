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
  <p:output port="result"/>
    
  <p:option name="docx" required="true"/>
  <p:option name="debug" required="false" select="'no'"/>
  <p:option name="debug-dir-uri" required="false" select="''"/>
  <p:option name="unwrap-tooltip-links" required="false" select="'no'"/>

  <!-- import libs with extension steps -->
  
  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl" />
  <p:import href="http://transpect.le-tex.de/calabash-extensions/ltx-unzip/ltx-lib.xpl" />
  <p:import href="wml2hub.lib.xpl" />

  <!-- unzip or print out error message -->
    
  <letex:unzip name="unzip">
    <p:with-option name="zip" select="$docx" />
    <p:with-option name="dest-dir" select="concat($docx, '.tmp')">
      <p:pipe step="docx2hub" port="source"/>
    </p:with-option>
    <p:with-option name="overwrite" select="'yes'" />
  </letex:unzip>

  <p:load name="designmap">
    <p:with-option name="href" select="concat(/c:files/@xml:base, 'word/document.xml')" />
  </p:load>

  <p:xslt initial-mode="insert-xpath" name="insert-xpath">
    <p:input port="stylesheet">
      <p:document href="main.xsl"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <docx2hub:store-debug pipeline-step="01.insert-xpath">
    <p:with-option name="active" select="$debug" />
    <p:with-option name="specified-uri" select="$debug-dir-uri"/>
    <p:with-option name="default-uri" select="concat(/c:files/@xml:base, '/debug')"><p:pipe step="unzip" port="result" /></p:with-option>
  </docx2hub:store-debug>
  
  <p:xslt initial-mode="docx2hub:add-props" name="add-props">
    <p:input port="source">
      <p:pipe port="result" step="insert-xpath"/>
    </p:input>
    <p:input port="stylesheet">
      <p:document href="main.xsl"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <docx2hub:store-debug pipeline-step="03.add-props">
    <p:with-option name="active" select="$debug" />
    <p:with-option name="specified-uri" select="$debug-dir-uri"/>
    <p:with-option name="default-uri" select="concat(/c:files/@xml:base, '/debug')"><p:pipe step="unzip" port="result" /></p:with-option>
  </docx2hub:store-debug>
  
  <p:xslt initial-mode="docx2hub:props2atts" name="props2atts">
    <p:input port="stylesheet">
      <p:document href="main.xsl"/>
    </p:input>
    <p:input port="source">
      <p:pipe port="result" step="add-props"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <docx2hub:store-debug pipeline-step="04.props2atts">
    <p:with-option name="active" select="$debug" />
    <p:with-option name="specified-uri" select="$debug-dir-uri"/>
    <p:with-option name="default-uri" select="concat(/c:files/@xml:base, '/debug')"><p:pipe step="unzip" port="result" /></p:with-option>
  </docx2hub:store-debug>
  
  <p:xslt initial-mode="docx2hub:remove-redundant-run-atts" name="remove-redundant-run-atts">
    <p:input port="stylesheet">
      <p:document href="main.xsl"/>
    </p:input>
    <p:input port="source">
      <p:pipe port="result" step="props2atts"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <docx2hub:store-debug pipeline-step="05.remove-redundant-run-atts">
    <p:with-option name="active" select="$debug" />
    <p:with-option name="specified-uri" select="$debug-dir-uri"/>
    <p:with-option name="default-uri" select="concat(/c:files/@xml:base, '/debug')"><p:pipe step="unzip" port="result" /></p:with-option>
  </docx2hub:store-debug>
  
  <p:xslt initial-mode="docx2hub:separate-field-functions" name="separate-field-functions">
    <p:input port="stylesheet">
      <p:document href="main.xsl"/>
    </p:input>
    <p:input port="source">
      <p:pipe port="result" step="remove-redundant-run-atts"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <docx2hub:store-debug pipeline-step="11.separate-field-functions">
    <p:with-option name="active" select="$debug" />
    <p:with-option name="specified-uri" select="$debug-dir-uri"/>
    <p:with-option name="default-uri" select="concat(/c:files/@xml:base, '/debug')"><p:pipe step="unzip" port="result" /></p:with-option>
  </docx2hub:store-debug>
  
  <p:xslt initial-mode="wml-to-dbk" name="wml-to-dbk">
    <p:with-param name="error-msg-file-path" select="replace( static-base-uri(), '/wml2hub.xpl', '' )"/>
    <p:with-param name="unwrap-tooltip-links" select="$unwrap-tooltip-links"/>
    <p:input port="stylesheet">
      <p:document href="main.xsl"/>
    </p:input>
    <p:input port="source">
      <p:pipe port="result" step="separate-field-functions"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <docx2hub:store-debug pipeline-step="20.wml-to-dbk">
    <p:with-option name="active" select="$debug" />
    <p:with-option name="specified-uri" select="$debug-dir-uri"/>
    <p:with-option name="default-uri" select="concat(/c:files/@xml:base, '/debug')"><p:pipe step="unzip" port="result" /></p:with-option>
  </docx2hub:store-debug>
  
  <p:xslt initial-mode="docx2hub:join-runs" name="join-runs">
    <p:input port="stylesheet">
      <p:document href="main.xsl"/>
    </p:input>
    <p:input port="source">
      <p:pipe port="result" step="wml-to-dbk"/>
    </p:input>
    <p:input port="parameters"><p:empty/></p:input>
  </p:xslt>

  <docx2hub:store-debug pipeline-step="24.join-runs">
    <p:with-option name="active" select="$debug" />
    <p:with-option name="specified-uri" select="$debug-dir-uri"/>
    <p:with-option name="default-uri" select="concat(/c:files/@xml:base, '/debug')"><p:pipe step="unzip" port="result" /></p:with-option>
  </docx2hub:store-debug>
  
  <p:add-attribute match="/*" attribute-name="xml:base" name="rebase">
    <p:input port="source">
      <p:pipe port="result" step="join-runs"/>
    </p:input>
    <p:with-option name="attribute-value" select="replace(/c:files/@xml:base, '\.\w+\.tmp/?$', '.hub.xml')" >
      <p:pipe step="unzip" port="result"/>
    </p:with-option>
  </p:add-attribute>

</p:declare-step>