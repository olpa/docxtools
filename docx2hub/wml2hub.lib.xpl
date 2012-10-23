<p:library 
  xmlns:p="http://www.w3.org/ns/xproc" 
  xmlns:c="http://www.w3.org/ns/xproc-step" 
  xmlns:cx="http://xmlcalabash.com/ns/extensions" 
  xmlns:pxf="http://exproc.org/proposed/steps/file" 
  xmlns:pxp="http://exproc.org/proposed/steps" 
  xmlns:ltx="http://le-tex.de/tools/unzip" 
  xmlns:pkg="http://expath.org/ns/pkg" 
  xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
  xmlns:tr="http://le-tex.de/namespace/transpect" 
  xmlns:docx2hub = "http://www.le-tex.de/namespace/docx2hub"
  pkg:import-uri="http://le-tex.de/tools/unzip.xpl" version="1.0">

  <p:documentation xmlns="http://www.w3.org/1999/xhtml">
    <html><title>Transpect Plugin Installer Extension Library</title></html>
    <body>
      <div>
        <h1>Transpect Plugin Installer Extension Library</h1>
        <h2>Version 1.0</h2>
        <p>These library provides some steps for installer tasks.</p>
      </div>
    </body>
  </p:documentation>

  <p:declare-step type="docx2hub:store-debug" name="store-debug">
    <p:input port="source" primary="true"/>
    <p:option name="active" required="false" select="'no'"/>
    <p:option name="pipeline-step" required="true"/>
    <p:option name="default-uri" required="false" select="resolve-uri('debug')"/>
    <p:option name="specified-uri" required="false"/>
    <p:choose>
      <p:when test="$active = 'yes'">
        <p:store indent="true">
          <p:input port="source">
            <p:pipe step="store-debug" port="source" />
          </p:input>
          <p:with-option name="href" select="if ($specified-uri != '') 
                                             then concat($specified-uri, '/', $pipeline-step, '.xml')
                                             else concat($default-uri, '/', $pipeline-step, '.xml')" />
        </p:store>
      </p:when>
      <p:otherwise>
        <p:sink/>
      </p:otherwise>
    </p:choose>
  </p:declare-step>


</p:library>
