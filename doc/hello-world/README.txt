Convert DocBook to .docx
========================

Instruction how to run the XSLT2 stylesheets without the calabash-infrastructure.

transpect
---------

Documentation:
http://www.le-tex.de/en/transpect.html

svn checkout https://subversion.le-tex.de/common/dbk2docx

make -f .../dbk2docx/frontend/Makefile conversion IN_FILE=hello.xml

saxon: preparing files
----------------------

1) Produce intermediate files

1a) Edit the file "dbk2docx/frontend/hub2docx/xpl/hub2docx.xpl".

Near other imports, add:

  <p:import href="http://transpect.le-tex.de/xproc-util/store-debug/store-debug.xpl" />

Before the step "transformed-hub", add:

  <letex:store-debug name="store-matched">
    <p:input port="source">
      <p:pipe step="split" port="matched"/>
    </p:input>
    <p:with-option name="active" select="$debug" />
    <p:with-option name="pipeline-step" select="'hub2docx/5_matched'" />
    <p:with-option name="base-uri" select="$debug-dir-uri" />
  </letex:store-debug>
  <letex:store-debug name="store-not-matched">
    <p:input port="source">
      <p:pipe step="split" port="not-matched"/>
    </p:input>
    <p:with-option name="active" select="$debug" />
    <p:with-option name="pipeline-step" select="'hub2docx/5_not_matched'" />
    <p:with-option name="base-uri" select="$debug-dir-uri" />
  </letex:store-debug>
  <p:sink/>

My version of "hub2docx.xpl" is in directory "support".

1b) Run generation in debug mode:

make -f .../dbk2docx/frontend/Makefile conversion \
   IN_FILE=hello.xml DEBUG=yes

Result) The directory "debug" appears, with the files "hub2docx/5_matched.xml" and "hub2docx/5_not_matched.xml".

2) Collect the files for saxon

Make a copy of the intermediate files "5_matched.xml" and "5_not_matched.xml". The example uses "intermediate_matched.xml" and "intermediate_not_matched.xml".

Saxon will need these files also as a collection. The definition in "collection.xml" looks like:

<collection stable="true">
  <doc href="intermediate_matched.xml"/>
  <doc href="intermediate_not_matched.xml"/>
  </collection>
</collection>

Put the file "resolver.jar" into "dbk2docx/frontend/calabash/resolver". I found one in other le-tex svn modules. A copy is in the directory "support".

Make directory "dbk2docx/frontend/xslt-util/xmlcatalog", put the file "support/catalog.xml" into this directory.

run saxon
---------

The content of "run-saxon.sh":

INFILE=intermediate_not_matched.xml
OUTFILE=document.xml
MODE=default # Second pass: clean, third: merge

CODEDIR=.../dbk2docx/frontend
java -cp $CODEDIR/calabash/resolver/resolver.jar:$CODEDIR/calabash/lib/saxon9he.jar \
  net.sf.saxon.Transform -s:$INFILE \
  --defaultCollection:collection.xml -o:$OUTFILE \
  -xsl:$CODEDIR/hub2docx/xsl/hub2docx.xsl \
  -catalog:$CODEDIR/xslt-util/xmlcatalog/catalog.xml';'$CODEDIR/docx_modify/xmlcatalog/catalog.xml \
  -im:'{http://www.le-tex.de/namespace/hub}'$MODE
# -explain:explain.xml -opt:0

The conversion is done in three passes. Each pass is defined by the mode. The first: "default", the second: "clean", the third: "merge".


