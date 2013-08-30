docx_modify: an XSLT 2 / XProc based tool to manipulate .docx files.

Usage: [DEBUG=yes|no] [HEAP=xxxxm] docx_modifiy XSL DOCX [XPL]

The prefixed DEBUG=yes or HEAP=2048m work only if your shell is
bash-compatible.

tcsh users (once per shell, and only if the default 1024m are insufficient):
setenv HEAP 2048m

Windows users: use the calabash/calabash.sh style invocation (see
below), except that you use calabash/calabash.bat instead of the .sh
script. Sorry no docx_modify.bat file yet.

Installation: 

svn checkout \
    https://subversion.le-tex.de/docxtools/trunk/docx_modify/frontend/ \
    docx_modify 
Will fetch XML Calabash with an extension and other prerequisites from
svn externals.  System requirements: 2 GB RAM, Java 1.6 or newer.

Sample invocation (identity with debug):
DEBUG=yes ./docx_modifiy.sh lib/xsl/identity.xsl /path/to/myfile.docx

The resulting file will end in .mod.docx, with the same base name.
The .docx file's directory will be used to temporarily extract the files.
So you need write permissions there (also for the resulting file).
There's no option yet to create the .mod.docx in another place.

The third argument [XPL] is the path (may be relative) to an optional
XProc pipeline that implements the modification. If the desired
modifications are such that they require multiple passes in different
XSLT modes, it's most convenient to encapsulate them into their own
XProc pipeline. You can derive your own pipeline from
lib/xpl/single-pass_modify and add more letex:xslt-mode steps.  This
pipeline is invoked by adding -i xpl=lib/xpl/single-pass_modify.xpl (in
case you want to explicitly specify the default pipeline) to the
calabash.sh invocation, or by supplying the xpl file URL as a third
argument to docx_modify.sh 

See lib/xsl for transformation examples (e.g., identity for
reproducing the .docx identically, or rename_pstyles.xsl for renaming
paragraphs. More complex, but also very specific examples are in
epub_formatsicherung.xsl (including generation of new character styles
from actual formatting) or page-bookmarks.xsl (this uses an XSLT
micropipeline for a multi-pass transformation -- if the XPL option had
been in place by the time we wrote this transformation, we would have
used it. Now there are two alternative invocations:

./docx_modifiy.sh lib/xsl/page-bookmarks.xsl /path/to/myfile.docx
./docx_modifiy.sh lib/xsl/page-bookmarks.xsl /path/to/myfile.docx lib/xpl/page-bookmarks.xpl

If you don't use the mode docx2hub:modify in your custom XSLT, it is important that
you invoke docx2hub:modify's handling of @xml:base attributes. So if one of your modes
is mymode, you should include the following template:
  <xsl:template match="@xml:base" mode="mymode">
    <xsl:apply-templates select="." mode="docx2hub:modify"/>
  </xsl:template>

Of course you may invoke Calabash directly, either by using
calabash/calabash.sh (see how the command line is assembled in
docx_modify.sh) or, bare metal, using the java command.

Other example:
calabash/calabash.sh \
		     -i xslt=lib/xsl/epub_formatsicherung.xsl \
		     lib/xpl/docx_modify.xpl \
		     file=tmp/9783451800139-color_hyphen_klein.docx \
		     debug=yes
is equivalent with:
./docx_modify.sh lib/xsl/epub_formatsicherung.xsl \
		 tmp/9783451800139-color_hyphen_klein.docx

A sample application is hub2docx (check out
https://subversion.le-tex.de/docxtools/trunk/hub2docx/frontend/),
which uses docx_modify to replace a docx templates content with
content that is converted from Hub XML (which is a superset of DocBook
Publishers 5.1, so it will work for many DocBook files).

Please note that this tool requires a Calabash extension
(calabash/lib/ltx/ltx-unzip/) that is included here as an svn external
but not in a vanilla XML Calabash distribution.


(C) 2013, le-tex publising services GmbH.  All rights reserved.

XML Calabash and other tools/libraries are attached as svn externals.
Their indivudual license terms pertain to these tools/libraries.

This tool is published under a Simplified BSD License:

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

   1. Redistributions of source code must retain the above copyright 
      notice, this list of conditions and the following disclaimer.

   2. Redistributions in binary form must reproduce the above copyright 
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY LE-TEX PUBLISING SERVICES ``AS IS'' AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL LE-TEX PUBLISING SERVICES OR CONTRIBUTORS 
BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR
BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE
OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN
IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.



