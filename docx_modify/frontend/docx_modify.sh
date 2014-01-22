#!/bin/bash
cygwin=false;
case "`uname`" in
  CYGWIN*) cygwin=true;
esac

DIR="$( cd -P "$(dirname $( readlink -f "${BASH_SOURCE[0]}" ))" && pwd )"
XSL="$( readlink -f "$1" )"
DOCX="$( readlink -f "$2" )"
MODIFY_XPL="$( readlink -f "$3" )"
XPL="$DIR"/lib/xpl/docx_modify.xpl

if [ -z $XSL ]; then
    echo "Usage: [DEBUG=yes|no] [HEAP=xxxxm] docx_modify XSL DOCX [XPL]";
    echo "(The prefixed DEBUG=yes or HEAP=2048m work only if your shell is bash-compatible.)";
    echo "";
    echo "Sample invocation (identity with debug): ";
    echo "DEBUG=yes ./docx_modifiy.sh lib/xsl/identity.xsl /path/to/myfile.docx";
    echo "";
    echo "The resulting file will end in .mod.docx, with the same base name.";
    echo "The .docx file's directory will be used to temporarily extract the files.";
    echo "So you need write permissions there (also for the resulting file)."
    echo "There's no option yet to create the .mod.docx in another place."
    echo "";
    echo "The third argument [XPL] is the path (may be relative) to an optional XProc pipeline ";
    echo "that implements the modification. If none is specified, lib/xpl/single-pass_modify.xpl";
    echo "will be used. See this XProc file as an example to build your own pipeline. Typically, ";
    echo "pipelines are built by chaining multiple letex:xslt-mode steps with the same stylesheet.";
    echo "";
    echo "See lib/xsl for transformation examples (e.g., identity for reproducing the .docx identically,";
    echo "or rename_pstyles.xsl for renaming paragraphs. More complex, but also very specific ";
    echo "examples are in epub_formatsicherung.xsl (including generation of new character styles";
    echo "from actual formatting) or page-bookmarks.xsl (this uses an XSLT micropipeline for";
    echo "a multi-pass transformation -- if the XPL option had been in place by the time we ";
    echo "wrote this transformation, we would have used it. Now there are two alternative";
    echo "invocations:";
    echo "./docx_modifiy.sh lib/xsl/page-bookmarks.xsl /path/to/myfile.docx";
    echo "./docx_modifiy.sh lib/xsl/page-bookmarks.xsl /path/to/myfile.docx lib/xpl/page-bookmarks.xpl";
    echo "";
    echo "If you don't use the mode docx2hub:modify in your custom XSLT, it is important that ";
    echo "you invoke docx2hub:modify's handling of @xml:base attributes. So if one of your modes";
    echo "is mymode, you should include the following template:";
    echo "  <xsl:template match=\"@xml:base\" mode=\"mymode\">";
    echo "    <xsl:apply-templates select=\".\" mode=\"docx2hub:modify\"/>";
    echo "  </xsl:template>";
    echo "";
    echo "Of course you may invoke Calabash directly, either by using calabash/calabash.sh";
    echo "(see how the command line is assembled in docx_modify.sh) or, bare metal, using";
    echo "the java command.";
    echo "Please note that this tool requires a Calabash extension (calabash/lib/ltx/ltx-unzip/)";
    echo "that is included here but not in a vanilla XML Calabash distribution.";
    exit 1;
fi

if [ -z $DOCX ]; then
    echo "Please supply a .docx file as second argument"
    exit 1
fi

if [ -z $DEBUG ]; then
    DEBUG=no
fi

if [ -z $DEBUGDIR ]; then
    DEBUGDIR=$DOCX.tmp/debug
fi

if [ -z $HEAP ]; then
    HEAP=1024m
fi

if $cygwin; then
  XSL=file:/$(cygpath -ma $XSL)
  DOCX=$(cygpath -ma $DOCX)
  XPL=file:/$(cygpath -ma $XPL)
  if [ ! -z $MODIFY_XPL ]; then
    MODIFY_XPL=file:/$(cygpath -ma $MODIFY_XPL)
  fi
fi

if [ ! -z $MODIFY_XPL ]; then
  MODIFY_XPL="-i xpl=$MODIFY_XPL"
fi

HEAP=$HEAP $DIR/calabash/calabash.sh -i xslt="$XSL" $MODIFY_XPL "$XPL" file="$DOCX" debug=$DEBUG debug-dir-uri=$DEBUGDIR
