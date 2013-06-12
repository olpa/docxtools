#!/bin/bash
cygwin=false;
case "`uname`" in
  CYGWIN*) cygwin=true;
esac

DIR="$( cd -P "$(dirname $( readlink -f "${BASH_SOURCE[0]}" ))" && pwd )"
XSL="$( readlink -f "$1" )"
DOCX="$( readlink -f "$2" )"
XPL="$DIR"/lib/xpl/docx_modify.xpl

if [ -z $DEBUG ]; then
    DEBUG=no
fi

if [ -z $HEAP ]; then
    HEAP=1024m
fi

if [ -z $XSL ]; then
    echo "Please supply an .xsl file as first argument (e.g., $DIR/lib/xsl/identity.xsl)"
    exit 1
fi

if [ -z $DOCX ]; then
    echo "Please supply a .docx file as second argument"
    exit 1
fi

if $cygwin; then
  XSL=file:/$(cygpath -ma $XSL)
  DOCX=$(cygpath -ma $DOCX)
  XPL=file:/$(cygpath -ma $XPL)
fi

$DIR/calabash/calabash.sh -i xslt="$XSL" "$XPL" file="$DOCX" debug=$DEBUG
