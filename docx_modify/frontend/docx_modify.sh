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
    echo "Usage: [DEBUG=yes|no] [HEAP=xxxxm] docx_modifiy XSL DOCX";
    echo "Sample invocation (identity): ./docx_modifiy.sh lib/xsl/identity.xsl /path/to/myfile.docx yes";
    exit 1;
fi

if [ -z $DOCX ]; then
    echo "Please supply a .docx file as second argument"
    exit 1
fi

if [ -z $DEBUG ]; then
    DEBUG=no
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

$DIR/calabash/calabash.sh -i xslt="$XSL" $MODIFY_XPL "$XPL" file="$DOCX" debug=$DEBUG
