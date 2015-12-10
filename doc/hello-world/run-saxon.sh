#
INFILE=intermediate_not_matched.xml
OUTFILE=document.xml
MODE=default # Second pass: clean, third: merge

CODEDIR=../..
java -cp $CODEDIR/calabash/resolver/resolver.jar:$CODEDIR/calabash/lib/saxon9he.jar net.sf.saxon.Transform -s:$INFILE --defaultCollection:collection.xml -o:$OUTFILE -xsl:$CODEDIR/hub2docx/lib/xsl/hub2docx.xsl -catalog:$CODEDIR/xslt-util/xmlcatalog/catalog.xml';'$CODEDIR/docx_modify/lib/xmlcatalog/catalog.xml -im:'{http://www.le-tex.de/namespace/hub}'$MODE
# -explain:explain.xml -opt:0
