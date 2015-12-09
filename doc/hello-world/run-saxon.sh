#
INFILE=intermediate_not_matched.xml
OUTFILE=document.xml
MODE=default # Second pass: clean, third: merge

#CODEDIR=/home/olpa/p/third_party/le-tex/dbk2docx/frontend
CODEDIR=/home/olpa/tmp/dbk2docx/frontend
java -cp $CODEDIR/calabash/resolver/resolver.jar:$CODEDIR/calabash/lib/saxon9he.jar net.sf.saxon.Transform -s:$INFILE --defaultCollection:collection.xml -o:$OUTFILE -xsl:$CODEDIR/hub2docx/xsl/hub2docx.xsl -catalog:$CODEDIR/xslt-util/xmlcatalog/catalog.xml';'$CODEDIR/docx_modify/xmlcatalog/catalog.xml -im:'{http://www.le-tex.de/namespace/hub}'$MODE
# -explain:explain.xml -opt:0
