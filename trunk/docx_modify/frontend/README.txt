Invocation examples:

calabash/calabash.sh -i xslt=lib/xsl/epub_formatsicherung.xsl lib/xpl/docx_modify.xpl file=tmp/9783451800139-color_hyphen_klein.docx debug=yes

# Front end script:
HEAP=2048m DEBUG=yes ./docx_modify.sh lib/xsl/epub_formatsicherung.xsl tmp/9783451800139-color_hyphen_klein.docx

# The same without debugging and retaining tmp dirs:
./docx_modify.sh lib/xsl/epub_formatsicherung.xsl tmp/9783451800139-color_hyphen_klein.docx
