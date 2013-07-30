Invocation examples:

calabash/calabash.sh -i xslt=lib/xsl/epub_formatsicherung.xsl lib/xpl/docx_modify.xpl file=tmp/9783451800139-color_hyphen_klein.docx debug=yes

# Front end script:
HEAP=2048m DEBUG=yes ./docx_modify.sh lib/xsl/epub_formatsicherung.xsl tmp/9783451800139-color_hyphen_klein.docx

# tcsh users (once per shell):
setenv HEAP 2048m
# and then: see below

# The same without debugging and without retaining tmp dirs:
./docx_modify.sh lib/xsl/epub_formatsicherung.xsl tmp/9783451800139-color_hyphen_klein.docx

# If the desired modifications are such that they require multiple passes in different XSLT modes,
# it's most convenient to encapsulate them into their own XProc pipeline. You can derive your own
# pipeline from lib/xpl/single-pass_modify and add more letex:xslt-mode steps.
# This pipeline is invoked by adding -i xpl=lib/xpl/single-pass_modify (in case you want to explicitly
# specify the default pipeline) to the calabash.sh invocation, or by supplying the xpl file URL as
# a third argument to docx_modify.sh
# A sample application is hub2docx, which uses docx_modify to replace a docx templates content
# with content that is converted from Hub XML.
