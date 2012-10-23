SAXON	= saxon
JING = jing
UMASK	= umask 002

MAKEFILEDIR := $(abspath $(dir $(lastword $(MAKEFILE_LIST))))

ERROR_MSG_PATH  = $(MAKEFILEDIR)

ifeq ($(shell uname -o),Cygwin)
win_path = $(shell cygpath -ma $(1))
uri = $(shell echo file:///$(call win_path,$(1))  | perl -pe 's/ /%20/g')
else
uri = $(shell echo file://$(abspath $(1))  | perl -pe 's/ /%20/g')
endif

usage:
	@echo "Usage:"
	@echo "  make -f $(MAKEFILEDIR)/Makefile path_to/xyz.hub.xml"
	@echo "  make -f $(MAKEFILEDIR)/Makefile path_to/xyz.hub.xml.val"
	@echo "where path_to/xyz.docx is the source docx file"
	@echo "Use XSLTPARAMS=\"srcpath=no\" if you don't need srcpath attributes (will speed things up) (is default now)"

%.hub.xml: %.docx.tmp/word/document.xml $(MAKEFILEDIR)/Makefile $(shell find $(MAKEFILEDIR) -name \*.xsl) $(MAKEFILEDIR)/*xml
	-$(UMASK); \
	$(SAXON) \
		$(SAXONOPTS) \
		-it:main \
		-s:$(call uri,$<) \
		-xsl:$(call uri,$(MAKEFILEDIR)/main.xsl) \
		error-msg-file-path=$(call uri,$(ERROR_MSG_PATH)) \
		$(XSLTPARAMS) \
		> $@ \
		2> $(abspath $(dir $@))/xslt_wml2dbk.log
		cat $(abspath $(dir $@))/xslt_wml2dbk.log

%.hub.xml.val: %.hub.xml $(MAKEFILEDIR)/Makefile $(shell find $(MAKEFILEDIR)/Hub -name \*.rng)
	$(JING) $(MAKEFILEDIR)/Hub/hub.rng "$<" > "$@"
	cat "$@"

%.docx.tmp/word/document.xml: %.docx
	-$(UMASK); mkdir "$<.tmp"
	$(UMASK); unzip -o -d "$<.tmp" "$<"
	$(UMASK); find "$<.tmp" -exec touch "{}" \;

# Keep all intermediate files:
.SECONDARY:
