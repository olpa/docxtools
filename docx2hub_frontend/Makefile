ifeq ($(shell uname -o),Cygwin)
        win_path = $(shell cygpath -ma "$(1)")
        uri = $(shell echo file:///$(call win_path,$(1))  | perl -pe 's/ /%20/g')
else
        win_path = $(shell readlink -f "$(1)")
        uri = $(shell echo file:$(abspath $(1))  | perl -pe 's/ /%20/g')
endif


CALABASH = calabash/calabash.sh
UMASK	= umask 002
HUBVERSION = 1.0
DEBUG = no
DEBUGDIR = debug
SRCPATHS = no

MAKEFILEDIR := $(abspath $(dir $(lastword $(MAKEFILE_LIST))))

usage:
	@echo "Usage:"
	@echo "  make -f $(MAKEFILEDIR)/Makefile path_to/xyz.hub.xml"
	@echo "  make -f $(MAKEFILEDIR)/Makefile path_to/xyz.hub.xml.val"
	@echo "where path_to/xyz.docx is the source docx file"
	@echo "Use XSLTPARAMS=\"srcpath=no\" if you don't need srcpath attributes (will speed things up) (is default now)"
	@echo "Use HUBVERSION=1.1 for Hub 1.1 output."

%.hub.xml: %.docx.tmp/word/document.xml 
	-$(UMASK); \
	$(CALABASH) \
	-o result=$*.hub.xml \
	docx2hub/wml2hub.xpl \
	docx=$*.docx \
	debug=$(DEBUG) \
	debug-dir-uri=$(call uri,$(DEBUGDIR))
