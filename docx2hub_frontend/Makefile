ifeq ($(shell uname -o),Cygwin)
        win_path = $(shell cygpath -ma "$(1)")
        uri = $(shell echo file:///$(call win_path,$(1))  | perl -pe 's/ /%20/g')
else
        win_path = $(shell readlink -f "$(1)")
        uri = $(shell echo file:$(abspath $(1))  | perl -pe 's/ /%20/g')
endif

MAKEFILEDIR := $(abspath $(dir $(lastword $(MAKEFILE_LIST))))

HEAP=1024m
CALABASH = $(MAKEFILEDIR)/calabash/calabash.sh HEAP=$(HEAP)
UMASK	= umask 002
DEBUG = no
DEBUGDIR = debug
SRCPATHS = no
PARAMS = docx=$(call win_path,$*).docx \
	 debug=$(DEBUG) \
	 debug-dir-uri=$(call uri,$(DEBUGDIR)) \
	 srcpaths=$(SRCPATHS)

usage:
	@echo "Usage:"
	@echo "  make -f $(MAKEFILEDIR)/Makefile path_to/xyz.hub.xml"
#	@echo "  make -f $(MAKEFILEDIR)/Makefile path_to/xyz.hub.xml.val"
	@echo "  make -f $(MAKEFILEDIR)/Makefile path_to/xyz.html"
	@echo "where path_to/xyz.docx is the source docx file"
	@echo "Parameters:"
	@echo "  DEBUG=yes|no (currently: $(DEBUG))"
	@echo "  DEBUGDIR=path_to/debug_files (currently: $(DEBUGDIR))"
	@echo "  SRCPATHS=yes|no (currently: $(SRCPATHS))"

%.hub.xml:
	-$(UMASK); \
	$(CALABASH) \
	-o result=$(call win_path,$*.hub.xml) \
	-o insert-xpath=$(call win_path,/dev/null) \
	-o report=$(call win_path,/dev/null) \
	-o zip-manifest=$(call win_path,/dev/null) \
	-o schema=$(call win_path,/dev/null) \
	$(call uri,$(MAKEFILEDIR)/docx2hub/wml2hub.xpl) \
	$(PARAMS)

%.html:
	-$(UMASK); \
	$(CALABASH) \
	-o result=$(call win_path,$*.html) \
	-o hub=$(call win_path,$*.hub.xml) \
	$(call uri,$(MAKEFILEDIR)/docx2hub/docx2html.xpl) \
	$(PARAMS)
