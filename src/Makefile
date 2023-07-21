.PHONY: all
all: addin

.PHONY: addin
addin: bin/GetPressedAddin.ppam

MODULES = Module.vb
CLASSES = AddinClass.vb
CUSTOMUI = $(wildcard customUI/*.xml)
CUSTOMUI_IMAGES = $(wildcard customUI/images/*.png)

bin/GetPressedAddin.ppam: $(MODULES) $(CLASSES) $(CUSTOMUI) $(CUSTOMUI_IMAGES)
	vbamc -m $(MODULES) -c $(CLASSES) -n "GetPressedAddin" --company "Acme" -f GetPressedAddin.ppam

.PHONY: clean
clean:
	rm -rf bin/ obj/
