#target indesign

var assetName = "Page Controls";

/* --------------------------------------------------------------------------------
Page Controls Placement Script

This script places interactive "Page Controls" (Back/Forward navigation buttons)
from an InDesign Library onto a specified range of pages in the active document.

!!- IMPORTANT -!!
Placement on the page depends on object styles being applied to the control items
the object style should have "position" defined


Workflow:
1. Prompts the user to select a Page Controls library file (.indl).
2. Allows the user to define a starting and ending page (by number or name).
   - Defaults to first/last page if left blank.
3. Opens the selected library and retrieves the "Page Controls" asset.
4. Loops through the specified page range and:
   - Places the asset on each page.
   - Moves the placed group to the correct page.
   - Clears object style overrides on all placed items.
   - Identifies navigation buttons by name ("Back" and "Forward").
5. Applies logic per page:
   - Removes the "Back" button on the first page in the range.
   - Removes the "Forward" button on the last page in the range.
   - Adds hyperlinks:
       • "Back" → previous page
       • "Forward" → next page
6. Optimizes performance by disabling redraw, preflight, and user interaction
   during execution, then restores original preferences afterward.

Notes:
- Assumes the library asset contains items named "Back" and "Forward".
- Page range supports both numeric input and page names (section-aware).
- Hyperlink destinations are reused if they already exist.

-------------------------------------------------------------------------------- */

var oldRedrawPrefs, oldPreflightPrefs, oldInteractionPrefs;
var asset;

var doc = app.activeDocument;
if (!doc) {
    alert("Error - No Active Document\rYou must have a document open for this script to work.\r\rOpen a document and try again.");
    exit(0);
}


//-----------------------------------------------------------------------------------------------------------
//-- User defined settings, set thru dialog input
var settings = showOutputAndLanguageDialog();
if (!settings) exit(0);

var libraryFile = settings.libFile;
var startIndex = settings.startPage;
var endIndex = settings.endPage;



try{
    setAppPrefs();

    lib = app.open(libraryFile);

    var asset = lib.assets.itemByName(assetName);

    if (!asset.isValid) {
        alert("Could not find asset named \"" + assetName + "\" in the selected library.");
        restoreAppPrefs();
        exit(0);
    }

    for (i = startIndex; i <= endIndex; i++) {
            var page = doc.pages[i];

            var backItem, forwardItem;

            var placedObjects = asset.placeAsset(doc, [0, 0]);
            var placedGroup = placedObjects[0];
            placedGroup.move(page);

            var controlObjects = placedGroup.allPageItems;

            for(var a = 0; a < controlObjects.length; a++) {
                controlObjects[a].clearObjectStyleOverrides();

                if (controlObjects[a].name == "Back") {
                    backItem = controlObjects[a];
                } else if (controlObjects[a].name == "Forward") {
                    forwardItem = controlObjects[a];
                }
            }


            // First page: remove Back
            if (i == startIndex) {
                if (backItem && backItem.isValid) {
                    backItem.remove();
                }
            } else {
                if (backItem && backItem.isValid) {
                    addPageLink(doc, backItem, doc.pages[i - 1], "Back_Page_" + (i + 1));
                }
            }

            // Last page: remove Forward
            if (i == endIndex) {
                if (forwardItem && forwardItem.isValid) {
                    forwardItem.remove();
                }
            } else {
                if (forwardItem && forwardItem.isValid) {
                    addPageLink(doc, forwardItem, doc.pages[i + 1], "Forward_Page_" + (i + 1));
                }
            }
        }


} catch (scriptError) {
    alert("Error!\rLine: " + scriptError.line + "\r\r" + scriptError.message);
} finally {
    restoreAppPrefs();
}
    





function showOutputAndLanguageDialog() {
    //-- persistence via app labels
    var pageControlsLabelKey = "pageControlsLibPath";

    //-- preload remembered value (if any)
    var rememberedLibPath = readRememberedPath(pageControlsLabelKey);

    // ---------------------------------------------------------------------
    //-- Launch Dialog
    var settingsDialog = new Window("dialog", "Define Page Controls Library");
    settingsDialog.alignChildren = "fill";

    // ---------- Panel 0: Library file ----------
    var controlsLibPanel = settingsDialog.add("panel", undefined, "Page Controls Library");
    controlsLibPanel.orientation = "column";
    controlsLibPanel.alignChildren = ["fill","top"];
    controlsLibPanel.margins = 12;

    var controlsLibGroup = controlsLibPanel.add("group");
    controlsLibGroup.orientation = "row";
    controlsLibGroup.alignChildren = ["fill","center"];
    controlsLibGroup.spacing = 10;

    var controlsLibPath = controlsLibGroup.add("edittext", undefined, rememberedLibPath || "");
    controlsLibPath.characters = 45;
    controlsLibPath.helpTip = "Path to the Page Controls library";

    var browseButton = controlsLibGroup.add("button", undefined, "Browse...");

    browseButton.onClick = function () {
        var libFile = File.openDialog("Select page controls library", "*.indl", false);
        if (libFile) controlsLibPath.text = libFile.fsName;
    };

    // ---------- Panel 1: Page range ----------
    var pageRangePanel = settingsDialog.add("panel", undefined, "Page Range");
    pageRangePanel.orientation = "column";
    pageRangePanel.alignChildren = ["left", "top"];
    pageRangePanel.margins = 12;

    var pageRangeGroup = pageRangePanel.add("group");
    pageRangeGroup.orientation = "row";
    pageRangeGroup.alignChildren = ["left", "center"];
    pageRangeGroup.spacing = 12;

    pageRangeGroup.add("statictext", undefined, "Starting Page:");
    var startPageInput = pageRangeGroup.add("edittext", undefined, "");
    startPageInput.characters = 6;

    pageRangeGroup.add("statictext", undefined, "Ending Page:");
    var endPageInput = pageRangeGroup.add("edittext", undefined, "");
    endPageInput.characters = 6;

    // ---------- Buttons ----------
    var buttonGroup = settingsDialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";

    var cancelButton = buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });
    var okButton = buttonGroup.add("button", undefined, "Continue", { name: "ok" });

    var result = null;

    okButton.onClick = function () {
        var libPath = (controlsLibPath.text || "").replace(/^\s+|\s+$/g, "");
        if (!libPath) {
            alert("Please choose a page controls library file.");
            return;
        }

        var libFile = new File(libPath);
        if (!libFile.exists) {
            alert("Library file not found:\n" + libPath);
            return;
        }

        var docPages = app.activeDocument.pages;

        var startPageText = (startPageInput.text || "").replace(/^\s+|\s+$/g, "");
        var endPageText = (endPageInput.text || "").replace(/^\s+|\s+$/g, "");

        var startIndex = resolvePageIndex(app.activeDocument, startPageText, 0);
        var endIndex = resolvePageIndex(app.activeDocument, endPageText, docPages.length - 1);

        //-- remember for next run
        writeRememberedPath(pageControlsLabelKey, libFile.fsName);

        result = {
            libFile: libFile,
            startPage: startIndex,
            endPage: endIndex
        };

        settingsDialog.close(1);
    };

    cancelButton.onClick = function () {
        settingsDialog.close(0);
        exit(0);
    };

    settingsDialog.center();
    settingsDialog.show();

    function resolvePageIndex(doc, input, fallbackIndex) {
        var txt = (input || "").replace(/^\s+|\s+$/g, "");

        // Blank → use fallback
        if (txt === "") return fallbackIndex;

        // Numeric → convert to index
        if (/^\d+$/.test(txt)) {
            var idx = parseInt(txt, 10) - 1;
            if (idx >= 0 && idx < doc.pages.length) {
                return idx;
            }
        }

        // Try page name
        var page = doc.pages.itemByName(txt);
        if (page.isValid) {
            return page.documentOffset; // <-- key property
        }

        return -1; // invalid
    }

    function readRememberedPath(key) {
        try {
            var s = app.extractLabel(key);
            return (s || "").replace(/^\s+|\s+$/g, "");
        } catch (e) {
            return "";
        }
    }

    function writeRememberedPath(key, path) {
        try {
            app.insertLabel(key, path);
            return true;
        } catch (e) {
            return false;
        }
    }

    return result;
}


function getOrCreatePageDestination(doc, page, destName) {
    var i, dest;

    for (i = 0; i < doc.hyperlinkPageDestinations.length; i++) {
        dest = doc.hyperlinkPageDestinations[i];
        try {
            if (dest.destinationPage === page) {
                return dest;
            }
        } catch (_) {}
    }

    return doc.hyperlinkPageDestinations.add(page, { name: destName });
}

function addPageLink(doc, pageItem, destinationPage, linkBaseName) {
    var dest = getOrCreatePageDestination(doc, destinationPage, "Dest_" + destinationPage.name);
    var source = doc.hyperlinkPageItemSources.add(pageItem);
    doc.hyperlinks.add(source, dest, { name: linkBaseName });
}

//-----------------------------------------------------------------------------------------------------------
//-- set preferences to optimize environment for running scripts efficiently
function setAppPrefs() {
    oldRedrawPrefs = app.scriptPreferences.enableRedraw;
    oldPreflightPrefs = app.preflightOptions.preflightOff;
    oldInteractionPrefs = app.scriptPreferences.userInteractionLevel;

    app.scriptPreferences.enableRedraw = false;
    app.preflightOptions.preflightOff = true;
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
}

//-----------------------------------------------------------------------------------------------------------
//-- restore original preferences
function restoreAppPrefs() {
    try { app.scriptPreferences.enableRedraw = oldRedrawPrefs; } catch (_) {}
    try { app.preflightOptions.preflightOff = oldPreflightPrefs; } catch (_) {}
    try { app.scriptPreferences.userInteractionLevel = oldInteractionPrefs; } catch (_) {}
}













