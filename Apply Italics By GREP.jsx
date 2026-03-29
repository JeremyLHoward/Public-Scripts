/* =======================================================================================================
-----------------------------------------------
SCRIPT OVERVIEW
-----------------------------------------------
Script iterates through grep statements that are stored in a csv file, performs a "find" 
by GREP and applies a character style named "EN Italic" to each found occurrence.

The name of the character style can be altered by changing the value for the 
"italic_char_style_name" variable near the top of the script.

Script Processing Steps:
- Checks active documents and determines if running on document or book
        This will work on the active document or the active book. If you have a document 
        and a book open, it will ask you which one you want to operate on. When operating 
        on a book, it will process each indd file in the book.
- Prompts user to choose a csv file that contains grep statements
- Reads csv and compiles a list of GREP statements
- For each target doc, check for "EN Italic" character style. If it does not exist then create it.
- Performs "findGrep()" for each grep statement derived from the csv
- Applies "EN Italic" style to any hits from the find
======================================================================================================= */


#targetengine "main"

var scriptName = "Apply Italics By GREP"

//-- calls "main()" - so the script is wrapped in a single undo
app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, scriptName);

function main() {

    var italic_char_style_name = "EN Italic"

    var scriptTarget, book, doc, oldRedrawPrefs, oldPreflightPrefs, oldInteractionPrefs;
    var bookOpen = false;
    var docOpen = false;
    var totalMatches = 0;
    var errors = [];
    var files = [];

    //-----------------------------------------------------------------------------------------------------------
    //-- determine if running on book or doc
    try {
        bookOpen = app.activeBook !== null;
        if (bookOpen) book = app.activeBook;
    } catch (noBook) {
        bookOpen = false;
    }

    try {
        docOpen = app.documents.length > 0 && app.activeDocument !== null;
        if (docOpen) doc = app.activeDocument;
    } catch (noDoc) {
        docOpen = false;
    }

    if (!bookOpen && !docOpen) {
        alert("Error!\rThere is no active book and no active document.");
        return;
    }

    if (docOpen && bookOpen) {
        scriptTarget = askBookOrDoc();
        if (!scriptTarget) return;
    } else {
        scriptTarget = docOpen ? "document" : "book";
    }
    //-----------------------------------------------------------------------------------------------------------

    //-----------------------------------------------------------------------------------------------------------
    //-- prompt user to choose grep csv, create grepList
    var csvFile = File.openDialog("Select a CSV file containing GREP expressions", "*.csv");
    if (!csvFile) return;

    var grepList = readGrepListFromCSV(csvFile);
    if (!grepList || grepList.length === 0) {
        alert("No GREP expressions were found in the selected file.");
        return;
    }
    //-----------------------------------------------------------------------------------------------------------


    setAppPrefs();


    //-----------------------------------------------------------------------------------------------------------
    //-- build target file array
    if (scriptTarget == "document") {
        files.push(doc.fullName);
    } else {
        var contents = book.bookContents;

        for (var i = 0; i < contents.length; i++) {
            if (contents[i].fullName) {
                files.push(contents[i].fullName);
            }
        }
    }
    //-----------------------------------------------------------------------------------------------------------

    //-----------------------------------------------------------------------------------------------------------
    //-- process docs in file array
    for (var f = 0; f < files.length; f++) {
        if (scriptTarget == 'document') {
            var currentDoc = app.activeDocument;
        } else {
            var currentDoc = app.open(files[f], false);
        }

        var italStyle = getItalStyle(currentDoc, italic_char_style_name);

        for (var i = 0; i < grepList.length; i++) {
            var grepPattern = grepList[i];

            if (!grepPattern || /^\s*$/.test(grepPattern)) continue;

            try {
                app.findGrepPreferences = NothingEnum.nothing;
                app.changeGrepPreferences = NothingEnum.nothing;

                app.findGrepPreferences.findWhat = grepPattern;

                var foundItems = currentDoc.findGrep();
                if (foundItems.length > 0) {
                    for (var j = 0; j < foundItems.length; j++) {
                        foundItems[j].appliedCharacterStyle = italStyle;
                    }
                    totalMatches += foundItems.length;
                }
            } catch (findErr) {
                errors.push("Pattern " + (i + 1) + ": " + grepPattern + "\r" + findErr);
            }
        }

        if (scriptTarget == 'book') currentDoc.close(SaveOptions.YES);
        
    }

    restoreAppPrefs();

    var msg = "Done.\r\rTotal matches styled: " + totalMatches;
    if (errors.length > 0) {
        msg += "\r\rSome patterns failed:\r\r" + errors.join("\r\r");
    }

    alert(msg);

    //==========================================================================================
    //-- FUNCTIONS FOR FUNCTIONING
    //==========================================================================================
    
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
        //app.findGrepPreferences = NothingEnum.nothing;
        //app.changeGrepPreferences = NothingEnum.nothing;
    }

    function askBookOrDoc() {
        var dlg = new Window("dialog", "Choose Target");
        dlg.orientation = "column";
        dlg.alignChildren = "left";

        dlg.add("statictext", undefined, "Both an active book and an active document were found.");
        dlg.add("statictext", undefined, "Choose the target for the script:");

        var panel = dlg.add("panel", undefined, "Target");
        panel.orientation = "column";
        panel.alignChildren = "left";

        var rbBook = panel.add("radiobutton", undefined, "All documents in active book");
        var rbDoc = panel.add("radiobutton", undefined, "Active document only");
        rbDoc.value = true;

        var btns = dlg.add("group");
        btns.alignment = "right";

        btns.add("button", undefined, "Cancel", {name:"cancel"});
        btns.add("button", undefined, "OK", {name:"ok"});

        if (dlg.show() != 1) return null;

        return rbBook.value ? "book" : "document";
    }

    function readGrepListFromCSV(fileObj) {
        var lines = [];
        fileObj.encoding = "UTF-8";

        if (!fileObj.open("r")) {
            alert("Could not open the selected file.");
            return lines;
        }

        while (!fileObj.eof) {
            var line = fileObj.readln();

            if (line === null || line === undefined) {
                continue;
            }

            line = trim(line);

            if (line === "") {
                continue;
            }

            if (line.charAt(0) === '"' && line.charAt(line.length - 1) === '"') {
                line = line.substring(1, line.length - 1);
                line = line.replace(/""/g, '"');
            }

            lines.push(line);
        }

        fileObj.close();
        return lines;
    }

    function getItalStyle(doc, styleName) {
        var style;

        try {
            style = doc.characterStyles.itemByName(styleName);
            var testName = style.name;
        } catch (e) {
            style = null;
        }

        if (!style) {
            var noneStyle = doc.characterStyles.itemByName("[None]");
            style = doc.characterStyles.add({
                name: styleName,
                basedOn: noneStyle
            });

            style.fontStyle = "Italic";
        }

        return style;
    }

    function trim(s) {
        if (s === null || s === undefined) return "";
        return String(s).replace(/^\s+|\s+$/g, "");
    }
}
