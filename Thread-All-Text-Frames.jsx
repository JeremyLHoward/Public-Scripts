/* =======================================================================================================
Written by: Jeremy Howard, 2026
https://www.linkedin.com/in/howarddesigns/
-----------------------------------------------
SCRIPT OVERVIEW
-----------------------------------------------
Script iterates through all pages in the active InDesign document, collects all text frames
on each page, sorts them from top to bottom, and threads them together into one continuous story.

Frames are processed in document order. The last text frame on each page is threaded to the
first text frame on the following page.

=========================== OPTIONAL FRAME BREAKS ===========================
To preserve the contents of each frame after threading, the script can insert a Frame Break at
the end of every text frame by changing the "addFrameBreaks" variable to true.
=============================================================================

Script Processing Steps:
- Checks that an InDesign document is open
- Iterates through each page in the active document
- Collects all text frames on each page
- Sorts the text frames from top to bottom
        if two frames share the same vertical position, they are sorted from left to right.
- Combines all page frames into a single array
- Checks whether any collected text frames are already threaded
        if any frame is already part of a thread, the script stops without making changes.
- IF "addFrameBreaks" is set to true then...
        inserts a Frame Break at the end of every text frame except the final frame
- Threads each text frame to the next frame
        the final frame on one page is threaded to the first frame on the next page.
======================================================================================================= */


#target "InDesign"

//-- Set to true to insert a frame break at the end of each frame's contents
var addFrameBreaks = false;

(function () {
    if (app.documents.length === 0) {
        alert("Please open an InDesign document first.");
        return;
    }

    app.doScript(
        main,
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "Thread Text Frames"
    );

    function main() {
        var doc = app.activeDocument;
        var orderedFrames = [];

        //-- Collect frames page by page
        for (var p = 0; p < doc.pages.length; p++) {
            var page = doc.pages[p];
            var pageFrames = page.textFrames.everyItem().getElements();

            //-- Sort frames on this page from top to bottom
            pageFrames.sort(function (a, b) {
                var aBounds = a.geometricBounds;
                var bBounds = b.geometricBounds;

                if (aBounds[0] !== bBounds[0]) {
                    return aBounds[0] - bBounds[0];
                }

                return aBounds[1] - bBounds[1];
            });

            //-- Add this page's frames to the document-wide array
            for (var f = 0; f < pageFrames.length; f++) {
                orderedFrames.push(pageFrames[f]);
            }
        }

        if (orderedFrames.length < 2) {
            alert("Fewer than two text frames were found.");
            return;
        }

        // Stop if any frames are already threaded
        for (var i = 0; i < orderedFrames.length; i++) {
            if (
                orderedFrames[i].previousTextFrame !== null ||
                orderedFrames[i].nextTextFrame !== null
            ) {
                alert(
                    "One or more text frames are already threaded.\r" +
                    "No changes were made."
                );
                return;
            }
        }

        if (addFrameBreaks == true) {
            for (var j = 0; j < orderedFrames.length - 1; j++) {
                orderedFrames[j].insertionPoints[-1].contents =
                    SpecialCharacters.FRAME_BREAK;
            }
        }

        //-- Thread all of the frames, in order
        for (var k = 0; k < orderedFrames.length - 1; k++) {
            orderedFrames[k].nextTextFrame = orderedFrames[k + 1];
        }

        alert(orderedFrames.length + " text frames were threaded.");
    }
})();