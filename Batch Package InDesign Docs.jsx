#target "InDesign"

/* =======================================================================================================
Written by: Jeremy Howard, 2026
https://www.linkedin.com/in/howarddesigns/
======================================================================================================= */

(function () {
    var SCRIPT_NAME = "Batch Package Documents";
    var oldRedrawPrefs, oldPreflightPrefs, oldInteractionPrefs;

    if (app.documents.length === 0) {
        // folder mode still allowed
    }

    var pdfPresetNames = getPdfPresetNames();
    var ui = buildDialog(pdfPresetNames);
    if (!ui || ui.result !== 1) return;

    var opts = ui.options;

    if (opts.sourceMode === "open" && app.documents.length === 0) {
        alert("There are no open documents.");
        return;
    }

    if (opts.sourceMode === "folder" && !opts.sourceFolder) {
        alert("Please choose a source folder.");
        return;
    }

    if (!opts.outputFolder) {
        alert("Please choose an output folder.");
        return;
    }

    var filesToProcess = [];
    var i, doc, file;

    if (opts.sourceMode === "open") {
        for (i = 0; i < app.documents.length; i++) {
            filesToProcess.push(app.documents[i]);
        }
    } else {
        filesToProcess = getDocumentFilesFromFolder(
            opts.sourceFolder,
            opts.recursive,
            opts.includeIndt
        );

        if (filesToProcess.length === 0) {
            alert("No matching InDesign files were found in the selected folder.");
            return;
        }
    }

    var logFile = createLogFile(opts.outputFolder);
    logLine(logFile, "=== " + SCRIPT_NAME + " ===");
    logLine(logFile, "Started: " + new Date());
    logLine(logFile, "Source mode: " + opts.sourceMode);
    if (opts.sourceMode === "folder") {
        logLine(logFile, "Source folder: " + opts.sourceFolder.fsName);
        logLine(logFile, "Recursive: " + opts.recursive);
        logLine(logFile, "Include INDT: " + opts.includeIndt);
    }
    logLine(logFile, "Output folder: " + opts.outputFolder.fsName);
    logLine(logFile, "");

    var progress = buildProgressWindow(filesToProcess.length);

    var processed = 0;
    var skipped = 0;
    var failed = 0;
    var skippedItems = [];
    var failedItems = [];

    try {
        setAppPrefs();

        for (i = 0; i < filesToProcess.length; i++) {
            doc = null;
            file = null;

            try {
                if (opts.sourceMode === "open") {
                    doc = filesToProcess[i];
                    file = doc.saved ? doc.fullName : null;
                    updateProgress(progress, i + 1, filesToProcess.length, "Processing open document: " + doc.name);
                } else {
                    file = filesToProcess[i];
                    updateProgress(progress, i + 1, filesToProcess.length, "Opening: " + file.name);
                    doc = app.open(file, false);
                }

                if (!doc.saved || !doc.fullName) {
                    skipped++;
                    skippedItems.push(doc.name + " (document must be saved before packaging)");
                    logLine(logFile, "[SKIP] " + doc.name + " - document must be saved before packaging");
                    if (opts.sourceMode === "folder") {
                        try { doc.close(SaveOptions.NO); } catch (eClose1) {}
                    }
                    continue;
                }

                logLine(logFile, "----------------------------------------------------------------");
                logLine(logFile, "Document: " + doc.name);
                try {
                    logLine(logFile, "Full path: " + doc.fullName.fsName);
                } catch (_) {}

                reportDocumentIssues(doc, logFile);

                var baseName = getBaseName(doc.name);
                var packageFolder = getUniqueFolder(opts.outputFolder, sanitizeName(baseName));
                if (!packageFolder.exists) {
                    packageFolder.create();
                }

                if (opts.forceSave) {
                    updateProgress(progress, i + 1, filesToProcess.length, "Saving: " + doc.name);
                    doc.save();
                }

                updateProgress(progress, i + 1, filesToProcess.length, "Packaging: " + doc.name);

                doc.packageForPrint(
                    packageFolder,
                    opts.copyFonts,
                    opts.copyLinkedGraphics,
                    opts.copyProfiles,
                    opts.updateGraphics,
                    opts.includeHiddenLayers,
                    opts.ignorePreflightErrors,
                    opts.createReport,
                    opts.includeIdml,
                    opts.includePdf,
                    opts.includePdf ? opts.pdfPresetName : "",
                    opts.useHyphenationExceptionsOnly,
                    opts.versionComments,
                    opts.forceSave
                );

                processed++;
                logLine(logFile, "[OK] " + doc.name + " -> " + packageFolder.fsName);

            } catch (e) {
                failed++;
                var itemName = doc ? doc.name : (file ? file.name : "Unknown");
                var errText = errorToString(e);
                failedItems.push(itemName + " -> " + errText);
                logLine(logFile, "[FAIL] " + itemName + " -> " + errText);

            } finally {
                if (opts.sourceMode === "folder" && doc != null) {
                    try {
                        doc.close(SaveOptions.NO);
                    } catch (eClose2) {}
                }
            }
        }

    } finally {
        closeProgressWindow(progress);
        restoreAppPrefs();
    }

    logLine(logFile, "");
    logLine(logFile, "Finished: " + new Date());
    logLine(logFile, "Processed: " + processed);
    logLine(logFile, "Skipped: " + skipped);
    logLine(logFile, "Failed: " + failed);
    try { logFile.close(); } catch (eCloseLog) {}

    var msg = "";
    msg += "Packaging complete.\r\r";
    msg += "Processed: " + processed + "\r";
    msg += "Skipped: " + skipped + "\r";
    msg += "Failed: " + failed + "\r";
    msg += "\rLog file:\r" + logFile.fsName;

    if (skippedItems.length > 0) {
        msg += "\r\rSkipped:\r- " + skippedItems.join("\r- ");
    }

    if (failedItems.length > 0) {
        msg += "\r\rFailed:\r- " + failedItems.join("\r- ");
    }

    alert(msg);

    function buildDialog(pdfPresetNames) {
        var rememberedInput = readRememberedInputPath();
        var rememberedOutput = readRememberedOutputPath();

        var selectedSourceFolder = null;
        var selectedOutputFolder = null;

        if (rememberedInput) {
            var rememberedInputFolder = new Folder(rememberedInput);
            if (rememberedInputFolder.exists) {
                selectedSourceFolder = rememberedInputFolder;
            }
        }

        if (rememberedOutput) {
            var rememberedOutputFolder = new Folder(rememberedOutput);
            if (rememberedOutputFolder.exists) {
                selectedOutputFolder = rememberedOutputFolder;
            }
        }

        var w = new Window("dialog", SCRIPT_NAME);
        w.orientation = "column";
        w.alignChildren = "fill";

        var sourcePanel = w.add("panel", undefined, "Source");
        sourcePanel.orientation = "column";
        sourcePanel.alignChildren = "left";
        sourcePanel.margins = 12;

        var rbOpen = sourcePanel.add("radiobutton", undefined, "All open documents");
        var rbFolder = sourcePanel.add("radiobutton", undefined, "Files from a folder");
        rbOpen.value = true;

        var sourceFolderGroup = sourcePanel.add("group");
        sourceFolderGroup.orientation = "row";
        sourceFolderGroup.alignChildren = "center";
        sourceFolderGroup.add("statictext", undefined, "Source folder:");
        var sourceFolderText = sourceFolderGroup.add("edittext", undefined, "");
        sourceFolderText.characters = 40;
        sourceFolderText.enabled = false;
        var sourceFolderBtn = sourceFolderGroup.add("button", undefined, "Browse...");
        sourceFolderBtn.enabled = false;

        var folderOptionsGroup = sourcePanel.add("group");
        folderOptionsGroup.orientation = "column";
        folderOptionsGroup.alignChildren = "left";
        var cbRecursive = folderOptionsGroup.add("checkbox", undefined, "Include subfolders recursively");
        cbRecursive.value = false;
        cbRecursive.enabled = false;
        var cbIncludeIndt = folderOptionsGroup.add("checkbox", undefined, "Also include INDT files");
        cbIncludeIndt.value = false;
        cbIncludeIndt.enabled = false;

        var outputPanel = w.add("panel", undefined, "Output");
        outputPanel.orientation = "row";
        outputPanel.alignChildren = "center";
        outputPanel.margins = 12;

        outputPanel.add("statictext", undefined, "Target folder:");
        var outputFolderText = outputPanel.add("edittext", undefined, "");
        outputFolderText.characters = 40;
        outputFolderText.enabled = false;
        var outputFolderBtn = outputPanel.add("button", undefined, "Browse...");

        var optionsPanel = w.add("panel", undefined, "Package Options");
        optionsPanel.orientation = "column";
        optionsPanel.alignChildren = "left";
        optionsPanel.margins = 12;

        var cbCopyFonts = optionsPanel.add("checkbox", undefined, "Copy fonts");
        cbCopyFonts.value = true;

        var cbCopyLinks = optionsPanel.add("checkbox", undefined, "Copy linked graphics");
        cbCopyLinks.value = true;

        var cbCopyProfiles = optionsPanel.add("checkbox", undefined, "Copy profiles");
        cbCopyProfiles.value = false;

        var cbUpdateGraphics = optionsPanel.add("checkbox", undefined, "Update graphic links in package");
        cbUpdateGraphics.value = true;

        var cbIncludeHiddenLayers = optionsPanel.add("checkbox", undefined, "Include fonts and links from hidden and non-printing content");
        cbIncludeHiddenLayers.value = false;

        var cbIgnorePreflight = optionsPanel.add("checkbox", undefined, "Ignore preflight errors");
        cbIgnorePreflight.value = false;

        var cbCreateReport = optionsPanel.add("checkbox", undefined, "Create printing instructions/report");
        cbCreateReport.value = true;

        var cbIncludeIdml = optionsPanel.add("checkbox", undefined, "Include IDML");
        cbIncludeIdml.value = false;

        var pdfGroup = optionsPanel.add("group");
        pdfGroup.orientation = "row";
        pdfGroup.alignChildren = "center";

        var cbIncludePdf = pdfGroup.add("checkbox", undefined, "Include PDF");
        cbIncludePdf.value = false;
        pdfGroup.add("statictext", undefined, "Preset:");
        var ddPdfPreset = pdfGroup.add("dropdownlist", undefined, pdfPresetNames);
        ddPdfPreset.enabled = false;
        if (pdfPresetNames.length > 0) ddPdfPreset.selection = 0;

        var cbHyphenOnly = optionsPanel.add("checkbox", undefined, "Use document hyphenation exceptions only");
        cbHyphenOnly.value = false;

        var commentsGroup = optionsPanel.add("group");
        commentsGroup.orientation = "row";
        commentsGroup.alignChildren = "center";
        commentsGroup.add("statictext", undefined, "Version comments:");
        var etComments = commentsGroup.add("edittext", undefined, "");
        etComments.characters = 35;

        var cbForceSave = optionsPanel.add("checkbox", undefined, "Force save before packaging");
        cbForceSave.value = false;

        var buttonGroup = w.add("group");
        buttonGroup.alignment = "right";
        var okBtn = buttonGroup.add("button", undefined, "OK", {name:"ok"});
        var cancelBtn = buttonGroup.add("button", undefined, "Cancel", {name:"cancel"});

        if (selectedSourceFolder) {
            sourceFolderText.text = selectedSourceFolder.fsName;
            rbFolder.value = true;
            rbOpen.value = false;
        }

        if (selectedOutputFolder) {
            outputFolderText.text = selectedOutputFolder.fsName;
        }

        function refreshSourceState() {
            var folderMode = rbFolder.value;
            sourceFolderText.enabled = folderMode;
            sourceFolderBtn.enabled = folderMode;
            cbRecursive.enabled = folderMode;
            cbIncludeIndt.enabled = folderMode;
        }

        rbOpen.onClick = refreshSourceState;
        rbFolder.onClick = refreshSourceState;

        sourceFolderBtn.onClick = function () {
            var f = Folder.selectDialog("Choose the source folder");
            if (f) {
                selectedSourceFolder = f;
                sourceFolderText.text = f.fsName;
            }
        };

        outputFolderBtn.onClick = function () {
            var f = Folder.selectDialog("Choose the target output folder");
            if (f) {
                selectedOutputFolder = f;
                outputFolderText.text = f.fsName;
            }
        };

        cbIncludePdf.onClick = function () {
            ddPdfPreset.enabled = cbIncludePdf.value;
        };

        refreshSourceState();

        okBtn.onClick = function () {
            if (rbFolder.value && !selectedSourceFolder) {
                alert("Please choose a source folder.");
                return;
            }

            if (!selectedOutputFolder) {
                alert("Please choose a target output folder.");
                return;
            }

            if (cbIncludePdf.value && !ddPdfPreset.selection) {
                alert("Please choose a PDF preset.");
                return;
            }

            if (rbFolder.value && selectedSourceFolder) {
                writeRememberedInputPath(selectedSourceFolder.fsName);
            }
            if (selectedOutputFolder) {
                writeRememberedOutputPath(selectedOutputFolder.fsName);
            }

            w.options = {
                sourceMode: rbOpen.value ? "open" : "folder",
                sourceFolder: selectedSourceFolder,
                outputFolder: selectedOutputFolder,
                recursive: cbRecursive.value,
                includeIndt: cbIncludeIndt.value,
                copyFonts: cbCopyFonts.value,
                copyLinkedGraphics: cbCopyLinks.value,
                copyProfiles: cbCopyProfiles.value,
                updateGraphics: cbUpdateGraphics.value,
                includeHiddenLayers: cbIncludeHiddenLayers.value,
                ignorePreflightErrors: cbIgnorePreflight.value,
                createReport: cbCreateReport.value,
                includeIdml: cbIncludeIdml.value,
                includePdf: cbIncludePdf.value,
                pdfPresetName: (cbIncludePdf.value && ddPdfPreset.selection) ? ddPdfPreset.selection.text : "",
                useHyphenationExceptionsOnly: cbHyphenOnly.value,
                versionComments: etComments.text,
                forceSave: cbForceSave.value
            };
            w.close(1);
        };

        cancelBtn.onClick = function () {
            w.close(0);
        };

        var result = w.show();
        if (result !== 1) return null;

        return {
            result: result,
            options: w.options
        };
    }

    function getPdfPresetNames() {
        var arr = [];
        var i;
        for (i = 0; i < app.pdfExportPresets.length; i++) {
            arr.push(app.pdfExportPresets[i].name);
        }
        return arr;
    }

    function getDocumentFilesFromFolder(folder, recursive, includeIndt) {
        var results = [];
        scanFolder(folder);
        return results;

        function scanFolder(currentFolder) {
            var items = currentFolder.getFiles();
            var i, it, name;

            for (i = 0; i < items.length; i++) {
                it = items[i];

                if (it instanceof Folder) {
                    if (recursive) {
                        scanFolder(it);
                    }
                } else if (it instanceof File) {
                    name = it.name;
                    if (/\.indd$/i.test(name) || (includeIndt && /\.indt$/i.test(name))) {
                        results.push(it);
                    }
                }
            }
        }
    }

    function reportDocumentIssues(doc, logFile) {
        var issues = [];
        var links = [];
        var i, link, linkName, statusLabel;

        try {
            links = doc.links.everyItem().getElements();
        } catch (e) {
            logLine(logFile, "[WARN] Could not inspect links for " + doc.name + " -> " + errorToString(e));
            return;
        }

        if (!links || links.length === 0) {
            logLine(logFile, "[INFO] No links found.");
            return;
        }

        for (i = 0; i < links.length; i++) {
            try {
                link = links[i];
                linkName = safeLinkName(link);
                statusLabel = getLinkStatusLabel(link);

                if (statusLabel !== "NORMAL") {
                    issues.push("[LINK] " + statusLabel + " -> " + linkName);
                }
            } catch (eLink) {
                issues.push("[LINK] ERROR READING LINK -> " + errorToString(eLink));
            }
        }

        if (issues.length === 0) {
            logLine(logFile, "[INFO] No link issues found.");
        } else {
            logLine(logFile, "[WARN] Link issues found: " + issues.length);
            for (i = 0; i < issues.length; i++) {
                logLine(logFile, "    " + issues[i]);
            }
        }
    }

    function safeLinkName(link) {
        try {
            if (link.filePath) {
                return link.filePath;
            }
        } catch (_) {}

        try {
            if (link.name) {
                return link.name;
            }
        } catch (_) {}

        return "[Unknown Link]";
    }

    function getLinkStatusLabel(link) {
        var statusValue;

        try {
            statusValue = link.status;
        } catch (e) {
            return "STATUS_UNAVAILABLE";
        }

        try {
            if (statusValue === LinkStatus.NORMAL) return "NORMAL";
        } catch (_) {}

        try {
            if (statusValue === LinkStatus.LINK_MISSING) return "MISSING";
        } catch (_) {}

        try {
            if (statusValue === LinkStatus.LINK_OUT_OF_DATE) return "OUT_OF_DATE";
        } catch (_) {}

        try {
            if (statusValue === LinkStatus.LINK_EMBEDDED) return "EMBEDDED";
        } catch (_) {}

        try {
            if (statusValue === LinkStatus.LINK_INACCESSIBLE) return "INACCESSIBLE";
        } catch (_) {}

        try {
            return statusValue.toString();
        } catch (e2) {
            return "UNKNOWN_STATUS";
        }
    }

    function getUniqueFolder(parentFolder, baseName) {
        var candidate = new Folder(parentFolder.fsName + "/" + baseName);
        if (!candidate.exists) return candidate;

        var n = 2;
        while (true) {
            candidate = new Folder(parentFolder.fsName + "/" + baseName + "_" + n);
            if (!candidate.exists) return candidate;
            n++;
        }
    }

    function getBaseName(name) {
        return name.replace(/\.[^\.]+$/, "");
    }

    function sanitizeName(name) {
        return name.replace(/[\\\/\:\*\?\"\<\>\|]/g, "_");
    }

    function buildProgressWindow(maxValue) {
        var w = new Window("palette", SCRIPT_NAME + " - Progress");
        w.orientation = "column";
        w.alignChildren = "fill";
        w.margins = 12;

        w.statusText = w.add("statictext", undefined, "Starting...");
        w.statusText.characters = 60;

        w.progressBar = w.add("progressbar", undefined, 0, maxValue);
        w.progressBar.preferredSize.width = 400;

        w.counterText = w.add("statictext", undefined, "0 / " + maxValue);

        w.show();
        return w;
    }

    function updateProgress(w, value, maxValue, message) {
        try {
            w.statusText.text = message;
            w.progressBar.value = value;
            w.counterText.text = value + " / " + maxValue;
            w.update();
        } catch (e) {}
    }

    function closeProgressWindow(w) {
        try { w.close(); } catch (e) {}
    }

    function createLogFile(outputFolder) {
        var stamp = makeTimestamp();
        var f = new File(outputFolder.fsName + "/BatchPackageLog_" + stamp + ".txt");
        f.encoding = "UTF-8";
        f.open("w");
        return f;
    }

    function logLine(fileObj, text) {
        try {
            fileObj.writeln(text);
            fileObj.flush();
        } catch (e) {}
    }

    function makeTimestamp() {
        var d = new Date();
        return d.getFullYear() +
            pad2(d.getMonth() + 1) +
            pad2(d.getDate()) + "_" +
            pad2(d.getHours()) +
            pad2(d.getMinutes()) +
            pad2(d.getSeconds());
    }

    function pad2(n) {
        return (n < 10 ? "0" : "") + n;
    }

    function errorToString(e) {
        try {
            if (e && e.number !== undefined) {
                return e.toString() + " (Error " + e.number + ")";
            }
            return e.toString();
        } catch (err) {
            return "Unknown error";
        }
    }

    function readRememberedInputPath() {
        try {
            var s = app.extractLabel("batchPackageInputPath");
            return (s || "").replace(/^\s+|\s+$/g, "");
        } catch (e) {
            return "";
        }
    }

    function writeRememberedInputPath(path) {
        try {
            app.insertLabel("batchPackageInputPath", path);
            return true;
        } catch (e) {
            return false;
        }
    }

    function readRememberedOutputPath() {
        try {
            var s = app.extractLabel("batchPackageOutputPath");
            return (s || "").replace(/^\s+|\s+$/g, "");
        } catch (e) {
            return "";
        }
    }

    function writeRememberedOutputPath(path) {
        try {
            app.insertLabel("batchPackageOutputPath", path);
            return true;
        } catch (e) {
            return false;
        }
    }

    function setAppPrefs() {
        oldRedrawPrefs = app.scriptPreferences.enableRedraw;
        oldPreflightPrefs = app.preflightOptions.preflightOff;
        oldInteractionPrefs = app.scriptPreferences.userInteractionLevel;

        app.scriptPreferences.enableRedraw = false;
        app.preflightOptions.preflightOff = true;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    }

    function restoreAppPrefs() {
        try { app.scriptPreferences.enableRedraw = oldRedrawPrefs; } catch (_) {}
        try { app.preflightOptions.preflightOff = oldPreflightPrefs; } catch (_) {}
        try { app.scriptPreferences.userInteractionLevel = oldInteractionPrefs; } catch (_) {}
    }

})();
