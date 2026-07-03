#target 'indesign'

/*
    GREP Style Report
    Generates a text report of the GREP styles applied within selected paragraph styles.

    Author:   Jeremy Howard
    Email:    howarddesigns@live.com
    LinkedIn: https://www.linkedin.com/in/howarddesigns/
 */

(function () {

    if (app.documents.length === 0) {
        alert('Please open a document before running this script.');
        return;
    }

    var doc        = app.activeDocument;
    var docName    = doc.name.replace(/\.indd$/i, '');
    var paraStyles = doc.allParagraphStyles;

    var styleNames = [];
    for (var i = 0; i < paraStyles.length; i++) {
        styleNames.push(paraStyles[i].name);
    }

    var chosenStyles = promptForStyles(styleNames);
    if (chosenStyles === null) { return; }

    //=============  BUILD THE REPORT  =============//
    var newLine    = '\r\n';
    var titleRule  = new Array(63).join('=');
    var lines      = [];

    for (var s = 0; s < chosenStyles.length; s++) {
        var style      = paraStyles[chosenStyles[s].index];
        var grepStyles = style.nestedGrepStyles;
        var base       = style.basedOn;
        var baseName   = (base && base.isValid) ? base.name : '[None]';

        if (s > 0) { lines.push(''); }
        lines.push(titleRule);
        lines.push('PARAGRAPH STYLE: ' + style.name);
        lines.push(titleRule);
        lines.push('Based On:   ' + baseName);
        lines.push('Font Name:  ' + style.appliedFont.name);
        lines.push('Font Style: ' + style.fontStyle);

        if (grepStyles.length === 0) {
            lines.push('');
            lines.push('\tNo GREP styles applied.');
        } else {
            for (var g = 0; g < grepStyles.length; g++) {
                lines.push('');
                lines.push('\tApplied Character Style: ' + grepStyles[g].appliedCharacterStyle.name);
                lines.push('\tGREP Expression:         ' + grepStyles[g].grepExpression);
            }
        }
    }

    //=============  WRITE TO THE DESKTOP  =============//
    var reportName    = docName + ' GREP Style Report.txt';
    var reportFile    = File('~/Desktop/' + reportName);

    reportFile.encoding = 'UTF-8';
    if (!reportFile.open('w')) {
        alert('Could not open the report file for writing.');
        return;
    }
    reportFile.write(lines.join(newLine) + newLine);
    reportFile.close();

    alert('GREP Report exported.\r\rLook for it on your desktop:\r' + reportName);


    //=============  DIALOG  =============//
    //-- Returns an array of selected ListItems, or null if cancelled / empty.
    function promptForStyles(names) {
        var stylePickerDialog = new Window('dialog', 'Source Paragraph Style', undefined, {resizeable: true});
        stylePickerDialog.preferredSize = [300, 500];
        stylePickerDialog.minimumSize   = [260, 300];
        stylePickerDialog.margins       = 23;
        stylePickerDialog.alignChildren = ['fill', 'fill'];

        var promptText = stylePickerDialog.add('statictext', undefined,
            'Select the style(s) to use when generating the GREP report\n(multiple selections allowed)',
            {multiline: true});
        promptText.alignment = ['fill', 'top'];

        var styleList = stylePickerDialog.add('listbox', undefined, names, {scrolling: true, multiselect: true});
        styleList.alignment     = ['fill', 'fill'];
        styleList.preferredSize = [350, 200];
        styleList.selection     = 0;

        var buttons = stylePickerDialog.add('group');
        buttons.orientation   = 'row';
        buttons.alignment     = ['right', 'bottom'];
        buttons.alignChildren = ['right', 'center'];
        buttons.add('button', undefined, 'Cancel',   {name: 'cancel'});
        buttons.add('button', undefined, 'Continue', {name: 'ok'});

        stylePickerDialog.onResizing = stylePickerDialog.onResize = function () { this.layout.resize(); };

        if (stylePickerDialog.show() !== 1) { return null; }
        if (!styleList.selection)  { return null; }

        return styleList.selection.length ? styleList.selection : [styleList.selection];
    }

}());
