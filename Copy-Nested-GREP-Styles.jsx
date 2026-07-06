//==========================================================================================================================================================//
/*

    Copy GREP Styles Between Paragraph Styles In A Document

    Use this script to copy nested GREP styles from one paragraph style to another in the same document.

    Author:   Jeremy Howard
    Email:    howarddesigns@live.com
    LinkedIn: https://www.linkedin.com/in/howarddesigns/

*/
//==========================================================================================================================================================//


//-- Ensure a document is open
if(app.documents.length == 0){
    alert('Please open a document before running this script.');
    exit(0);
}

var aDoc = app.activeDocument;
var allStyles = aDoc.allParagraphStyles;

//-- Build one filtered list of style objects (skipping "[No Paragraph Style]").
var styles = [];
for(var i=0;i<allStyles.length;i++){
    if(allStyles[i].name == '[No Paragraph Style]'){
        continue;
    }
    styles.push(allStyles[i]);
}


//-- Prompt user to choose the source and target styles
var mainDialog = new Window('dialog', 'Copy GREP Styles');
mainDialog.alignChildren = 'fill';
mainDialog.margins = 16;
mainDialog.spacing = 12;

//-- Source: a single style whose nested GREP rules will be copied.
var sourcePanel = mainDialog.add('panel {text: "Copy the nested GREP styles from:",alignChildren: "left",orientation: "column"}');
sourcePanel.margins = [16, 20, 16, 16];

var pStyleDropdown = sourcePanel.add('dropdownlist');
pStyleDropdown.preferredSize.width = 320;

//-- Target: one or more styles to receive the copied rules.
var targetPanel = mainDialog.add('panel {text: "Paste them into:",alignChildren: "left",orientation: "column"}');
targetPanel.margins = [16, 20, 16, 16];

var targetList = targetPanel.add('listbox', undefined, undefined, {scrolling: true, multiselect: true});
targetList.preferredSize.width = 320;
targetList.preferredSize.height = 320;
targetList.maximumSize.height = 500;

//-- Populate both controls from the same filtered list, tagging each item with its style object.
for(var i=0;i<styles.length;i++){
    var sourceItem = pStyleDropdown.add('item', styles[i].name);
    sourceItem.styleRef = styles[i];

    var targetItem = targetList.add('item', styles[i].name);
    targetItem.styleRef = styles[i];
}
pStyleDropdown.selection = 0;

var buttonGroup = mainDialog.add('group {alignment: "right",orientation: "row"}');
buttonGroup.add('button', undefined, 'Cancel', {name:'cancel'});
var continueButton = buttonGroup.add('button', undefined, 'Continue', {name:'ok'});

//-- Validate on Continue and keep the dialog open on a bad selection instead of forcing a relaunch.
continueButton.onClick = function(){
    var chosenSource = pStyleDropdown.selection;
    if(chosenSource == null){
        alert('Please select a source paragraph style.');
        return;
    }
    if(chosenSource.styleRef.nestedGrepStyles.length == 0){
        alert('The source style "' + chosenSource.styleRef.name + '" has no nested GREP styles to copy.');
        return;
    }
    var chosenTargets = targetList.selection;
    if(chosenTargets == null || chosenTargets.length == 0){
        alert('Please select at least one target paragraph style.');
        return;
    }
    mainDialog.close(1);
};

if(mainDialog.show() != 1){
    exit(0);
}

//-- The selections were already validated above, so just read the style objects back.
var currentSourceStyle = pStyleDropdown.selection.styleRef;
var sourceGrepStyles = currentSourceStyle.nestedGrepStyles;

var targetSelection = targetList.selection;
var targetStyles = [];
for(var t=0;t<targetSelection.length;t++){
    targetStyles.push(targetSelection[t].styleRef);
}


//-- Copy the GREP styles
var rulesAdded = 0;
var stylesTouched = 0;

//-- Paste the GREP styles
app.doScript(copyGrepStyles, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Copy GREP Styles Between Paragraph Styles');


//-- Report results
if(rulesAdded == 0){
    alert('No new rules were added. Every target style already contained these GREP styles.');
}else{
    var ruleWord = (rulesAdded == 1) ? 'rule' : 'rules';
    var styleWord = (stylesTouched == 1) ? 'style' : 'styles';
    alert('Done. Added ' + rulesAdded + ' GREP ' + ruleWord + ' to ' + stylesTouched + ' paragraph ' + styleWord + '.');
}


//-- Copy function
function copyGrepStyles(){
    for(var t=0;t<targetStyles.length;t++){
        var currentTargetStyle = targetStyles[t];
        var currentTargetGrepRules = currentTargetStyle.nestedGrepStyles;
        var addedToThisStyle = 0;

        for(var sg=0;sg<sourceGrepStyles.length;sg++){
            var thisGrepStyle = sourceGrepStyles[sg];
            var thisExpression = thisGrepStyle.grepExpression;
            var thisCharStyle = thisGrepStyle.appliedCharacterStyle;

            //-- Check whether this rule already exists on the target.
            //-- Character styles are compared by id, not by object reference, so re-running the script does not quietly add duplicates.
            var styleExists = false;
            for(var ctg=0;ctg<currentTargetGrepRules.length;ctg++){
                var thisTargetGrepRule = currentTargetGrepRules[ctg];
                if(thisExpression == thisTargetGrepRule.grepExpression && thisCharStyle.id == thisTargetGrepRule.appliedCharacterStyle.id){
                    styleExists = true;
                    break;
                }
            }

            if(styleExists == false){
                currentTargetStyle.nestedGrepStyles.add({grepExpression: thisExpression, appliedCharacterStyle: thisCharStyle});
                rulesAdded++;
                addedToThisStyle++;
            }
        }

        if(addedToThisStyle > 0){
            stylesTouched++;
        }
    }
}