/**

@title Multiple Find/Replace
@author Mitchell Ray
@version 1.0.3

Based on the GREP Query Runner script by Peter Kahrel
http://www.kahrel.plus.com/indesign/grep_query_runner.html

*/
#targetengine "session";

// Setup
var w = new Window("palette", "Multiple Find/Replace 1.0.3");
var scopePanel = w.add("panel", undefined, "Run queries on");
var scopeGroup = scopePanel.add("group");
var scopeSelection;

scopeGroup.add("radiobutton", undefined, "Selected Objects");
scopeGroup.add("radiobutton", undefined, "Document");

scopeGroup.children[0].value = true; // Default to 'Selected Objects'

var typesPanel = w.add("tabbedpanel");
typesPanel.alignChildren = ["fill", "fill"];

var searchTypeResults = {};
var tabs = {};
var searchTypes = [
    {key: "textMode", folder: "Text", tab: "Text"},
    {key: "grepMode", folder: "Grep", tab: "GREP"},
    {key: "glyphMode", folder: "Glyph", tab: "Glyph"},
    {key: "objectMode", folder: "Object", tab: "Object"}
];
var buttons = {};

for (i = 0; i < searchTypes.length; i++) {
    searchTypeResults[searchTypes[i].key] = {};
    searchTypeResults[searchTypes[i].key].results = [];
    searchTypeResults[searchTypes[i].key].folder = searchTypes[i].folder;
    tabs[searchTypes[i].key] = {};
    tabs[searchTypes[i].key].name = searchTypes[i].key;
    tabs[searchTypes[i].key].tab = typesPanel.add("tab", undefined, searchTypes[i].tab);
}

getAll();

//
// INIT
//

w.show();

//
// FUNCTIONS
//

function getScope() {
    for (var i = 0; i < scopeGroup.children.length; i++) {
        if (scopeGroup.children[i].value == true) {
            return scopeGroup.children[i].text;
        }
    }
}

//------------------------------------------------------------------------
// Some things to restore the previous session's selection

function scriptPath() {
    try {
        return app.activeScript;
    }
    catch (e) {
        return File(e.fileName);
    }
}

function saveData(obj) {
    var f = File(scriptPath().fullName.replace(/\.jsx$/, '.txt'));
    f.open('w');
    f.write(obj.toSource());
    f.close();
}

function getPrevious() {
    var f = File(scriptPath().fullName.replace(/\.jsx$/, '.txt'));
    var obj = {};
    if (f.exists) {
        f.open('r');
        var temp = f.read();
        f.close();
        obj = eval(temp);
    }
    return obj;
}

//------------------------------------------------------------------------
// Convert an array of ListItems to an array of strings

function getQueryNames(sel) {
    var arr = [];

    for (var i = 0; i < sel.length; i++) {
        arr.push(sel[i].text);
    }

    return arr;
}

//------------------------------------------------------------------------
// Create an array of GREP queries from the app and the user folder

function findQueries(queryType) {

    function findQueriesSub(dir) {
        var f = Folder(dir).getFiles('*.xml');
        return f;
    }

    var indesignFolder = function() {
        if ($.os.indexOf ('Mac') > -1) {
            return Folder.appPackage.parent;
        } else {
           return Folder.appPackage;
        }
    }

    var queryFolder = app.scriptPreferences.scriptsFolder.parent.parent + "/Find-Change Queries/" + searchTypeResults[queryType].folder + "/";
    var appFolder = indesignFolder() + '/Presets/Find-Change Queries/' + searchTypeResults[queryType].folder + '/' + $.locale;

    // Remove dummy separator file from old version
    var dummy = File(queryFolder + "----------.xml");
    dummy.remove('w');

    var list = findQueriesSub(appFolder);
    list = list.concat(findQueriesSub(queryFolder));
    for (var i = list.length - 1; i >= 0; i--) {
        list[i] = decodeURI(list[i].name.replace('.xml', ''));
    }

    list.sort();

    searchTypeResults[queryType].results = list;
}

function getAll() {

    var previous = getPrevious();

    for (n = 0; n < searchTypes.length; n++) {
        findQueries(searchTypes[n].key);

        tabs[searchTypes[n].key].list = tabs[searchTypes[n].key].tab.add('listbox', undefined, searchTypeResults[searchTypes[n].key].results, {multiselect: true});
        tabs[searchTypes[n].key].list.preferredSize = [300, 300];
        tabs[searchTypes[n].key].buttons = tabs[searchTypes[n].key].tab.add('group');
        buttons[searchTypes[n].key] = tabs[searchTypes[n].key].buttons.add('button', [0, 0, 200, 30], 'Run Selected ' + searchTypes[n].tab + ' Queries');
    }

    w.onShow = function () {
        if (previous.hasOwnProperty('location')) {
            w.location = previous.location;
        }

        if (previous.hasOwnProperty('queries')) {
            for (n = 0; n < searchTypes.length; n++) {
                if (typeof previous.queries[searchTypes[n].key] !== "undefined" && previous.queries[searchTypes[n].key].length > 0) {
                    for (var i = 0; i < previous.queries[searchTypes[n].key].length; i++) {
                        tabs[searchTypes[n].key].list.selection = tabs[searchTypes[n].key].list.find(previous.queries[searchTypes[n].key][i]);
                    }
                }
            }
        }
    }

}

function saveSelectedQueries() {
    var selectedQueries = {};

    for (i = 0; i < searchTypes.length; i++) {
        var selections = tabs[searchTypes[i].key].list.selection;

        if (selections !== null) {
            selectedQueries[searchTypes[i].key] = getQueryNames(selections);
        }
    }

    saveData({
        queries: selectedQueries,
        location: [w.location.x, w.location.y]
    });

    return selectedQueries;
}

function warnDocumentScope() {
    var myDialog = new Window("dialog", "Confirmation");
    myDialog.add("statictext", undefined, "Processing the entire document may take a while");
    var myDialogButtons = myDialog.add("group");
    myDialogButtons.alignment = "right";
    myDialogButtons.add("button", undefined, "Continue", {name: "ok"});
    myDialogButtons.add("button", undefined, "Cancel");

    if (myDialog.show() !== 1) {
        exit();
    }
}

function warnSelectedObjects(d) {
    var s = d.selection;

    //alert(app.selection[0].getElements()[0].constructor.name);

    if (s.length < 1) {
        var myDialog = new Window("dialog", "Warning");
        myDialog.add("statictext", undefined, "No objects were selected");
        var myDialogButtons = myDialog.add("group");
        myDialogButtons.alignment = "right";
        myDialogButtons.add("button", undefined, "Continue", {name: "cancel"});

        if (myDialog.show() !== 1) {
            exit();
        }
    }

    if (app.selection[0].getElements()[0].constructor.name == "InsertionPoint" || app.selection[0].getElements()[0].constructor.name == "Text") {
        var myDialog = new Window("dialog", "Warning");
        myDialog.add("statictext", undefined, "Please select objects instead of text");
        var myDialogButtons = myDialog.add("group");
        myDialogButtons.alignment = "right";
        myDialogButtons.add("button", undefined, "Continue", {name: "cancel"});

        if (myDialog.show() !== 1) {
            exit();
        }
    }
}

function runTextQueries() {
    var queries = saveSelectedQueries();
    var scope = getScope();
    var thisDocument = app.activeDocument;
    var selections = thisDocument.selection;
    var newSelections = [];

    if (scope == "Selected Objects") {
        warnSelectedObjects(thisDocument);

        newSelections = getAllObjectsInSelection(selections);
        subTextQueries(queries, newSelections);
    } else {
        warnDocumentScope();
        subTextQueries(queries);
    }

    /**
     *
     * @param q Queries selected to be run
     * @param s Selection of objects
     */
    function subTextQueries(q, s) {
        if (typeof q.textMode !== "undefined" && q.textMode.length > 0) {
            createLocalProgressBar("Running " + q.textMode.length + " Text Queries", q.textMode.length);

            // Loop through each selected query
            for (var i = 0; i < q.textMode.length; i++) {
                updateLocalProgressBar();

                app.loadFindChangeQuery(q.textMode[i], SearchModes.textSearch);
                app.findChangeTextOptions.includeMasterPages = true;

                if (typeof s !== "undefined" && s.length > 0) {
                    for (x = 0; x < s.length; x++) {
                        try {
                            if (s[x] instanceof TextFrame) {
                                s[x].changeText();
                            }
                        } catch(e) {
                            // object was not a textframe and threw an exception
                        }
                    }
                } else if (scope == "Document") {
                    thisDocument.changeText();
                }
            }

            destroyLocalProgressBar();
        }
    }
}

function runGrepQueries() {
    var queries = saveSelectedQueries();
    var scope = getScope();
    var thisDocument = app.activeDocument;
    var selections = thisDocument.selection;
    var newSelections = [];

    if (scope == "Selected Objects") {
        warnSelectedObjects(thisDocument);

        newSelections = getAllObjectsInSelection(selections);
        subGrepQueries(queries, newSelections);
    } else {
        warnDocumentScope();
        subGrepQueries(queries);
    }

    /**
     *
     * @param q Queries selected to be run
     * @param s Selection of objects
     */
    function subGrepQueries(q, s) {
        if (typeof q.grepMode !== "undefined" && q.grepMode.length > 0) {
            createLocalProgressBar("Running " + q.grepMode.length + " GREP Queries", q.grepMode.length);

            // Loop through each selected query
            for (var i = 0; i < q.grepMode.length; i++) {
                updateLocalProgressBar();

                app.loadFindChangeQuery(q.grepMode[i], SearchModes.grepSearch);
                app.findChangeGrepOptions.includeMasterPages = true;

                if (typeof s !== "undefined" && s.length > 0) {
                    for (x = 0; x < s.length; x++) {
                        if (s[x] instanceof TextFrame) {
                            s[x].changeGrep();
                        }
                    }
                } else if (scope == "Document") {
                    thisDocument.changeGrep();
                }
            }

            destroyLocalProgressBar();
        }
    }
}

function runGlyphQueries() {
    var queries = saveSelectedQueries();
    var scope = getScope();
    var thisDocument = app.activeDocument;
    var selections = thisDocument.selection;
    var newSelections = [];

    if (scope == "Selected Objects") {
        warnSelectedObjects(thisDocument);

        newSelections = getAllObjectsInSelection(selections);
        subGlyphQueries(queries, newSelections);
    } else {
        warnDocumentScope();
        subGlyphQueries(queries);
    }

    /**
     *
     * @param q Queries selected to be run
     * @param s Selection of objects
     */
    function subGlyphQueries(q, s) {
        if (typeof q.glyphMode !== "undefined" && q.glyphMode.length > 0) {
            createLocalProgressBar("Running " + q.glyphMode.length + " Glyph Queries", q.glyphMode.length);

            // Loop through each selected query
            for (var i = 0; i < q.glyphMode.length; i++) {
                updateLocalProgressBar();

                app.loadFindChangeQuery(q.glyphMode[i], SearchModes.glyphSearch);
                app.findChangeGlyphOptions.includeMasterPages = true;

                if (typeof s !== "undefined" && s.length > 0) {
                    for (x = 0; x < s.length; x++) {
                        try {
                            if (s[x] instanceof TextFrame) {
                                s[x].changeGlyph();
                            }
                        } catch(e) {
                            // object was not a textframe and threw an exception
                        }
                    }
                } else if (scope == "Document") {
                    thisDocument.changeGlyph();
                }
            }

            destroyLocalProgressBar();
        }
    }
}

function runObjectQueries() {
    var queries = saveSelectedQueries();
    var scope = getScope();
    var thisDocument = app.activeDocument;
    var selections = thisDocument.selection;
    var newSelections;

    if (scope == "Selected Objects") {
        warnSelectedObjects(thisDocument);

        // Loop through each selected object
        for (n = 0; n < selections.length; n++) {
            app.select(selections[n].allPageItems, SelectionOptions.ADD_TO);
        }

        // Load our new selection
        newSelections = thisDocument.selection;

        subObjectQueries(queries, newSelections);

        //Restore previous selection
        thisDocument.select(selections);
    } else {
        warnDocumentScope();
        subObjectQueries(queries);
    }

    /**
     *
     * @param q Queries selected to be run
     * @param s Selection of objects
     */
    function subObjectQueries(q, s) {
        if (typeof q.objectMode !== "undefined" && q.objectMode.length > 0) {
            createLocalProgressBar("Running " + q.objectMode.length + " Object Queries", q.objectMode.length);

            // Loop through each selected query
            for (var i = 0; i < q.objectMode.length; i++) {
                updateLocalProgressBar();

                app.loadFindChangeQuery(q.objectMode[i], SearchModes.objectSearch);
                app.findChangeObjectOptions.includeMasterPages = true;

                if (typeof s !== "undefined" && s.length > 0) {
                    for (x = 0; x < s.length; x++) {
                        s[x].changeObject();
                    }
                }  if (scope == "Document") {
                    thisDocument.changeObject();
                }
            }

            destroyLocalProgressBar();
        }
    }
}

/**
 *  Local task progress bar
 */
function createLocalProgressBar(message, total) {
    localProgressBar = new Window('palette', message, undefined, {closeButton:false});
    localProgressBar.progressbar = localProgressBar.add('progressbar', undefined, 0, total);
    localProgressBar.progressbar.preferredSize.width = 300;
    localProgressBar.show();
}

function updateLocalProgressBar() {
    localProgressBar.progressbar.value++;
    localProgressBar.update();
};

function destroyLocalProgressBar() {
    localProgressBar.close();
    // Prevents InDesign from losing focus
    indesign.active = true;
}

/**
 *  Function to recursively gather all pageItems, including nested (anchored) objects
 */
function getAllObjectsInSelection(selection) {
    var allPageItems = [];

    // Helper function to recursively explore pageItems and nested items
    function recurseItems(items) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];

            // Add the item itself to the collection if it's not already added
            allPageItems.push(item);

            // Check if the item is a group, which might contain other pageItems
            if (item.hasOwnProperty("allPageItems")) {
                recurseItems(item.allPageItems); // Recurse into the group's pageItems
            }
        }
    }

    // Start the recursion on the selected items
    recurseItems(selection);

    return allPageItems;
}

//
// EVENTS
//

buttons.textMode.onClick = function(){app.doScript("runTextQueries()", undefined, undefined, UndoModes.ENTIRE_SCRIPT, "Multiple Find/Replace (Text)")};
buttons.grepMode.onClick = function(){app.doScript("runGrepQueries()", undefined, undefined, UndoModes.ENTIRE_SCRIPT, "Multiple Find/Replace (GREP)")};
buttons.glyphMode.onClick = function(){app.doScript("runGlyphQueries()", undefined, undefined, UndoModes.ENTIRE_SCRIPT, "Multiple Find/Replace (Glyph)")};
buttons.objectMode.onClick = function(){app.doScript("runObjectQueries()", undefined, undefined, UndoModes.ENTIRE_SCRIPT, "Multiple Find/Replace (Object)")};
