/**

@title Multiple Find/Replace
@author Mitchell Ray
@version 1.0.0

Based on the GREP Query Runner script by Peter Kahrel
http://www.kahrel.plus.com/indesign/grep_query_runner.html

*/
#targetengine "session";

// Setup
var w = new Window("palette", "Multiple Find/Replace");
var scopePanel = w.add("panel", undefined, "Run queries on");
var scopeGroup = scopePanel.add("group");
var scopeSelection;

scopeGroup.add("radiobutton", undefined, "Selected Objects");
scopeGroup.add("radiobutton", undefined, "Document");

scopeGroup.children[1].value = true; // Default to 'Document'

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
    
    // Create dummy separator file
    var dummy = File(queryFolder + "----------.xml");
    dummy.open('w');
    dummy.write('');
    dummy.close();
    
    var list = findQueriesSub(appFolder);
    list = list.concat(findQueriesSub(queryFolder));
    for (var i = list.length - 1; i >= 0; i--) {
        list[i] = decodeURI(list[i].name.replace('.xml', ''));
    }

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
    var newSelections;

    if (scope == "Selected Objects") {
        warnSelectedObjects(thisDocument);        
        
        // Loop through each selected object
        for (n = 0; n < selections.length; n++) {
            app.select(selections[n].allPageItems, SelectionOptions.ADD_TO);
        }

        // Load our new selection
        newSelections = thisDocument.selection;

        subTextQueries(queries, newSelections);

        //Restore previous selection
        thisDocument.select(selections);
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
            var textProgress = new Window('palette', "Running " + q.textMode.length + " Text Queries", undefined, {closeButton:false});
            textProgress.progressbar = textProgress.add('progressbar', undefined, 0, q.textMode.length);
            textProgress.progressbar.preferredSize.width = 300;
            textProgress.show();

            // Loop through each selected query
            for (var i = 0; i < q.textMode.length; i++) {
                textProgress.progressbar.value = i+1;
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
        }
    }
}

function runGrepQueries() {
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

        subGrepQueries(queries, newSelections);

        //Restore previous selection
        thisDocument.select(selections);
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
            var grepProgress = new Window('palette', "Running " + q.grepMode.length + " GREP Queries", undefined, {closeButton:false});
            grepProgress.progressbar = grepProgress.add('progressbar', undefined, 0, q.grepMode.length);
            grepProgress.progressbar.preferredSize.width = 300;
            grepProgress.show();

            // Loop through each selected query
            for (var i = 0; i < q.grepMode.length; i++) {
                grepProgress.progressbar.value = i+1;
                app.loadFindChangeQuery(q.grepMode[i], SearchModes.grepSearch);
                app.findChangeGrepOptions.includeMasterPages = true;

                if (typeof s !== "undefined" && s.length > 0) {
                    for (x = 0; x < s.length; x++) {
                        try {
                            if (s[x] instanceof TextFrame) {
                                s[x].changeGrep();
                            }
                        } catch(e) {
                            // object was not a textframe and threw an exception
                        }
                    }
                } else if (scope == "Document") {
                    thisDocument.changeGrep();
                }
            }
        }
    }
}

function runGlyphQueries() {
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

        subGlyphQueries(queries, newSelections);

        //Restore previous selection
        thisDocument.select(selections);
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
            var glyphProgress = new Window('palette', "Running " + q.glyphMode.length + " Glyph Queries", undefined, {closeButton:false});
            glyphProgress.progressbar = glyphProgress.add('progressbar', undefined, 0, q.glyphMode.length);
            glyphProgress.progressbar.preferredSize.width = 300;
            glyphProgress.show();

            // Loop through each selected query
            for (var i = 0; i < q.glyphMode.length; i++) {
                glyphProgress.progressbar.value = i+1;
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
            var objectProgress = new Window('palette', "Running " + q.objectMode.length + " Object Queries", undefined, {closeButton:false});
            objectProgress.progressbar = objectProgress.add('progressbar', undefined, 0, q.objectMode.length);
            objectProgress.progressbar.preferredSize.width = 300;
            objectProgress.show();

            // Loop through each selected query
            for (var i = 0; i < q.objectMode.length; i++) {
                objectProgress.progressbar.value = i+1;
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
        }
    }
}

//
// EVENTS
//

buttons.textMode.onClick = function(){app.doScript("runTextQueries()", undefined, undefined, UndoModes.ENTIRE_SCRIPT, "Multiple Find/Replace (Text)")};
buttons.grepMode.onClick = function(){app.doScript("runGrepQueries()", undefined, undefined, UndoModes.ENTIRE_SCRIPT, "Multiple Find/Replace (GREP)")};
buttons.glyphMode.onClick = function(){app.doScript("runGlyphQueries()", undefined, undefined, UndoModes.ENTIRE_SCRIPT, "Multiple Find/Replace (Glyph)")};
buttons.objectMode.onClick = function(){app.doScript("runObjectQueries()", undefined, undefined, UndoModes.ENTIRE_SCRIPT, "Multiple Find/Replace (Object)")};
