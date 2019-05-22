/// <reference path="../../typings/globals/jquery/index.d.ts" />
var elemNum = 6;

//virtualtype,long,short,init,declaration,constantinit,precommand
var dictTypes = {
    'String': ['String', 'str', 's', ' = vbNullString', 'String', ' = "text"', ''],
    'Long': ['Long', 'lng', 'l', ' = 0', 'Long', ' = 0', ''],
    'Integer': ['Integer', 'int', 'i', ' = 0', 'Integer', ' = 0', ''],
    'Boolean': ['Boolean', 'bln', 'b', ' = False', 'Boolean', ' = False', ''],
    'Double': ['Double', 'dbl', 'd', ' = 0.1', 'Double', ' = 0.1', ''],
    'Date': ['Date', 'dat', 'dat', ' = CDate("04/22/2016 12:00 AM")', 'Date', ' = CDate("04/22/2016 12:00 AM")', ''],
    'Variant': ['Variant', 'vnt', 'v', ' = 0', 'Variant', ' = 0', ''],
    'Object': ['Object', 'obj', 'o', ' = Nothing', 'Object', ' = Nothing', 'Set '],
    'SheetName': ['SheetName', 'sh', 'sh', ' = "Munka1"', 'String', ' = "Munka1"', ''],
    'Worksheet': ['Worksheet', 'wsh', 'wsh', ' =  ThisComponent.CurrentController.ActiveSheet', 'Object', ' =  ThisComponent.CurrentController.ActiveSheet', 'Set '],
    'WorkbookName': ['WorkbookName', 'wb', 'wb', ' = "ThisBook"', 'String', ' = "ThisBook"', ''],
    'Workbook': ['Workbook', 'wbk', 'wbk', ' = ThisComponent', 'Object', ' = ThisComponent', 'Set '],
    'ColumnName': ['ColumnName', 'col', 'col', ' = "Header"', 'String', ' = "Header"', ''],
    'ColumnNumber': ['ColumnNumber', 'col', 'col', ' = 1', 'Long', ' = 1', ''],
    'Outlook': ['Outlook', 'oul', 'ou', ' = Nothing', 'Outlook', ' = Nothing', 'Set ']
};

//prefix,declaration,constdeclaration
var scopeTypes = {
    'Procedure': ['', 'Dim', ''],
    'Module': ["m", 'Private', 'Private'],
    'Global': ['g', 'Global', 'Global']
}
var prefIndex = 1;

$(document).ready(function () {
    var i = 1;
    for (i = 1; i < elemNum + 1; i++) {
        $("#tabla").find('tbody')
            .append($('<tr>')
                .append(
                    $('<td class="nameColumn"><input type="text" id="Name' + i +
                        '" />'))
                .append(
                    $('<td class="otherColumns"><select id="Dimension' + i +
                        '">'))
                .append(
                    $('<td class="otherColumns"><select id="Scope' + i +
                        '">'))
                .append(
                    $('<td class="otherColumns"><select id="Type' + i +
                        '">'))
            );
    }

    for (i = 1; i < elemNum + 1; i++) {
        $('#Dimension' + i)
            .append($('<option value="Normal">Normal</option>'))
            .append($('<option value="Constant">Constant</option>'))
            .append($('<option value="Array">Array</option>'));
    }

    for (i = 1; i < elemNum + 1; i++) {
        $('#Scope' + i)
            .append($('<option value="Procedure">Procedure</option>'))
            .append($('<option value="Module">Module</option>'))
            .append($('<option value="Global">Global</option>'));
    }

    for (i = 1; i < elemNum + 1; i++) {
        $('#Type' + i)
            .append($('<option value="String">String</option>'))
            .append($('<option value="Long">Long</option>'))
            .append($('<option value="Integer">Integer</option>'))
            .append($('<option value="Boolean">Boolean</option>'))
            .append($('<option value="Double">Double</option>'))
            .append($('<option value="Date">Date</option>'))
            .append($('<option value="Variant">Variant</option>'))
            .append($('<option value="Object">Object</option>'))
            .append($('<option value="SheetName">SheetName</option>'))
            .append($('<option value="Worksheet">Worksheet</option>'))
            .append($('<option value="WorkbookName">WorkbookName</option>'))
            .append($('<option value="Workbook">Workbook</option>'))
            .append($('<option value="ColumnName">ColumnName</option>'))
            .append($('<option value="ColumnNumber">ColumnNumber</option>'))
            .append($('<option value="Outlook">Outlook</option>'));
    }

    $('#generate').click(function () {
        generateCode();
    });
    $('#clear').click(function () {
        resetNames();
    });
});

/** Main function
 */
function generateCode() {
    var declarations = '';
    var initvals = '';
    var i = 1;
    if ($('#ShortPrefix').is(':checked')) {
        prefIndex = 2;
    } else {
        prefIndex = 1;
    };
    for (i = 1; i < elemNum + 1; i++) {
        // var name = document.getElementById('Name' + i.toString()).value;
        var name = $('#Name' + i).val();
        var dimension = $('#Dimension' + i).val()
        var scope = $('#Scope' + i).val()
        var type = $('#Type' + i).val()
        if (name != '') {
            declarations = declarations + getDeclars(name, dimension, scope, type) + '\n';
            initvals = initvals + getInitValues(name, dimension, scope, type) + '\n';
        }
    };
    declarations = declarations + '--------------------------\n'
    declarations = declarations + initvals;

    $('#Code').val(declarations);
}

/** Create declaration
 * @param  {string} name
 * @param  {string} dimension
 * @param  {string} scope
 * @param  {string} type
 */
function getDeclars(name, dimension, scope, type) {
    var declaration = '';

    if (dimension == 'Normal') {
        declaration = declaration + scopeTypes[scope][1] + ' ' + scopeTypes[scope][0] +
            dictTypes[type][prefIndex] + capitalizeFirstLetter(name) +
            ' As ' + dictTypes[type][4] + '\n';
    }
    if (dimension == 'Constant') {
        declaration = declaration + scopeTypes[scope][2] +
            ' ' + 'Const c' + scopeTypes[scope][0] + dictTypes[type][prefIndex] +
            capitalizeFirstLetter(name) + ' As ' + dictTypes[type][4] +
            dictTypes[type][5] + '\n';
    }
    return declaration;
}

/** Create initial values
 * @param  {string} name
 * @param  {string} dimension
 * @param  {string} scope
 * @param  {string} type
 */
function getInitValues(name, dimension, scope, type) {
    var declaration = '';
    if (dimension == 'Normal') {
        declaration = declaration + dictTypes[type][6] + scopeTypes[scope][0] +
            dictTypes[type][prefIndex] + capitalizeFirstLetter(name) + dictTypes[type][3] + '\n';
    }
    return declaration;
}

/** Capitalize the first letter of the text
 * @param  {string} text
 */
function capitalizeFirstLetter(text) {
    return text.charAt(0).toUpperCase() + text.slice(1);
}

/** Reset #Name objects' names
 */
function resetNames() {
    for (var i = 1; i < elemNum + 1; i++) {
        $('#Name' + i).val('');
    }
}