/// <reference path="../../typings/globals/jquery/index.d.ts" />
var elemNum = 6;

var placeHolder = "QWQWQWQW";
var pref = '';
//virtualtype,long,short,init,declaration,constantinit,precommand


var dictCases = {
    0: ['    Case IDOK\n        \'code'],
    1: ['    Case IDOK\n        \'code\n    Case IDCANCEL\n        \'code'],
    2: ['    Case IDABORT\n        \'code\n    Case IDRETRY\n        \'code\n    Case IDIGNORE\n        \'code'],
    3: ['    Case IDYES\n        \'code\n    Case IDNO\n        \'code\n    Case IDCANCEL\n        \'code'],
    4: ['    Case IDYES\n        \'code\n    Case IDNO\n        \'code'],
    5: ['    Case IDRETRY\n        \'code\n    Case IDCANCEL\n        \'code'],
}


var dictPredef = {
    '-': ['', '', '', 0, 0],
    'msgConfirmation': ['confirm', 'Are you sure?', 'Confirm action', 4, 32],
    'msgMessage': ['message', 'This is the message', 'Message', 0, 64]
}
var prefIndex = 1;

$(document).ready(function () {
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
    var initvals = '';
    var i = 1;
    if ($('#ShortPrefix').is(':checked')) {
        pref = 'mg';
    } else {
        pref = 'msg';
    };
    var name = $('#MsgVar').val();
    if (name != '') {
        var declarations = 'Dim ' + pref + capitalizeFirstLetter(name) + ' As Integer\n\n';
        declarations = declarations + pref + capitalizeFirstLetter(name) + ' = MsgBox ("' + $('#MsgPromt').val() + '"';
        var bVal = parseInt($('#VBButton').val());
        var buttonVal = bVal + parseInt($('#VBIcon').val());
        declarations = declarations + ', ' + buttonVal
        var msgTitle = $('#MsgTitle').val();
        if (msgTitle != '') {
            declarations = declarations + ', "' + msgTitle + '"';
        }
        declarations = declarations + ')\n';
        declarations = declarations + 'Select Case ' + pref + capitalizeFirstLetter(name) + '\n';
        declarations = declarations + dictCases[bVal][0] + '\n';
        declarations = declarations + 'End Select'
    }

    $('#Code').val(declarations);
}

/** Change field values
 */
function predefChange() {
    $('#MsgVar').val(dictPredef[$('#VBPredefMsg').val()][0]);
    $('#MsgPromt').val(dictPredef[$('#VBPredefMsg').val()][1]);
    $('#MsgTitle').val(dictPredef[$('#VBPredefMsg').val()][2]);
    $('#VBButton').val(dictPredef[$('#VBPredefMsg').val()][3]);
    $('#VBIcon').val(dictPredef[$('#VBPredefMsg').val()][4]);
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