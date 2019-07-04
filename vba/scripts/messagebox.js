/// <reference path="../../typings/globals/jquery/index.d.ts" />
var elemNum = 6;

var placeHolder = "QWQWQWQW";
var pref = '';
//virtualtype,long,short,init,declaration,constantinit,precommand


var dictCases = {
    0: ['    Case vbOK\n        \'code', '', '', '', '', '', ''],
    1: ['    Case vbOK\n        \'code', '    Case vbCancel\n        \'code', '', '', '', '', ''],
    2: ['', '', '    Case vbAbort\n        \'code', '    Case vbRetry\n        \'code', '    Case vbIgnore\n        \'code', '', ''],
    3: ['', '    Case vbCancel\n        \'code', '', '', '', '    Case vbYes\n        \'code', '    Case vbNo\n        \'code'],
    4: ['', '', '', '', '', '    Case vbYes\n        \'code', '    Case vbNo\n        \'code'],
    5: ['', '    Case vbCancel\n        \'code', '', '    Case vbRetry\n        \'code', '', '', ''],
}


var dictPredef = {
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
    if ($('#VBPredefMsg').val() != '-') {
        $('#MsgVar').val(dictPredef[$('#VBPredefMsg').val()][0]);
        $('#MsgPromt').val(dictPredef[$('#VBPredefMsg').val()][1]);
        $('#MsgTitle').val(dictPredef[$('#VBPredefMsg').val()][2]);
        $('#VBButton').val(dictPredef[$('#VBPredefMsg').val()][3]);
        $('#VBIcon').val(dictPredef[$('#VBPredefMsg').val()][4]);
    }
    if ($('#ShortPrefix').is(':checked')) {
        pref = 'mg';
    } else {
        pref = 'msg';
    };
    // for (i = 1; i < elemNum + 1; i++) {
    //     // var name = document.getElementById('Name' + i.toString()).value;
    var name = $('#MsgVar').val();
    if (name != '') {
        var declarations = 'Dim ' + pref + capitalizeFirstLetter(name) + ' As Integer\n\n';
        declarations = declarations + pref + capitalizeFirstLetter(name) + ' = MsgBox("' + $('#MsgPromt').val() + '"';
        var bVal = parseInt($('#VBButton').val());
        var buttonVal = bVal + parseInt($('#VBIcon').val());
        declarations = declarations + ', ' + buttonVal
        var msgTitle = $('#MsgTitle').val();
        if (msgTitle != '') {
            declarations = declarations + ', "' + msgTitle + '"';
        }
        declarations = declarations + ')\n';
        //     var scope = $('#Scope' + i).val()
        //     var type = $('#Type' + i).val()

        declarations = declarations + 'Select Case ' + pref + capitalizeFirstLetter(name) + '\n';
        if (dictCases[bVal][0] != '') {
            declarations = declarations + dictCases[bVal][0] + '\n';
        }
        if (dictCases[bVal][1] != '') {
            declarations = declarations + dictCases[bVal][1] + '\n';
        }
        if (dictCases[bVal][2] != '') {
            declarations = declarations + dictCases[bVal][2] + '\n';
        }
        if (dictCases[bVal][3] != '') {
            declarations = declarations + dictCases[bVal][3] + '\n';
        }
        if (dictCases[bVal][4] != '') {
            declarations = declarations + dictCases[bVal][4] + '\n';
        }
        if (dictCases[bVal][5] != '') {
            declarations = declarations + dictCases[bVal][5] + '\n';
        }

        declarations = declarations + 'End Select'
    }

    $('#Code').val(declarations);
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