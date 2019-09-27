// / <reference path="../../typings/globals/jquery/index.d.ts" />
var elemNum = 6;

var pref = '';
// Virtualtype,long,short,init,declaration,constantinit,precommand

// prettier-ignore
var dictCases = {
  0: ['    Case vbOK\n        \'code'],
  1: ['    Case vbOK\n        \'code\n    Case vbCancel\n        \'code'],
  2: ['    Case vbAbort\n        \'code\n    Case vbRetry\n        \'code\n    Case vbIgnore\n        \'code'],
  3: ['    Case vbYes\n        \'code\n    Case vbNo\n        \'code\n    Case vbCancel\n        \'code'],
  4: ['    Case vbYes\n        \'code\n    Case vbNo\n        \'code'],
  5: ['    Case vbRetry\n        \'code\n    Case vbCancel\n        \'code'],
};

// prettier-ignore
var dictPredef = {'-': ['', '', '', 0, 0],
  'msgConfirmation': ['confirm', 'Are you sure?', 'Confirm action', 4, 32],
  'msgMessage': ['message', 'This is the message', 'Message', 0, 64]
};


$(document).ready(function() {
  $('#generate').click(function() {
    generateCode();
  });
  $('#clear').click(function() {
    resetNames();
  });
});

/**
 * Main function
 */
function generateCode() {
  var name = '';
  var declarations = '';
  var bVal;
  var buttonVal;
  var msgTitle;
  if ($('#ShortPrefix').is(':checked')) {
    pref = 'mg';
  } else {
    pref = 'msg';
  }
  name = $('#MsgVar').val();
  if (name !== '') {
    declarations =
      'Dim ' + pref + capitalizeFirstLetter(name) + ' As Integer\n\n';
    declarations =
      declarations +
      pref +
      capitalizeFirstLetter(name) +
      ' = MsgBox("' +
      $('#MsgPromt').val() +
      '"';
    bVal = parseInt($('#VBButton').val());
    buttonVal = bVal + parseInt($('#VBIcon').val());
    declarations = declarations + ', ' + buttonVal;
    msgTitle = $('#MsgTitle').val();
    if (msgTitle !== '') {
      declarations = declarations + ', "' + msgTitle + '"';
    }
    declarations += ')\n';
    declarations =
      declarations + 'Select Case ' + pref + capitalizeFirstLetter(name) + '\n';
    declarations = declarations + dictCases[bVal][0] + '\n';
    declarations += 'End Select';
  }

  $('#Code').val(declarations);
}

/**
 * Change field values
 * called from HTML
 */
function predefChange() {
  $('#MsgVar').val(dictPredef[$('#VBPredefMsg').val()][0]);
  $('#MsgPromt').val(dictPredef[$('#VBPredefMsg').val()][1]);
  $('#MsgTitle').val(dictPredef[$('#VBPredefMsg').val()][2]);
  $('#VBButton').val(dictPredef[$('#VBPredefMsg').val()][3]);
  $('#VBIcon').val(dictPredef[$('#VBPredefMsg').val()][4]);
}

/**
 * Capitalize the first letter of the text
 * @param  {string} text
 */
function capitalizeFirstLetter(text) {
  return text.charAt(0).toUpperCase() + text.slice(1);
}

/**
 * Reset #Name objects' names
 */
function resetNames() {
  for (i = 1; i < elemNum + 1; i++) {
    $('#Name' + i).val('');
  }
}
