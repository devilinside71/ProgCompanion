var elemNum = 6;
var placeHolder = 'QWQWQWQW';

// Virtualtype,long,short,init,declaration,constantinit,precommand

// prettier-ignore
var dictTypes = {
  String: ['String', 'str', 's', ' = vbNullString', 'String', ' = "text"', ''],
  Long: ['Long', 'lng', 'l', ' = 0', 'Long', ' = 0', ''],
  Integer: ['Integer', 'int', 'i', ' = 0', 'Integer', ' = 0', ''],
  Boolean: ['Boolean', 'bln', 'b', ' = False', 'Boolean', ' = False', ''],
  Double: ['Double', 'dbl', 'd', ' = 0.1', 'Double', ' = 0.1', ''],
  Date: ['Date', 'dat', 'dat', ' = CDate("04/22/2016 12:00 AM")', 'Date', ' = CDate("04/22/2016 12:00 AM")', ''],
  Variant: ['Variant', 'vnt', 'v', ' = 0', 'Variant', ' = 0', ''],
  Object: ['Object', 'obj', 'o', ' = Nothing', 'Object', ' = Nothing', 'Set '],
  SheetName: ['SheetName', 'sh', 'sh', ' = "Munka1"', 'String', ' = "Munka1"', ''],
  Worksheet: ['Worksheet', 'wsh', 'wsh', ' = ActiveSheet', 'Worksheet', ' = ActiveSheet', 'Set '],
  WorkbookName: ['WorkbookName', 'wb', 'wb', ' = "ThisBook"', 'String', ' = "ThisBook"', ''],
  Workbook: ['Workbook', 'wbk', 'wbk', ' = ActiveWorkbook', 'Workbook', ' = ActiveWorkbook', 'Set '],
  ColumnName: ['ColumnName', 'col', 'col', ' = "Header"', 'String', ' = "Header"', ''],
  ColumnNumber: ['ColumnNumber', 'col', 'col', ' = 1', 'Long', ' = 1', ''],
  RowNumber: ['RowNumber', 'row', 'row', ' = 1', 'Long', ' = 1', ''],
  MsgConfirmation: ['MsgConfirmation', 'msg', 'mg', ' = MsgBox("Are you sure?", vbYesNo + vbQuestion)\nSelect Case ' + placeHolder + '\n    Case vbYes\n        \'code\n    Case vbNo\n        \'code\nEnd Select', 'Integer', ' = 1', ''],
  Outlook: ['Outlook', 'oul', 'ou', ' = Nothing', 'Outlook', ' = Nothing', 'Set '],
};

// Prefix,declaration,constdeclaration
var scopeTypes = {
  Procedure: ['', 'Dim', ''],
  Module: ['m', 'Private', 'Private'],
  Global: ['g', 'Global', 'Global'],
};
var prefIndex = 1;
$(document).ready(function() {
  var i = 1;
  for (i = 1; i < elemNum + 1; i++) {
    $('#tabla')
      .find('tbody')
      .append(
        $('<tr>')
          .append($('<td class="nameColumn">Parameter' + i + ':</td>'))
          .append(
            $(
              '<td class="otherColumns"><input type="text" id="Name' +
                i +
                '" />'
            )
          )
          .append($('<td class="otherColumns"><select id="TypePar' + i + '">'))
          .append($('<td class="otherColumns">'))
      );
  }
  for (i = 1; i < elemNum + 1; i++) {
    $('#TypePar' + i)
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
      .append($('<option value="RowNumber">RowNumber</option>'))
      .append($('<option value="MsgConfirmation">MsgConfirmation</option>'))
      .append($('<option value="Outlook">Outlook</option>'));
  }

  $('#TypeFunc')
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
    .append($('<option value="RowNumber">RowNumber</option>'))
    .append($('<option value="MsgConfirmation">MsgConfirmation</option>'))
    .append($('<option value="Outlook">Outlook</option>'));

  $('#generate').click(function() {
    createFunction();
  });
  $('#clear').click(function() {
    resetNames();
  });
});

function capitalizeFirstLetter(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

function resetNames() {
  for (i = 1; i < elemNum + 1; i++) {
    $('#Name' + i).val('');
  }
}

function createFunction() {
  var i = 1;
  var inputPars = '';
  var funcText = '';
  var dimPars = '';
  var testPars = '';
  if ($('#ShortPrefix').is(':checked')) {
    prefIndex = 2;
  } else {
    prefIndex = 1;
  }
  for (i = 1; i < elemNum + 1; i++) {
    if ($('#Name' + i).val() !== '') {
      inputPars =
        inputPars +
        capitalizeFirstLetter($('#Name' + i).val()) +
        ' As ' +
        dictTypes[$('#TypePar' + i).val()][4] +
        ', ';
    }
  }
  funcText = '';
  if (inputPars !== '') {
    inputPars = inputPars.slice(0, inputPars.length - 2);
  }
  if ($('#classFunction').prop('checked')) {
    funcText = 'Function ';
  } else {
    funcText = 'Private Function ';
  }
  funcText =
    funcText +
    capitalizeFirstLetter($('#NameFunc').val()) +
    '(' +
    inputPars +
    ') As ' +
    dictTypes[$('#TypeFunc').val()][4];
  if ($('#arrayFunction').prop('checked')) {
    funcText += '()';
  }
  funcText += '\n';
  funcText = funcText + '    \'' + $('#remarkText').val() + '\n';
  funcText += '    \'Parameters:\n';

  for (i = 1; i < elemNum + 1; i++) {
    if ($('#Name' + i).val() !== '') {
      funcText =
        funcText +
        '    \'           ' +
        '{' +
        dictTypes[$('#TypePar' + i).val()][4];
      if (
        dictTypes[$('#TypePar' + i).val()][4] !==
        dictTypes[$('#TypePar' + i).val()][0]
      ) {
        funcText = funcText + ', ' + dictTypes[$('#TypePar' + i).val()][0];
      }
      funcText =
        funcText + '} ' + capitalizeFirstLetter($('#Name' + i).val()) + '\n';
    }
  }
  funcText =
    funcText +
    '    \'Returns:' +
    '{' +
    dictTypes[$('#TypeFunc').val()][4] +
    '}\n';
  funcText += '    \'Created by: Laszlo Tamas\n\n';

  funcText =
    funcText + '    Dim ' + dictTypes[$('#TypeFunc').val()][prefIndex] + 'Res';
  if ($('#arrayFunction').prop('checked')) {
    funcText += '()';
  }
  funcText = funcText + ' As ' + dictTypes[$('#TypeFunc').val()][4] + '\n';
  funcText += '\n    On Error GoTo FUNC_ERR\n\n';
  funcText =
    funcText + '    ' + dictTypes[$('#TypeFunc').val()][prefIndex] + 'Res';
  if ($('#arrayFunction').prop('checked')) {
    funcText =
      funcText +
      ' = Split(' +
      capitalizeFirstLetter(document.getElementById('Name1').value) +
      '," ")\n';
  } else {
    funcText = funcText + dictTypes[$('#TypeFunc').val()][5] + '\n';
  }
  funcText += '    \'Code here\n\n';
  funcText =
    funcText +
    '    ' +
    capitalizeFirstLetter($('#NameFunc').val()) +
    ' = ' +
    dictTypes[$('#TypeFunc').val()][prefIndex] +
    'Res' +
    '\n';
  funcText += '    \'---------------\n';
  funcText += 'FUNC_EXIT:\n';
  funcText += '    On Error GoTo 0\n';
  funcText += '    Exit Function\n';
  funcText += 'FUNC_ERR:\n';
  funcText =
    funcText +
    '    Debug.Print  "Error in Function ' +
    capitalizeFirstLetter($('#NameFunc').val()) +
    '"\n';
  funcText += '    If Err.Number Then\n';
  funcText += '        Debug.Print  Err.Description\n';
  funcText += '    End If\n';
  funcText += '    Resume FUNC_EXIT\n';

  funcText += 'End Function\n';
  $('#Code').val(funcText);

  // Test code
  dimPars = '';
  for (i = 1; i < elemNum + 1; i++) {
    if ($('#Name' + i).val() !== '') {
      dimPars =
        dimPars +
        '    Dim ' +
        dictTypes[$('#TypePar' + i).val()][prefIndex] +
        capitalizeFirstLetter($('#Name' + i).val()) +
        ' As ' +
        dictTypes[$('#TypePar' + i).val()][4] +
        '\n';
    }
  }

  dimPars += '\n';
  for (i = 1; i < elemNum + 1; i++) {
    if ($('#Name' + i).val() !== '') {
      dimPars =
        dimPars +
        '    ' +
        dictTypes[$('#TypePar' + i).val()][prefIndex] +
        capitalizeFirstLetter($('#Name' + i).val()) +
        dictTypes[$('#TypePar' + i).val()][5] +
        '\n';
    }
  }

  testPars = '';
  for (i = 1; i < elemNum + 1; i++) {
    if ($('#Name' + i).val() !== '') {
      testPars =
        testPars +
        dictTypes[$('#TypePar' + i).val()][prefIndex] +
        capitalizeFirstLetter($('#Name' + i).val()) +
        ', ';
    }
  }

  if (testPars !== '') {
    testPars = testPars.slice(0, testPars.length - 2);
  }

  funcText = 'Private Sub ';
  funcText = funcText + capitalizeFirstLetter($('#NameFunc').val()) + 'Test\n';
  if ($('#classFunction').prop('checked')) {
    funcText =
      funcText +
      '    \'Test procedure for ' +
      capitalizeFirstLetter($('#className').val()) +
      '.' +
      capitalizeFirstLetter($('#NameFunc').val()) +
      '\n';
  } else {
    funcText =
      funcText +
      '    \'Test procedure for ' +
      capitalizeFirstLetter($('#NameFunc').val()) +
      '\n';
  }
  funcText = funcText + '    \'' + $('#remarkText').val() + '\n';
  if ($('#classFunction').prop('checked')) {
    funcText =
      funcText +
      '    Dim cl' +
      capitalizeFirstLetter($('#className').val()) +
      ' As New ' +
      capitalizeFirstLetter($('#className').val()) +
      '\n';
  }
  funcText += '    Dim dtmStartTime As Date\n';
  funcText = funcText + dimPars + '\n\n';
  if ($('#arrayFunction').prop('checked')) {
    funcText =
      funcText +
      '    Dim Arr() As ' +
      dictTypes[$('#TypeFunc').val()][4] +
      '\n';
  }
  funcText += '    dtmStartTime = Now()\n';

  if ($('#arrayFunction').prop('checked')) {
    funcText += '    Dim iTer As Long\n';
    if ($('#classFunction').prop('checked')) {
      funcText =
        funcText +
        '    Arr = cl' +
        capitalizeFirstLetter($('#className').val()) +
        '.' +
        capitalizeFirstLetter($('#NameFunc').val()) +
        '(' +
        testPars +
        ')\n';
    } else {
      funcText =
        funcText +
        '    Arr = ' +
        capitalizeFirstLetter($('#NameFunc').val()) +
        '(' +
        testPars +
        ')\n';
    }
    funcText += '    For iTer = LBound(Arr) to UBound(Arr)\n';
    funcText =
      funcText +
      '        Debug.Print  "Function ' +
      capitalizeFirstLetter($('#NameFunc').val()) +
      ' test " & iTer & " >> " & Arr(iTer)\n';
    funcText += '    Next iTer\n';
  } else if ($('#classFunction').prop('checked')) {
    funcText =
      funcText +
      '    Debug.Print  "Function ' +
      capitalizeFirstLetter($('#NameFunc').val()) +
      ' test: >> " & cl' +
      capitalizeFirstLetter($('#className').val()) +
      '.' +
      capitalizeFirstLetter($('#NameFunc').val()) +
      '(' +
      testPars +
      ')\n';
    funcText =
      funcText +
      '    Set cl' +
      capitalizeFirstLetter($('#className').val()) +
      ' = Nothing\n';
  } else {
    funcText =
      funcText +
      '    Debug.Print  "Function ' +
      capitalizeFirstLetter($('#NameFunc').val()) +
      ' test: >> " & ' +
      capitalizeFirstLetter($('#NameFunc').val()) +
      '(' +
      testPars +
      ')\n';
  }

  funcText += 'End Sub\n';

  $('#CodeTest').val(funcText);
}
