// / <reference path="../../typings/globals/jquery/index.d.ts" />
$(document).ready(function() {
  $('#format').click(function() {
    formatVBA();
  });
  $('#clear').click(function() {
    clearCode();
  });
});

function formatVBA() {
  // prettier-ignore
  var commandsUp = new Array('AppActivate', 'Beep', 'Call', 'ChDir', 'ChDrive',
    'Close', 'Const', 'Date', 'Declare', 'DeleteSetting', 'Dim', 'Do', 'Do While',
    'Loop', 'End', 'Erase', 'Error', 'Exit Do', 'Exit For', 'Exit Function',
    'Exit Property', 'Exit Sub', 'FileCopy', 'For', 'Each', 'Next', 'For Each',
    'Function', 'Get', 'GoSub', 'Return', 'GoTo', 'If', 'Then', 'Else',
    'Input #', 'Kill', 'Let', 'Line Input #', 'Load', 'Lock', 'Unlock', 'Mid',
    'MkDir', 'Name', 'On Error', 'On', 'Open', 'Option Base', 'Option Compare',
    'Option Explicit', 'Option Private', 'Print #', 'Private', 'Property Get',
    'Property Let', 'Property Set', 'Public', 'Put', 'RaiseEvent', 'Randomize',
    'ReDim', 'REM', 'Reset', 'Resume', 'RmDir', 'SaveSetting', 'Seek',
    'Select Case', 'SendKeys', 'Set', 'SetAttr', 'Static', 'Stop', 'Sub',
    'Time', 'Type', 'Unload', 'While', 'Wend', 'Width #', 'With', 'Write #',
    'End Sub', 'End Function', 'Debug.Print', 'MsgBox', 'Wait', 'Private Sub',
    '#If', '#Else', '#End If');
  // prettier-ignore
  var funcsUp = new Array('Abs', 'Array', 'Asc', 'Atn', 'CBool', 'CByte',
    'CCur', 'CDate', 'CDbl', 'CDec', 'Choose', 'Chr', 'CInt', 'CLng', 'Cos',
    'CurDir', 'CVar', 'CVErr', 'CSng', 'CStr', 'Date', 'DateAdd', 'DateDiff',
    'DatePart', 'DateSerial', 'DateValue', 'Day', 'DDB', 'Dir', 'Error', 'Exp',
    'FileAttr', 'FileDateTime', 'FileLen', 'Filter', 'Fix', 'Format',
    'FormatCurrency', 'FormatDateTime', 'FormatNumber', 'FormatPercent',
    'FV', 'GetAttr', 'Hex', 'Hour', 'IIf', 'InputBox', 'InStr', 'InStrRev',
    'Int', 'IPmt', 'IRR', 'IsArray', 'IsDate', 'IsEmpty', 'IsError',
    'IsMissing', 'IsNull', 'IsNumeric', 'IsObject', 'Join', 'LBound',
    'LCase', 'Left', 'Len', 'Log', 'LTrim', 'Mid', 'Minute', 'MIRR',
    'Month', 'MonthName', 'MsgBox', 'Now', 'NPer', 'NPV', 'Oct', 'Pmt',
    'PPmt', 'PV', 'Rate', 'Replace', 'Right', 'Rnd', 'Round', 'RTrim',
    'Second', 'Sgn', 'Sin', 'SLN', 'Space', 'Split', 'Sqr', 'Str',
    'StrComp', 'StrConv', 'String', 'StrReverse', 'Switch', 'SYD',
    'Tan', 'Time', 'Timer', 'TimeSerial', 'TimeValue', 'Trim',
    'UBound', 'UCase', 'Val', 'Weekday', 'WeekdayName', 'Year',
    'addItem', 'getCellRangeByName', 'getCellByPosition', 'getByName',
    'setActiveSheet', 'Worksheets', 'Sheets', 'findSheetIndex',
    'InsertNewByName', 'LoadLibrary', 'getURL', 'DirectoryNameoutofPath',
    'callFunction', 'hasLocation', 'Wait', 'FileNameOutOfPath',
    'GetDocumentType', 'HasUnoInterfaces', 'getComponents', 'createEnumeration',
    'hasMoreElements', 'nextElement', 'loadComponentFromURL', 'Open',
    'getCount');
  // prettier-ignore
  var typesUp = new Array(' As String', ' As Integer', ' As Double',
    ' As WorkSheet', ' As WorkBook', ' As Long', ' As Variant', ' As Boolean',
    ' As Object', ' As Date', ' Then');
  // prettier-ignore
  var objectsUp = new Array('ThisComponent', 'CurrentController', 'ActiveSheet',
    'ActiveWorkbook', 'GlobalScope', 'BasicLibraries', 'StarDesktop',
    'RunAutoMacros');
  // prettier-ignore
  var activityUp = new Array('Activate', 'ActiveSheet', 'getCurrentSelection',
    'ScreenUpdating', 'LockControllers', 'Open', 'Name', 'Value', 'String',
    'Address', 'Select');
  // prettier-ignore
  var commands = new Array('if', 'else', 'else if', 'end if', 'while', 'wend',
    'sub', 'private sub', 'public sub', 'function', 'private function',
    'public function', 'end sub', 'end function', 'enum', 'private enum',
    'public enum', 'property', 'private property', 'public property',
    'end enum', 'end property', 'for', 'next', 'with', 'end with', 'do', 'loop',
    'select case', 'case', 'end select', '#if', '#else', '#end if');
  // prettier-ignore
  var commandsBefore = new Array('', '-', '-', '-', '', '-', '0', '0', '0', '0',
    '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '', '-', '',
    '-', '', '-', '', '-', '--', '', '-', '-');
  // prettier-ignore
  var commandsAfter = new Array('+', '+', '+', '', '+', '', '+', '+', '+', '+',
    '+', '+', '0', '0', '+', '+', '+', '+', '+', '+', '0', '0', '+', '', '+',
    '', '+', '', '++', '+', '', '+', '+', '');
  // Sorokra bontás
  var lines = $('#Code')
    .val()
    .split('\n');
  var outText = '';
  var line = '';
  var beforeIndent = 0;
  var afterIndent = 0;
  var i = 0;
  var k = 0;
  for (i = 0; i < lines.length; i++) {
    line = lines[i].trim();
    // Remove extra spaces, add needed spaces
    line = addSpaceToOperators(line);
    line = removeSpaces(line);

    for (k = 0; k < commands.length; k++) {
      if (
        (commands[k] + ' ').toLowerCase() ===
        line
          .slice(0, commands[k].length + 1)
          .toLowerCase()
          .trim() +
          ' '
      ) {
        // OutText = outText + getIndent(beforeIndent) + line + "QQQ\n";
        if (commandsAfter[k] === '+') {
          // BeforeIndent++;
          afterIndent = 1;
        }
        if (commandsAfter[k] === '++') {
          // BeforeIndent++;
          afterIndent = 2;
        }
        if (commandsAfter[k] === '0') {
          beforeIndent = 0;
        }
        if (commandsBefore[k] === '-') {
          beforeIndent--;
        }
        if (commandsBefore[k] === '--') {
          beforeIndent--;
          beforeIndent--;
        }
        if (commandsBefore[k] === '0') {
          beforeIndent = 0;
        }
      }
    }
    // Make uppercase statement
    for (k = 0; k < commandsUp.length; k++) {
      if (
        (commandsUp[k] + ' ').toLowerCase() ===
        line
          .slice(0, commandsUp[k].length + 1)
          .toLowerCase()
          .trim() +
          ' '
      ) {
        line = line.replace(
          new RegExp(line.slice(0, commandsUp[k].length), 'i'),
          commandsUp[k]
        );
      }
    }
    // Make uppercase functions
    for (k = 0; k < funcsUp.length; k++) {
      line = line.replace(
        new RegExp(funcsUp[k] + '\\(', 'gi'),
        funcsUp[k] + '('
      );
    }
    // Make uppercase types
    for (k = 0; k < typesUp.length; k++) {
      line = line.replace(new RegExp(typesUp[k], 'gi'), typesUp[k]);
    }
    // Make uppercase objects
    for (k = 0; k < objectsUp.length; k++) {
      line = line.replace(
        new RegExp(objectsUp[k] + '\\.', 'gi'),
        objectsUp[k] + '.'
      );
    }
    // Make uppercase activity
    for (k = 0; k < activityUp.length; k++) {
      line = line.replace(
        new RegExp('\\.' + activityUp[k], 'gi'),
        '.' + activityUp[k]
      );
    }
    outText = outText + getIndent(beforeIndent) + line + '\n';
    if (afterIndent === 1) {
      afterIndent = 0;
      beforeIndent++;
    }
    if (afterIndent === 2) {
      // AfterIndent = 1;
      beforeIndent++;
      beforeIndent++;
    }
  }
  $('#CodeFormat').val(outText);
}

/**
 * Remove extra spaces except within quotation marks
 * @param  {string} lineText
 */
function removeSpaces(lineText) {
  // eslint-disable-next-line id-length
  var newString = lineText.replace(/([^"]+)|("[^"]+")/g, function($0, $1, $2) {
    if ($1) {
      return $1.replace(/\s{2,}/g, ' ');
    }
    return $2;
  });
  return newString;
}

function addSpaceToOperators(lineText) {
  var newString = '';
  // =
  newString = lineText.replace(/=(?=[^\s])/g, '= ');
  newString = newString.replace(/(?=[^\s])=/g, ' =');
  // <
  newString = lineText.replace(/<(?=[^\s])/g, '< ');
  newString = newString.replace(/(?=[^\s])</g, ' <');
  // >
  newString = lineText.replace(/>(?=[^\s])/g, '> ');
  newString = newString.replace(/(?=[^\s])>/g, ' >');
  // &
  // eslint-disable-next-line no-useless-escape
  newString = lineText.replace(/\&(?=[^\s])/g, '& ');
  // eslint-disable-next-line no-useless-escape
  newString = newString.replace(/(?=[^\s])\&/g, ' &');

  return newString;
}

function getIndent(num) {
  var res = '';
  var indent = '';
  for (index = 0; index < $('#indentSize').val(); index++) {
    indent += ' ';
  }
  for (i = 0; i < num; i++) {
    res += indent;
  }
  return res;
}

function clearCode() {
  $('#Code').val('');
  $('#CodeFormat').val('');
}
