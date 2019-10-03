// / <reference path="../../typings/globals/jquery/index.d.ts" />
$(document).ready(function() {
  $('#format').click(function() {
    formatVBA();
  });
  $('#clear').click(function() {
    clearCode();
  });
});
var lineNum = 0;
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
  lineNum = 0;
  for (i = 0; i < lines.length; i++) {
    line = lines[i].trim();
    // Remove extra spaces, add needed spaces
    lineNum++;
    line = addSpaceToOperators(line);
    line = removeSpaces(line);

    outText += line + '\n';
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
  var myRegexp;
  var matches;
  var match;
  var result = [];
  var resultIndex = [];
  var res;
  var strLine;
  var pos = 0;
  newString = lineText;
  myRegexp = /(=>|=<|>=|<=|<>|=|<|>|&|\+|-|\/)(?=[^\s])/g;

  strLine = newString;
  while ((matches = myRegexp.exec(newString))) {
    result.push(matches[1]);
    resultIndex.push(matches.index);
    // if (matches.index > pos) {
    //   strLine = strLine.replace(matches[1], matches[1] + ' ');
    //   pos = matches.index;
    // }
    // console.log(matches.index, matches[1]);

    /*
     * Result.push(matches[1]);
     * console.log(lineNum, matches[1], matches.index);
     */
  }

  // newString = strLine;
  var indOfArr = -1;
  for (res of result) {
    indOfArr++;
    console.log(res.length, indOfArr, resultIndex[indOfArr], res);
    // newString = newString.replace(
    //   newString.substring(
    //     resultIndex[indOfArr],
    //     resultIndex[indOfArr] + res.length
    //   ),
    //   res + ' '
    // );
    str1 = newString.substring(0, resultIndex[indOfArr] - 1);
    str2 = newString.substring(
      resultIndex[indOfArr],
      resultIndex[indOfArr] + res.length - 1
    );
    str2 = res + ' ';
    str3 = newString.substring(resultIndex[indOfArr] + res.length + 1);
    console.log(str1 + '*' + str2 + '*' + str3);
    newString = str1 + str2 + str3;
    // newString = newString.replace(res, ' ' + res + ' ');
    // newString = replaceAll(newString, res, ' ' + res + ' ');
    // console.log(newString);
  }

  /*
   * Result = [];
   * myRegexp = /(?=[^\s])(=>|=<|>=|<=|<>|=|<|>|&|\+|-|\/)/g;
   */

  /*
   * While ((matches = myRegexp.exec(newString))) {
   *   result.push(matches[1]);
   * }
   * for (res of result) {
   *   newString = newString.replace(res, ' ' + res);
   *   // console.log(newString);
   * }
   */
  return newString;
}

function searchString(string, pattern) {
  var result = [];
  var i;
  var matches = string.match(new RegExp(pattern.source, pattern.flags));

  for (i = 0; i < matches.length; i++) {
    result.push(new RegExp(pattern.source, pattern.flags).exec(matches[i]));
  }

  return result;
}

function getMatches(string, regex, index) {
  var matches = [];
  var match;
  while ((match = regex.exec(string))) {
    matches.push(match[index]);
  }
  return matches;
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
function replaceAll(str, find, replace) {
  return str.replace(new RegExp(find, 'g'), replace);
}
