/* eslint-disable capitalized-comments */
// / <reference path="../../typings/globals/jquery/index.d.ts" />

// prettier-ignore
var commandsUp = new Array('AppActivate', 'Beep', 'Call', 'ChDir', 'ChDrive',
  'Close', 'Const', 'Date', 'Declare', 'DeleteSetting', 'Dim', 'Do', 'Do While',
  'Loop', 'End', 'Erase', 'Error', 'Exit Do', 'Exit For', 'Exit Function',
  'Exit Property', 'Exit Sub', 'FileCopy', 'For', 'Each', 'Next', 'For Each',
  'Function', 'Get', 'GoSub', 'Return', 'GoTo', 'If', 'Then', 'Else', 'ElseIf',
  'Input #', 'Kill', 'Let', 'Line Input #', 'Load', 'Lock', 'Unlock', 'Mid',
  'MkDir', 'Name', 'On Error', 'On', 'Open', 'Option Base', 'Option Compare',
  'Option Explicit', 'Option Private', 'Print #', 'Private', 'Property Get',
  'Property Let', 'Property Set', 'Public', 'Put', 'RaiseEvent', 'Randomize',
  'ReDim', 'REM', 'Reset', 'Resume', 'RmDir', 'SaveSetting', 'Seek',
  'Select Case', 'SendKeys', 'Set', 'SetAttr', 'Static', 'Stop', 'Sub',
  'Time', 'Type', 'Unload', 'While', 'Wend', 'Width #', 'With', 'Write #',
  'End Sub', 'End Function', 'Debug.Print', 'MsgBox', 'Wait', 'Private Sub',
  '#If', '#Else', '#End If', 'Case', 'End Select', 'End Property', 'Mac',
  'Win32', 'Win64', 'Vba6', 'Vba7', 'Append', 'Output');

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
  'getCount', 'MacScript');
// prettier-ignore
var typesUp = new Array(' As String', ' As Integer', ' As Double',
  ' As WorkSheet', ' As WorkBook', ' As Long', ' As Variant', ' As Boolean',
  ' As Object', ' As Date');
// prettier-ignore
var objectsUp = new Array('ThisComponent', 'CurrentController', 'ActiveSheet',
  'ActiveWorkbook', 'GlobalScope', 'BasicLibraries', 'StarDesktop',
  'RunAutoMacros');
// prettier-ignore
var activityUp = new Array('Activate', 'ActiveSheet', 'getCurrentSelection',
  'ScreenUpdating', 'LockControllers', 'Open', 'Name', 'Value', 'String',
  'Address', 'Select');
// prettier-ignore
var subobjectsUp = new Array('Cells', 'Sheets', 'Range');

var currentIndent = 0;
var firstCase = false;

$(document).ready(function() {
  $('#format').click(function() {
    formatVBA();
  });
  $('#clear').click(function() {
    clearCode();
  });
});
function formatVBA() {
  var lines = $('#Code')
    .val()
    .split('\n');
  var outText = '';
  var line = '';
  var i = 0;
  for (i = 0; i < lines.length; i++) {
    line = lines[i].trim();
    line = removeSpaces(line);
    line = formatConstDeclarationLine(line);
    line = formatSubLine(line);
    line = formatFuncLine(line);
    line = formatVBACommand(line);
    line = formatVBAFunction(line);
    line = formatVBAType(line);
    line = formatVBAObject(line);
    line = formatVBAActivity(line);
    line = formatVBASubobject(line);
    line = getIndentedLine(line);
    outText += line + '\n';
  }
  $('#CodeFormat').val(outText);
}

function getIndentedLine(line) {
  var ret = getIndent(currentIndent) + line.trim();
  var regex;
  var match;
  regex = /^\s*(private|public|global|option explicit|end sub|end function|end property)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = line.trim();
    // console.log('MainDeclare Indent:' + currentIndent + ' ' + line);
  }
  // Main line
  regex = /^\s*(private sub|private function|public sub|public function|global sub|global function|sub|function|private property|public property|global property)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    currentIndent = 1;
    ret = line.trim();
    // console.log('Main Indent:' + currentIndent + ' ' + line);
  }
  regex = /^\s*(for|while|with)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = getIndent(currentIndent) + line.trim();
    currentIndent++;
    // console.log('ForWhile Indent:' + currentIndent + ' ' + line);
  }
  regex = /^\s*(select case)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = getIndent(currentIndent) + line.trim();
    currentIndent++;
    firstCase = true;
    // console.log('Select Case Indent:' + currentIndent + ' ' + firstCase + ' ' + line);
  }
  regex = /^\s*(case)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    currentIndent--;
    // console.log('Case Indent firstCase: ' + firstCase + ' ' + line);
    if (firstCase) {
      currentIndent++;
      firstCase = false;
    }
    // console.log('Case Indent:' + currentIndent + ' ' + line);
    ret = getIndent(currentIndent) + line.trim();
    currentIndent++;
  }
  regex = /^\s*(end select)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    currentIndent -= 2;
    ret = getIndent(currentIndent) + line.trim();
    // console.log('End Select Indent:' + currentIndent + ' ' + line);
  }
  regex = /\s*(then| _|#then)\b\s*$/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = getIndent(currentIndent) + line.trim();
    currentIndent++;
    // console.log('Then Indent:' + currentIndent + ' ' + line);
  }

  regex = /^\s*(next|end if|#end if|wend|end with)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    currentIndent--;
    ret = getIndent(currentIndent) + line.trim();
    // console.log('NextEndif Indent:' + currentIndent + ' ' + line);
  }
  regex = /^\s*(else|elseif|#else|#elseif)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    currentIndent--;
    ret = getIndent(currentIndent) + line.trim();
    currentIndent++;
    // console.log('Else Indent:' + currentIndent + ' ' + line);
  }
  return ret;
}
function formatVBACommand(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < commandsUp.length; k++) {
    regex = new RegExp('\\b' + commandsUp[k] + '\\b', 'gi');
    ret = ret.replace(regex, commandsUp[k]);
  }
  return ret;
}
function formatVBAFunction(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < funcsUp.length; k++) {
    regex = new RegExp('\\b' + funcsUp[k] + '\\s*\\(', 'gi');
    ret = ret.replace(regex, funcsUp[k] + '(');
  }
  return ret;
}
function formatVBAType(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < typesUp.length; k++) {
    regex = new RegExp('\\b' + typesUp[k] + '\\b', 'gi');
    ret = ret.replace(regex, typesUp[k]);
  }
  return ret;
}
function formatVBAObject(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < objectsUp.length; k++) {
    regex = new RegExp('\\b' + objectsUp[k] + '\\s*\\.', 'gi');
    ret = ret.replace(regex, objectsUp[k] + '.');
  }
  return ret;
}
function formatVBAActivity(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < activityUp.length; k++) {
    regex = new RegExp('\\.\\s*' + activityUp[k], 'gi');
    ret = ret.replace(regex, '.' + activityUp[k]);
  }
  return ret;
}
function formatVBASubobject(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < subobjectsUp.length; k++) {
    regex = new RegExp('\\.\\s*' + subobjectsUp[k] + '\\s*\\(', 'gi');
    ret = ret.replace(regex, '.' + subobjectsUp[k] + '(');
  }
  return ret;
}
function formatFuncLine(line) {
  var ret = line;
  var subret = '';
  var mainRegexp;
  var subRegexp;
  var mainMatch;
  var subMatch;
  var subElems;
  var index;
  var elem;
  mainRegexp = /(private\s*function |global\s*function |function )\s*(.*)\((.*)\)\s*as\s*(\b[a-zA-Z0-9_]+\b)$/gi;
  mainMatch = mainRegexp.exec(line);
  try {
    subElems = mainMatch[3].split(',');
    for (index = 0; index < subElems.length; index++) {
      subRegexp = /(\b[a-zA-Z0-9_]+\b)\s*as\s*(\b[a-zA-Z0-9_]+\b)/gi;
      elem = subElems[index].trim();
      // console.log('E:' + elem);
      subMatch = subRegexp.exec(elem);
      // console.log(elem + ' SM:' + subMatch.length);
      try {
        subret += subMatch[1] + ' As ' + capitalize(subMatch[2]);
        if (index + 1 !== subElems.length) {
          subret += ', ';
        }
      } catch (error) {
        subret += formatOptionalPart(elem);
      }
    }
    ret =
      capitalize(mainMatch[1]) +
      mainMatch[2].trim() +
      '(' +
      subret +
      ') As ' +
      capitalize(mainMatch[4]);
    ret = ret.replace('function', 'Function');
  } catch (error) {}
  return ret;
}

function formatOptionalPart(line) {
  var ret = line;
  var mainRegexp;
  var mainMatch;
  // console.log('Try OPTIONAL ' + line);
  mainRegexp = /(\boptional\b)\s*(.*)\s*=\s*(.*)\s*/gi;
  mainMatch = mainRegexp.exec(line);
  try {
    // console.log('MOPT:' + mainMatch.length);
    ret = 'Optional ' + mainMatch[2] + ' = ' + mainMatch[3];
  } catch (error) {
    // console.log('Not OPTIONAL');
  }
  return ret;
}
function formatSubLine(line) {
  var ret = line;
  var subret = '';
  var mainRegexp;
  var subRegexp;
  var mainMatch;
  var subMatch;
  var subElems;
  var index;
  var elem;
  mainRegexp = /(private\s*sub |global\s*sub |sub )\s*(.*)\((.*)\)\s*$/gi;
  mainMatch = mainRegexp.exec(line);
  try {
    subElems = mainMatch[3].split(',');
    for (index = 0; index < subElems.length; index++) {
      subRegexp = /(\b[a-zA-Z0-9_]+\b)\s*as\s*(\b[a-zA-Z0-9_]+\b)/gi;
      elem = subElems[index].trim();
      // console.log('E:' + elem);
      subMatch = subRegexp.exec(elem);
      // console.log(elem + ' SM:' + subMatch.length);
      try {
        subret += subMatch[1] + ' As ' + capitalize(subMatch[2]);
        if (index + 1 !== subElems.length) {
          subret += ', ';
        }
      } catch (error) {
        subret += formatOptionalPart(elem);
      }
    }
    ret = capitalize(mainMatch[1]) + mainMatch[2].trim() + '(' + subret + ')';
    ret = ret.replace('sub', 'Sub');
  } catch (error) {}
  return ret;
}

function formatConstDeclarationLine(line) {
  var ret = line;
  myRegexp = /(private |public )\s*const\s*(.*)\s*as\s*(\b[a-zA-Z0-9_]+\b)\s*=\s*(.*$)/gi;
  match = myRegexp.exec(line);
  try {
    // console.log(capitalize(match[1]));
    ret =
      capitalize(match[1]) +
      'Const ' +
      match[2].trim() +
      ' As ' +
      capitalize(match[3]) +
      ' = ' +
      match[4];
  } catch (error) {
    // console.log(error);
  }
  return ret;
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

function clearCode() {
  $('#Code').val('');
  $('#CodeFormat').val('');
}

function capitalize(string) {
  var ret;
  if (string.charAt(0) === '#') {
    ret = '#' + string.charAt(1).toUpperCase() + string.slice(2).toLowerCase();
  } else {
    ret = string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
  }
  return ret;
}

function getIndent(num) {
  var res = '';
  var indent = '';
  if (currentIndent < 0) {
    currentIndent = 0;
  }
  if (num > 0) {
    for (index = 0; index < $('#indentSize').val(); index++) {
      indent += ' ';
    }
    for (i = 0; i < num; i++) {
      res += indent;
    }
  }
  return res;
}
