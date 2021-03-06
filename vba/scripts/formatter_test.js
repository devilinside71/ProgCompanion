/* eslint-disable no-negated-condition */
/* eslint-disable capitalized-comments */
// / <reference path="../../typings/globals/jquery/index.d.ts" />

// #region Keywords

// prettier-ignore
var commandsUp = new Array('#Else', '#End If', '#If', 'AppActivate', 'Append',
  'Beep', 'ByRef', 'ByVal', 'Call', 'Case', 'ChDir', 'ChDrive', 'Close',
  'Const', 'Date', 'Debug.Print', 'Declare', 'DeleteSetting', 'Dim', 'Do',
  'Do While', 'Each', 'Else', 'ElseIf', 'End', 'End Function', 'End Property',
  'End Select', 'End Sub', 'Erase', 'Error', 'Exit Do', 'Exit For',
  'Exit Function', 'Exit Property', 'Exit Sub', 'FileCopy', 'For', 'For Each',
  'Function', 'Get', 'GoSub', 'GoTo', 'If', 'Iif', 'Input #', 'Kill', 'Let',
  'Line Input #', 'Load', 'Lock', 'Loop', 'Mid', 'MkDir', 'MsgBox',
  'Name', 'Next', 'On', 'On Error', 'Open', 'Option Base', 'Option Compare',
  'Option Explicit', 'Option Private', 'Output', 'Print #', 'Private',
  'Private Sub', 'Property Get', 'Property Let', 'Property Set', 'Public',
  'Put', 'REM', 'RaiseEvent', 'Randomize', 'ReDim', 'Reset', 'Resume', 'Return',
  'RmDir', 'SaveSetting', 'Seek', 'Select Case', 'SendKeys', 'Set', 'SetAttr',
  'Static', 'Stop', 'Sub', 'Then', 'Time', 'Type', 'Unload', 'Unlock', 'Vba6',
  'Vba7', 'Wait', 'Wend', 'While', 'Width #', 'With', 'Write #');

// prettier-ignore
var funcsUp = new Array('Abs', 'Array', 'Asc', 'Atn', 'CBool', 'CByte', 'CCur',
  'CDate', 'CDbl', 'CDec', 'CInt', 'CLng', 'CSng', 'CStr', 'CVErr', 'CVar',
  'Choose', 'Chr', 'Cos', 'CurDir', 'DDB', 'Date', 'DateAdd', 'DateDiff',
  'DatePart', 'DateSerial', 'DateValue', 'Day', 'Dir', 'DirectoryNameoutofPath',
  'Error', 'Exp', 'FV', 'FileAttr', 'FileDateTime', 'FileLen',
  'FileNameOutOfPath', 'Filter', 'Fix', 'Format', 'FormatCurrency',
  'FormatDateTime', 'FormatNumber', 'FormatPercent', 'GetAttr',
  'GetDocumentType', 'HasUnoInterfaces', 'Hex', 'Hour', 'IIf', 'IPmt', 'IRR',
  'InStr', 'InStrRev', 'InputBox', 'InsertNewByName', 'Int', 'IsArray',
  'IsDate', 'IsEmpty', 'IsError', 'IsMissing', 'IsNull', 'IsNumeric',
  'IsObject', 'Join', 'LBound', 'LCase', 'LTrim', 'Left', 'Len', 'LoadLibrary',
  'Log', 'MIRR', 'MacScript', 'Mid', 'Minute', 'Month', 'MonthName', 'MsgBox',
  'NPV', 'NPer', 'Now', 'Oct', 'Open', 'PPmt', 'PV', 'Pmt', 'RTrim', 'Rate',
  'Replace', 'Right', 'Rnd', 'Round', 'SLN', 'SYD', 'Second', 'Sgn', 'Sheets',
  'Sin', 'Space', 'Split', 'Sqr', 'Str', 'StrComp', 'StrConv', 'StrReverse',
  'String', 'Switch', 'Tan', 'Time', 'TimeSerial', 'TimeValue', 'Timer', 'Trim',
  'UBound', 'UCase', 'Val', 'Wait', 'Weekday', 'WeekdayName', 'Worksheets',
  'Year', 'addItem', 'callFunction', 'createEnumeration', 'findSheetIndex',
  'getByName', 'getCellByPosition', 'getCellRangeByName', 'getComponents',
  'getCount', 'getURL', 'hasLocation', 'hasMoreElements', 'loadComponentFromURL',
  'nextElement', 'setActiveSheet');
// prettier-ignore
var typesUp = new Array('As Boolean', 'As Date', 'As Double', 'As Integer',
  'As Long', 'As Object', 'As String', 'As Variant', 'As WorkBook',
  'As WorkSheet', 'As Byte', 'As Single', 'As Currency', 'As Decimal');
// prettier-ignore
var objectsUp = new Array('ActiveSheet', 'ActiveWorkbook', 'BasicLibraries',
  'CurrentController', 'GlobalScope', 'RunAutoMacros', 'StarDesktop', 'ThisComponent');
// prettier-ignore
var activityUp = new Array('Activate', 'ActiveSheet', 'Address', 'LockControllers',
  'Name', 'Open', 'ScreenUpdating', 'Select', 'String', 'Value', 'getCurrentSelection');
// prettier-ignore
var subobjectsUp = new Array('Cells', 'Range', 'Sheets');
// prettier-ignore
var constUp = new Array('vbTrue', 'vbFalse', 'vbCr', 'vbCrLf', 'vbFormFeed',
  'vbLf', 'vbNewLine', 'vbNullChar', 'vbNullString', 'vbTab', 'vbVerticalTab',
  'vbBinaryCompare', 'vbTextCompare', 'vbSunday', 'vbMonday', 'vbTuesday',
  'vbWednesday', 'vbThursday', 'vbFriday', 'vbSaturday', 'vbUseSystemDayOfWeek',
  'vbFirstJan1', 'vbFirstFourDays', 'vbFirstFullWeek', 'vbGeneralDate', 'vbLongDate',
  'vbShortDate', 'vbLongTime', 'vbShortTime', 'vbObjectError', 'vbEmpty', 'vbNull',
  'vbInteger', 'vbLong', 'vSingle', 'vbDouble', 'vbCurrency', 'vbDate', 'vbString',
  'vbObject', 'vbError', 'vbBoolean', 'vbVariant', 'vbDataObject', 'vbDecimal',
  'vbByte', 'vbArray', 'Mac', 'Win64', 'Win32', 'Vba6', 'Vba7');
// prettier-ignore
var msgConstUp = new Array('vbOKOnly', 'vbOKCancel', 'vbAbortRetryIgnore',
  'vbYesNoCancel', 'vbYesNo', 'vbRetryCancel', 'vbCritical', 'vbQuestion',
  'vbExclamation', 'vbInformation', 'vbDefaultButton1', 'vbDefaultButton2',
  'vbDefaultButton3', 'vbDefaultButton4', 'vbApplicationModal', 'vbSystemModal',
  'vbMsgBoxHelpButton', 'VbMsgBoxSetForeground', 'vbMsgBoxRight', 'vbMsgBoxRtlReading',
  'vbOK', 'vbCancel', 'vbAbort', 'vbRetry', 'vbIgnore', 'vbYes', 'vbNo');
// #endregion

var currentIndent = 0;
var firstCase = false;
var underscored = true;
var underscoreCount = 0;
var breakPoint = 80;

/**
 * Main function
 * @param  {} document
 * @param  {} .ready(function(
 */
$(document).ready(function() {
  $('#format').click(function() {
    formatVBA();
  });
  $('#clear').click(function() {
    clearCode();
  });
  $('#test').click(function() {
    test();
  });
});

function test() {
  var newLines = [];
  var lines = $('#Code')
    .val()
    .split('\n');
  var outText = '';
  var line = '';
  var i = 0;
  for (i = 0; i < lines.length; i++) {
    line = lines[i].trim();
    newLines.push(splitLine(line));
  }
  for (i = 0; i < newLines.length; i++) {
    line = newLines[i].trim();

    outText += line + '\n';
  }

  $('#CodeFormat').val(outText);
}

/**
 * Format VBA code
 */
function formatVBA() {
  var lines = $('#Code')
    .val()
    .split('\n');
  var outText = '';
  var line = '';
  var i = 0;

  for (i = 0; i < lines.length; i++) {
    line = lines[i].trim();
    line = addSpaceToOperators(line);
    line = removeSpaceAroundBrackets(line);
    line = formatSpecialLine(line);

    line = removeSpaces(line);
    line = formatConstDeclarationLine(line);
    // line = formatSubLine(line);
    // line = formatFuncLine(line);
    line = formatVBACommand(line);
    line = formatVBAFunction(line);
    line = formatVBAType(line);
    line = formatVBAObject(line);
    line = formatVBAActivity(line);
    line = formatVBASubobject(line);
    line = formatVBAConstant(line);
    line = formatVBAMsgConstant(line);
    line = getIndentedLine(line);
    outText += line + '\n';
  }
  outText = getSplitLines(outText);
  $('#CodeFormat').val(outText);
}

function getSplitLines(tempText) {
  var ret = '';
  var newLines = [];
  var tempLine = '';
  newLines = tempText.split('\n');
  breakPoint = $('#lineLength').val();
  for (i = 0; i < newLines.length; i++) {
    line = newLines[i];
    tempLine = splitLine(line) + '\n';
    ret += tempLine;
  }
  return ret;
}

/**
 * Determine if line is a remark one
 * @param  {} line
 */
function remLine(line) {
  var ret = false;
  regex = /^\s*(rem\b|')/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = true;
    console.log('REM LINE: ' + match[1] + ' ' + line);
  }
  return ret;
}

function brokenLine(line) {
  var ret = false;
  regex = /( _)$/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = true;
    console.log('BrokenLine LINE: ' + match[1] + ' ' + line);
  }
  return ret;
}
/**
 * Split line to specific length
 * @param  {} line
 */
function splitLine(line) {
  var retVal;
  var lineLengthCounter = 0;
  var regex;
  var parts;
  var partsIndex;
  var lineIndent;
  var tempPart = '';
  var retVal2 = '';
  var retVal3 = '';
  // Determine initial indent
  lineIndent = '';
  regex = /^(\s*).*$/;
  match = regex.exec(line);
  if (match !== null) {
    lineIndent = match[1];
  }
  // console.log('Initial indent: ' + lineIndent.length + ' ' + line);
  if (line.length < breakPoint || remLine(line) || brokenLine(line)) {
    // console.log('Line is smaller than ' + breakPoint + ': ' + line);
    retVal = line;
  } else {
    // Operators except between quotation
    // prettier-ignore
    regex = new RegExp('(>|<|=|\\+|-|&|\\/|,)(?=(?:[^"]*"[^"]*")*[^"]*$)', 'gi');
    // Operators except between quotation an brackets
    // prettier-ignore
    // regex = new RegExp('(>|<|=|\\+|-|&|\\/|,)(?=(?=(?:[^"]*"[^"]*")*[^"]*$)(?![^\\(]*\\)))', 'gi');
    // eslint-disable-next-line id-length
    retVal = line.replace(regex, function($0, $1) {
      // console.log($1+' '+line);
      return $1+' _';
    });
    parts = retVal.split(' _');
    retVal2 = '';
    retVal3 = '_\n' + lineIndent + '   ';
    // console.log('Parts: ' + parts.length + ' ' + line);
    lineLengthCounter = 0;
    for (partsIndex = 0; partsIndex < parts.length; partsIndex++) {
      tempPart = parts[partsIndex].replace(/ _$/gi, '').trim() + ' ';
      lineLengthCounter += tempPart.length;
      if (lineLengthCounter < breakPoint || partsIndex === 0) {
        if (partsIndex === 0) {
          retVal2 += lineIndent + tempPart;
        } else {
          retVal2 += tempPart;
        }
      } else {
        retVal3 += tempPart;
      }
    }
  }
  if (retVal2 !== '') {
    retVal = retVal2 + retVal3;
  }
  return retVal;
}

/**
 * Create indented line
 * @param  {} line
 */
function getIndentedLine(line) {
  var ret = getIndent(currentIndent) + line.trim();
  var regex;
  var match;

  // Lines with no indent
  regex = /^\s*(private|public|global|option explicit|end sub|end function|end property|'#region|'#endregion)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = line.trim();
    // console.log('MainDeclare Indent:' + currentIndent + ' ' + +match[1] + ' ' + line);
  }

  // Lines with no indent but indent after
  regex = /^\s*(private sub|private function|public sub|public function|global sub|global function|sub|function|private property|public property|global property)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    currentIndent = 1;
    ret = line.trim();
    // console.log('Main Indent:' + currentIndent + ' ' + line);
  }

  // Loop start, indent after
  regex = /^\s*(for|while|with|do while)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = getIndent(currentIndent) + line.trim();
    currentIndent++;
    // console.log('ForWhile Indent:' + currentIndent + ' ' + line);
  }

  // Select Case, indent after
  regex = /^\s*(select case)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = getIndent(currentIndent) + line.trim();
    currentIndent++;
    firstCase = true;
    // console.log('Select Case Indent:' + currentIndent + ' ' + firstCase + ' ' + line);
  }

  // Case, indent after, outdent after last Case
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

  // End Select, outdent after
  regex = /^\s*(end select)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    currentIndent -= 2;
    ret = getIndent(currentIndent) + line.trim();
    // console.log('End Select Indent:' + currentIndent + ' ' + line);
  }

  // Then at end of line, indent after
  regex = /\s*(then|#then)\b\s*$/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = getIndent(currentIndent) + line.trim();
    currentIndent++;
    // console.log('Then Indent:' + currentIndent + ' ' + line);
  }

  // _ continue line, indent after, outdent after last _
  regex = /\s*( _)\b\s*$/gi;
  match = regex.exec(line);
  if (match !== null) {
    ret = getIndent(currentIndent) + line.trim();
    currentIndent += 3;
    underscoreCount += 3;
    underscored = true;
    // console.log(underscoreCount);
  } else if (underscored) {
    // console.log('Not undersored ' + line);
    underscored = false;
    currentIndent -= underscoreCount;
    underscoreCount = 0;
    // console.log('New indent:' + currentIndent);
  }

  // End of loop, outdent line
  regex = /^\s*(next|end if|#end if|wend|end with|loop)\b/gi;
  match = regex.exec(line);
  if (match !== null) {
    currentIndent--;
    ret = getIndent(currentIndent) + line.trim();
    // console.log('NextEndif Indent:' + currentIndent + ' ' + line);
  }

  // Else, outdent line and indent after
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

// #region Format keywords

/**
 * Format special, officially not supported line,
 * i.e. #region
 * @param  {} line
 */
function formatSpecialLine(line) {
  var ret = line;
  var mainRegexp;
  mainRegexp = /^\s*'\s*#\s*(endregion|region)\b(.*)$/gi;
  match = mainRegexp.exec(line);
  try {
    ret = '\'#' + match[1] + ' ' + match[2];
  } catch (error) {}
  return ret;
}

/**
 * Format commands
 * @param  {} line
 */
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

/**
 * Format built-in constants
 * @param  {} line
 */
function formatVBAConstant(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < constUp.length; k++) {
    regex = new RegExp('\\b' + constUp[k] + '\\b', 'gi');
    ret = ret.replace(regex, constUp[k]);
  }
  return ret;
}

/**
 * Format built-in messagebox constants
 * @param  {} line
 */
function formatVBAMsgConstant(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < msgConstUp.length; k++) {
    regex = new RegExp('\\b' + msgConstUp[k] + '\\b', 'gi');
    ret = ret.replace(regex, msgConstUp[k]);
  }
  return ret;
}

/**
 * Format functions
 * @param  {} line
 */
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

/**
 * Format variable and function types
 * @param  {} line
 */
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

/**
 * Format objects
 * @param  {} line
 */
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

/**
 * Format activities
 * @param  {} line
 */
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

/**
 * Format subobjects
 * @param  {} line
 */
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

// #endregion

/**
 * Format Const declaration line
 * @param  {} line
 */
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

// #region Format operators

/**
 * Add space around operators, except between ""
 * @param  {} line
 */
function addSpaceToOperators(line) {
  var ret = line;
  ret = ret.replace(
    /(>|<|=|\+|-|&|\/)(?=(?:[^"]*"[^"]*")*[^"]*$)/gi,
    replaceSingleOperator
  );
  ret = ret.replace(
    /(>|<|=)\s*(>|<|=)(?=(?:[^"]*"[^"]*")*[^"]*$)/gi,
    replaceDoubleOperator
  );
  return ret;
}
function replaceSingleOperator(str, group1) {
  return ' ' + group1 + ' ';
}
function replaceDoubleOperator(str, group1, group2) {
  return ' ' + group1 + group2 + ' ';
}

/**
 * Remove space around brackets
 * @param  {} line
 */
function removeSpaceAroundBrackets(line) {
  var ret = line;
  regex=/"\s*\(/;
  match=regex.exec(line);
  if (match === null){
    ret = ret.replace(/\s*\(\s*(?=(?:[^"]*"[^"]*")*[^"]*$)/gi, '(');
  }
  ret = ret.replace(/\s*\)(?=(?:[^"]*"[^"]*")*[^"]*$)/gi, ')');
  return ret;
}

// #endregion

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

/**
 * Clear textfields
 */
function clearCode() {
  $('#Code').val('');
  $('#CodeFormat').val('');
}

/**
 * Capitalize string, considering #
 * @param  {} string
 */
function capitalize(string) {
  var ret;
  if (string.charAt(0) === '#') {
    ret = '#' + string.charAt(1).toUpperCase() + string.slice(2).toLowerCase();
  } else {
    ret = string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
  }
  return ret;
}

/**
 * Create indent with spaces
 * @param  {} num Number of indent
 */
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
