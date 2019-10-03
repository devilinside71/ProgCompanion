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
  '#If', '#Else', '#End If');
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
  var remVals;
  for (i = 0; i < lines.length; i++) {
    line = lines[i].trim();
    remVals = formatREMlIne(line);
    line = remVals[0];
    // console.log(remVals[1]);
    if (remVals[1] === false) {
      // console.log(line);
      line = removeSpaces(line);
      line = formatDeclarationLine(line);
      line = formatConstDeclarationLine(line);
      line = formatSubLine(line);
      line = formatFuncLine(line);
      // line = formatIfLine(line);
      line = formatCommand(line);
    }
    outText += line + '\n';
  }
  $('#CodeFormat').val(outText);
}

function formatCommand(line) {
  var k;
  var ret = line;
  var regex;
  for (k = 0; k < commandsUp.length; k++) {
    regex = new RegExp('\\b' + commandsUp[k] + '\\b', 'gi');
    ret = ret.replace(regex, commandsUp[k]);
    // console.log('\b' + commandsUp[k].toLocaleLowerCase() + '\b');
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
  console.log('Try OPTIONAL ' + line);
  mainRegexp = /(\boptional\b)\s*(.*)\s*=\s*(.*)\s*/gi;
  mainMatch = mainRegexp.exec(line);
  try {
    console.log('MOPT:' + mainMatch.length);
    ret = 'Optional ' + mainMatch[2] + ' = ' + mainMatch[3];
  } catch (error) {
    console.log('Not OPTIONAL');
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
function formatDeclarationLine(line) {
  var ret = line;
  var subret = '';
  var mainRegexp;
  var subRegexp;
  var mainMatch;
  var subMatch;
  var subElems;
  var index;
  var elem;
  mainRegexp = /(dim |private |global )(?!.*\s*const\s*)(?!.*\s*sub\s*)(?!.*\s*function\s*)(.*)/gi;
  mainMatch = mainRegexp.exec(line);
  try {
    subElems = mainMatch[2].split(',');
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
        // console.log('hiba ' + subMatch[1]);
      }
      // subret += elem;
    }
    // console.log(mainMatch.length);
    ret = capitalize(mainMatch[1]) + subret;
    // console.log(ret);
  } catch (error) {
    // console.log('nem DECLAR ' + line);
  }
  return ret;
}
function formatREMlIne(line) {
  var ret = line;
  var retVal = false;
  if (line.substring(0, 1) === '\'') {
    ret = line;
    retVal = true;
    // console.log('REM line:' + line);
  }
  if (line.substring(0, 4).toLowerCase() === 'rem ') {
    ret = 'REM ' + line.substring(4);
    retVal = true;
    // console.log('REM line:' + line);
  }
  return [ret, retVal];
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
