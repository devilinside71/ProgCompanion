$(document).ready(function() {
  $('#convert').click(function() {
    convertText();
  });
  $('#clear').click(function() {
    clearCode();
  });
});

function convertText() {
  var varName = '';
  var outText = '';
  var line = '';
  var inText = $('#Code').val();
  var lines = inText.split('\n');
  var i = 0;
  var strPrefix = 'str';
  if ($('#ShortPrefix').is(':checked')) {
    strPrefix = 's';
  } else {
    strPrefix = 'str';
  }
  if ($('#langChoice').val() === 'VBA') {
    varName = strPrefix + capitalizeFirstLetter($('#textVar').val());
    outText = 'Dim ' + varName + ' As String\n';
    outText = outText + varName + ' = vbNullString\n';
    for (i = 0; i < lines.length; i++) {
      line = lines[i];
      line = replaceAll(line, '"', '" & Chr(34) & "');
      outText =
        outText +
        varName +
        ' = ' +
        varName +
        ' & ' +
        '"' +
        line +
        '" & vbCrLf\n';
    }
  }
  if ($('#langChoice').val() === 'SB') {
    varName = strPrefix + capitalizeFirstLetter($('#textVar').val());
    outText = 'Dim ' + varName + ' As String\n';
    outText = outText + varName + ' = ""\n';
    for (i = 0; i < lines.length; i++) {
      line = lines[i];
      line = replaceAll(line, '"', '" & Chr(34) & "');
      outText =
        outText + varName + ' = ' + varName + ' & "' + line + '\\n"\n';
    }
  }
  if ($('#langChoice').val() === 'XML') {
    varName = decapitalizeFirstLetter($('#textVar').val());
    outText = '<' + varName.toLowerCase() + '>\n';
    for (i = 0; i < lines.length; i++) {
      line = lines[i];
      line = replaceAll(line, '&', '&amp;');
      line = replaceAll(line, '<', '&lt;');
      line = replaceAll(line, '>', '&gt;');
      outText = outText + line + '\n';
    }
    outText = outText + '</' + varName.toLowerCase() + '>\n';
  }
  $('#CodeFormat').val(outText);
}

function clearCode() {
  $('#Code').val('');
  $('#CodeFormat').val('');
}

function capitalizeFirstLetter(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

function decapitalizeFirstLetter(string) {
  return string.charAt(0).toLowerCase() + string.slice(1);
}

function replaceAll(str, find, replace) {
  return str.replace(new RegExp(find, 'g'), replace);
}
