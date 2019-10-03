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
  var lines = $('#Code')
    .val()
    .split('\n');
  var outText = '';
  var line = '';
  var i = 0;
  for (i = 0; i < lines.length; i++) {
    line = lines[i].trim();
    if (isRemLine(line)) {
      console.log('Remline: ' + line);
      if (line.substring(0, 4).toLowerCase() === 'rem ') {
        line = 'REM ' + line.substring(4);
      }
    }
    outText += line + '\n';
  }
  $('#CodeFormat').val(outText);
}

function isRemLine(line) {
  var ret = false;
  if (
    line.substring(0, 1) === '\'' ||
    line.substring(0, 4).toLowerCase() === 'rem '
  ) {
    ret = true;
  }
  return ret;
}

function clearCode() {
  $('#Code').val('');
  $('#CodeFormat').val('');
}
