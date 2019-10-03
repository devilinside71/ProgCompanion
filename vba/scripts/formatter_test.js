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

    outText += line + '\n';
  }
  $('#CodeFormat').val(outText);
}


function clearCode() {
  $('#Code').val('');
  $('#CodeFormat').val('');
}

