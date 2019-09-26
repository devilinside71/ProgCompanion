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
  var textString = '';
  if ($('#VBPredefDlg').val() === 'dlgOpenFile') {
    textString += '    Dim FileNum As Integer\n';
    textString += '    Dim DataLine As String\n';
    textString += '    Dim sPath As String\n';
    textString += '    Dim iChoice As Integer\n';
    textString += '\n';
    textString += '    \'only allow the user to select one file\n';
    textString +=
      '    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False\n';
    textString += '    \'make the file dialog visible to the user\n';
    textString +=
      '    iChoice = Application.FileDialog(msoFileDialogOpen).Show\n';
    textString += '    \'determine what choice the user made\n';
    textString += '    If iChoice <> 0 Then\n';
    textString += '        \'get the file path selected by the user\n';
    textString += '        sPath = Application.FileDialog( _\n';
    textString += '        msoFileDialogOpen).SelectedItems(1)\n';
    textString += '        FileNum = FreeFile()\n';
    textString += '        Open sPath For Input As #FileNum\n';
    textString += '        \n';
    textString += '        While Not EOF(FileNum)\n';
    textString =
      textString +
      '            Line Input #FileNum, DataLine ' +
      '\'' +
      ' read in data 1 line at a time\n';
    textString += '            \' decide what to do with dataline,\n';
    textString =
      textString +
      '            ' +
      '\'' +
      ' depending on what processing you need to do for each case\n';
    textString += '        Wend\n';
    textString += '    End If\n';
  }
  $('#Code').val(textString);
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
