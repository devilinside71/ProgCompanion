/// <reference path="../../typings/globals/jquery/index.d.ts" />
$(document).ready(function () {
    $('#generate').click(function () {
        generateCode();
    });
    $('#clear').click(function () {
        resetNames();
    });
});

/** Main function
 */
function generateCode() {
    var textString = "";
    if ($('#VBPredefDlg').val() == 'dlgOpenFile') {
        textString = textString + '    Dim FileNum As Integer\n';
        textString = textString + '    Dim DataLine As String\n';
        textString = textString + '    Dim sPath As String\n';
        textString = textString + '    Dim iChoice As Integer\n';
        textString = textString + '\n';
        textString = textString + '    ' + "'" + 'only allow the user to select one file\n';
        textString = textString + '    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False\n';
        textString = textString + '    ' + "'" + 'make the file dialog visible to the user\n';
        textString = textString + '    iChoice = Application.FileDialog(msoFileDialogOpen).Show\n';
        textString = textString + '    ' + "'" + 'determine what choice the user made\n';
        textString = textString + '    If iChoice <> 0 Then\n';
        textString = textString + '        ' + "'" + 'get the file path selected by the user\n';
        textString = textString + '        sPath = Application.FileDialog( _\n';
        textString = textString + '        msoFileDialogOpen).SelectedItems(1)\n';
        textString = textString + '        FileNum = FreeFile()\n';
        textString = textString + '        Open sPath For Input As #FileNum\n';
        textString = textString + '        \n';
        textString = textString + '        While Not EOF(FileNum)\n';
        textString = textString + '            Line Input #FileNum, DataLine ' + "'" + ' read in data 1 line at a time\n';
        textString = textString + '            ' + "'" + ' decide what to do with dataline,\n';
        textString = textString + '            ' + "'" + ' depending on what processing you need to do for each case\n';
        textString = textString + '        Wend\n';
        textString = textString + '    End If\n';
    }
    $('#Code').val(textString);
}


/** Capitalize the first letter of the text
 * @param  {string} text
 */
function capitalizeFirstLetter(text) {
    return text.charAt(0).toUpperCase() + text.slice(1);
}

/** Reset #Name objects' names
 */
function resetNames() {
    for (var i = 1; i < elemNum + 1; i++) {
        $('#Name' + i).val('');
    }
}