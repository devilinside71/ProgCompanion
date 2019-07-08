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
    if ($('#SBPredefDlg').val() == 'dlgOpenFile') {
        textString = textString + '' + "'" + 'https://www.debugpoint.com/2015/01/use-open-file-dialog-using-macro-in-libreofficeopenoffice/\n';
        textString = textString + 'Sub pick_a_file()\n';
        textString = textString + '    Dim fName As String\n';
        textString = textString + '    fName = open_file()\n';
        textString = textString + '    MsgBox fName & Chr(10) & ConvertFromUrl(fName)\n';
        textString = textString + 'End Sub\n';
        textString = textString + '\n';
        textString = textString + 'Function open_file() As String\n';
        textString = textString + '    \n';
        textString = textString + '    Dim file_dialog As Object\n';
        textString = textString + '    Dim status As Integer\n';
        textString = textString + '    Dim file_path As String\n';
        textString = textString + '    Dim init_path As String\n';
        textString = textString + '    Dim ucb As Object\n';
        textString = textString + '    Dim filterNames(3) As String\n';
        textString = textString + '    \n';
        textString = textString + '    filterNames(0) = "*.*"\n';
        textString = textString + '    filterNames(1) = "*.png"\n';
        textString = textString + '    filterNames(2) = "*.jpg"\n';
        textString = textString + '    \n';
        textString = textString + '    GlobalScope.BasicLibraries.LoadLibrary("Tools")\n';
        textString = textString + '    file_dialog = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")\n';
        textString = textString + '    ucb = createUnoService("com.sun.star.ucb.SimpleFileAccess")\n';
        textString = textString + '    \n';
        textString = textString + '    AddFiltersToDiaLog(FilterNames(), file_dialog)\n';
        textString = textString + '    ' + "'" + 'Set your initial path here!\n';
        textString = textString + '    init_path = ConvertToUrl("/usr")\n';
        textString = textString + '    \n';
        textString = textString + '    If ucb.Exists(init_path) Then\n';
        textString = textString + '        file_dialog.SetDisplayDirectory(init_path)\n';
        textString = textString + '    End If\n';
        textString = textString + '    \n';
        textString = textString + '    status = file_dialog.Execute()\n';
        textString = textString + '    If status = 1 Then\n';
        textString = textString + '        file_path = file_dialog.Files(0)\n';
        textString = textString + '        open_file = file_path\n';
        textString = textString + '    End If\n';
        textString = textString + '    file_dialog.Dispose()\n';
        textString = textString + '    \n';
        textString = textString + 'End Function\n';
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