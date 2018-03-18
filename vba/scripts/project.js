/// <reference path="../../typings/globals/jquery/index.d.ts" />
var menuNum = 8;
//min 5 properties
var propNum = 8;

$(document).ready(function () {
    var i = 1;
    var typeOptions = new Array(
        "String", "Long", "Integer", "Boolean", "Double", "Date", "Variant",
        "Object", "SheetName", "Worksheet", "Outlook"
    );
    for (i = 1; i < menuNum + 1; i++) {
        $("#tableMenu").find('tbody')
            .append($('<tr>')
                .append($('<td class="Col1">Menu' + i + '</td>'))
                .append($(
                    '<td class="checkBoxCol">' +
                    '<input type="checkbox" id="MenuCheck' +
                    i + '" /></td>'))
                .append($('<td class="Col2"><input id="Cap' +
                    i + '" type="text" value="Menu' + i +
                    '" class="textBox" /></td>'))
                .append($('<td class="Col3"><input id="Face' +
                    i + '" type="text" value="1244" class="textBox" /></td>'))
                .append($('<td class="Col4"><input id="Onaction' + i +
                    '" type="text" value="MenuNULL" class="textBox" /></td>'))
                .append($('<td class="Col5"><input id="TTip' +
                    i + '" type="text" value="Menu' + i +
                    ' végrehajtása" class="textBox" /></td>'))
            );
    }
    for (i = 1; i < propNum + 1; i++) {
        $("#tableProps").find('tbody')
            .append($('<tr>')
                .append($('<td class="Col1">Property' + i + '</td>'))
                .append($(
                    '<td class="checkBoxCol">' +
                    '<input type="checkbox" id="ClassCheck' +
                    i + '" /></td>'))
                .append($('<td class="Col2"><input id="Clprop' + i +
                    '" type="text" value="Prop' + i +
                    '" class="textBox" /></td>'))
                .append($('<td class="Col3"><input id="Clpar' + i +
                    '" type="text" value="Par' + i +
                    '" class="textBox" /></td>'))
                .append($('<td class="Col4"><select id="Type' + i + '">'))
                .append($('<td class="Col5"><select id="Mode' + i + '">'))
            );
    }
    for (i = 1; i < propNum + 1; i++) {
        $('#Type' + i)
            .append($('<option value="' + typeOptions[0] + '">' +
                typeOptions[0] + '</option>'))
            .append($('<option value="' + typeOptions[1] + '">' +
                typeOptions[1] + '</option>'))
            .append($('<option value="' + typeOptions[2] + '">' +
                typeOptions[2] + '</option>'))
            .append($('<option value="' + typeOptions[3] + '">' +
                typeOptions[3] + '</option>'))
            .append($('<option value="' + typeOptions[4] + '">' +
                typeOptions[4] + '</option>'))
            .append($('<option value="' + typeOptions[5] + '">' +
                typeOptions[5] + '</option>'))
            .append($('<option value="' + typeOptions[6] + '">' +
                typeOptions[6] + '</option>'))
            .append($('<option value="' + typeOptions[7] + '">' +
                typeOptions[7] + '</option>'))
            .append($('<option value="' + typeOptions[8] + '">' +
                typeOptions[8] + '</option>'))
            .append($('<option value="' + typeOptions[9] + '">' +
                typeOptions[9] + '</option>'))
            .append($('<option value="' + typeOptions[10] + '">' +
                typeOptions[10] + '</option>'));
    }
    for (i = 1; i < propNum + 1; i++) {
        $('#Mode' + i)
            .append($('<option value="Let and Get">Let and Get</option>'))
            .append($('<option value="Let">Let</option>'))
            .append($('<option value="Get">Get</option>'));

    }


    $('#createproject').click(function () {
        createProject();
    });
    $('#clear').click(function () {
        resetNames();
    });

    $('#createclass').click(function () {
        createClass();
    });
    $('#setpredef').click(function () {
        setPredef();
    });
    $('#NameProj').on('input', (function () {
        $('#NameClass').val($('#NameProj').val() + "Class");
    }));
});

function createProject() {
    createMenu();
    createMainModule();
    createClass();
}

function resetNames() {
    for (var i = 1; i < menuNum + 1; i++) {
        $('#Cap' + i).val("Button" + i);
        $('#Face' + i).val("1244");
        $('#Onaction' + i).val("MenuNULL");
        $('#TTip' + i).val("Button" + i + " végrahajtása");
    }
}

function createClass() {
    var className = capitalizeFirstLetter($('#NameClass').val());
    //Classtest
    var outtextClasstest = "";
    var i = 1;
    outtextClasstest = outtextClasstest + 'Private Sub ' +
        className + '_ClassTest()\n';
    outtextClasstest = outtextClasstest + '    Dim cl' +
        className + ' As New ' + className + '\n';
    outtextClasstest = outtextClasstest + '    \n';
    for (i = 1; i < propNum + 1; i++) {
        if ($('#ClassCheck' + 1).prop('checked')) {
            outtextClasstest = outtextClasstest + '    cl' +
                className + '.' + $('#Clprop' + i).val() +
                getConstInitValue($('#Type' + i).val()) + '\n';
            outtextClasstest = outtextClasstest + '    Debug.Print "cl' +
                className + '.' + $('#Clprop' + i).val() +
                ': " & cl' + className + '.' + $('#Clprop' + i).val() + '\n';
        }
    }

    outtextClasstest = outtextClasstest + '    Set cl' +
        className + ' = Nothing\n';
    outtextClasstest = outtextClasstest + 'End Sub\n';
    $('#textMain').val($('#textMain').val() + outtextClasstest);

    //Definition
    var outtxtClass = "";
    outtxtClass = outtxtClass + 'Option Explicit\n';
    outtxtClass = outtxtClass + '\n';

    for (i = 1; i < propNum + 1; i++) {
        if ($('#ClassCheck' + 1).prop('checked')) {
            outtxtClass = outtxtClass + 'Private m_' +
                getPrefix($('#Type' + i).val()) +
                $('#Clprop' + i).val() + ' As ' +
                getDeclareType($('#Type' + i).val()) + '\n';
            outtxtClass = outtxtClass + 'Private Const cm' +
                getPrefix($('#Type' + i).val()) +
                $('#Clprop' + i).val() + ' As ' +
                getDeclareType($('#Type' + i).val()) +
                getConstInitValue($('#Type' + i).val()) + '\n';
        }
    }

    //Properties
    for (i = 1; i < propNum + 1; i++) {
        if ($('#ClassCheck' + i).prop('checked')) {
            if ($('#Mode' + i).val() ==
                "Let and Get" || $('#Mode' + i).val() == "Let") {
                outtxtClass = outtxtClass + 'Public Property Let ' +
                    $('#Clprop' + i).val() +
                    '(' + $('#Clpar' + i).val() + ' As ' +
                    getDeclareType($('#Type' + i).val()) + ')\n';
                outtxtClass = outtxtClass + '    \n';
                outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
                outtxtClass = outtxtClass + '    \n';
                outtxtClass = outtxtClass + '    m_' +
                    getPrefix($('#Type' + i).val()) +
                    $('#Clprop' + i).val() + ' = ' +
                    $('#Clpar' + i).val() + '\n';
                outtxtClass = outtxtClass + '    Debug.Print "' +
                    className + '.' + $('#Clprop' + i).val() +
                    ' has been set to: " & m_' +
                    getPrefix($('#Type' + i).val()) +
                    $('#Clprop' + i).val() + '\n';
                outtxtClass = outtxtClass + '    \n';
                outtxtClass = outtxtClass + 'PROC_EXIT:\n';
                outtxtClass = outtxtClass + '    Exit Property\n';
                outtxtClass = outtxtClass + '    \n';
                outtxtClass = outtxtClass + 'PROC_ERR:\n';
                outtxtClass = outtxtClass + '    Err.Raise Err.Number\n';
                outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
                outtxtClass = outtxtClass + 'End Property\n';
            }
            if ($('#Mode' + i).val() ==
                "Let and Get" || $('#Mode' + i).val() == "Get") {
                outtxtClass = outtxtClass + '\n';
                outtxtClass = outtxtClass + 'Public Property Get ' +
                    $('#Clprop' + i).val() + '() As ' +
                    getDeclareType($('#Type' + i).val()) + '\n';
                outtxtClass = outtxtClass + '    \n';
                outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
                outtxtClass = outtxtClass + '    \n';
                outtxtClass = outtxtClass + '    ' + $('#Clprop' + i).val() +
                    ' = m_' + getPrefix($('#Type' + i).val()) +
                    $('#Clprop' + i).val() + '\n';
                outtxtClass = outtxtClass + '    \n';
                outtxtClass = outtxtClass + 'PROC_EXIT:\n';
                outtxtClass = outtxtClass + '    Exit Property\n';
                outtxtClass = outtxtClass + '    \n';
                outtxtClass = outtxtClass + 'PROC_ERR:\n';
                outtxtClass = outtxtClass + '    Err.Raise Err.Number\n';
                outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
                outtxtClass = outtxtClass + 'End Property\n';
            }
        }
    }

    //Class body
    outtxtClass = outtxtClass + 'Private Sub Class_Initialize()\n';
    outtxtClass = outtxtClass + '    Debug.Print "Class ' +
        className + ' initialized"\n';
    outtxtClass = outtxtClass + '    \n';
    for (i = 1; i < propNum + 1; i++) {
        if ($('#ClassCheck' + i).prop('checked')) {
            outtxtClass = outtxtClass + '    m_' +
                getPrefix($('#Type' + i).val()) + $('#Clprop' + i).val() +
                ' = cm' + getPrefix($('#Type' + i).val()) +
                $('#Clprop' + i).val() + '\n';
            outtxtClass = outtxtClass + '    Debug.Print "' +
                className + ' Default value for ' + $('#Clprop' + i).val() +
                ': " & m_' + getPrefix($('#Type' + i).val()) +
                $('#Clprop' + i).val() + '\n';
        }

    }

    outtxtClass = outtxtClass + 'End Sub\n';
    outtxtClass = outtxtClass + 'Private Sub Class_Terminate()\n';
    outtxtClass = outtxtClass + '    Debug.Print "Class ' +
        className + ' terminated"\n';
    outtxtClass = outtxtClass + 'End Sub\n';
    outtxtClass = outtxtClass + 'Sub Reset()\n';
    outtxtClass = outtxtClass + '    \n';

    for (i = 1; i < propNum + 1; i++) {
        if ($('#ClassCheck' + i).prop('checked')) {
            outtxtClass = outtxtClass + '    m_' +
                getPrefix($('#Type' + i).val()) + $('#Clprop' + i).val() +
                ' = cm' + getPrefix($('#Type' + i).val()) +
                $('#Clprop' + i).val() + '\n';
            outtxtClass = outtxtClass + '    Debug.Print "' +
                className + ' Default value for ' + $('#Clprop' + i).val() +
                ': " & m_' + getPrefix($('#Type' + i).val()) +
                $('#Clprop' + i).val() + '\n';
        }
    }

    outtxtClass = outtxtClass + 'End Sub\n';


    //Others
    if ($('#ClassCompCheck1').prop('checked')) {
        outtxtClass = outtxtClass + '' + "'" + '----------------\n';
        outtxtClass = outtxtClass + '' + "'" + 'Columns and Rows\n';
        outtxtClass = outtxtClass + '' + "'" + '----------------\n';
        outtxtClass = outtxtClass +
            'Private Function Col_Letter(lngCol As Long) As String\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Get letter from column number\n';
        outtxtClass = outtxtClass + '    Dim vArr\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    ' + "'" + '  On Error Resume Next\n';
        outtxtClass = outtxtClass +
            '    vArr = Split(Cells(1, lngCol).Address(True, False), "$")\n';
        outtxtClass = outtxtClass + '    Col_Letter = vArr(0)\n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass +
            'Private Function Col_LetterHeader(sheetName As String,' +
            ' headText As String, Optional headRow = 1) As String\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Get column letter from header text\n';
        outtxtClass = outtxtClass + '    Dim lngColNumber As Long\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    lngColNumber = ' +
            'Col_NumberHeader(sheetName, headText, headRow)\n';
        outtxtClass = outtxtClass + '    Col_LetterHeader = ' +
            'Col_Letter(lngColNumber)\n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass +
            'Private Function Col_Number(colLetter) As Long\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Get column number from column letter\n';
        outtxtClass = outtxtClass +
            '    Col_Number = Range(colLetter & "1").Column\n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass +
            'Private Function Col_NumberHeader(sheetName As String, ' +
            'headText As String, Optional headRow = 1) As Long\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Get column number from header text\n';
        outtxtClass = outtxtClass + '    Dim i As Long\n';
        outtxtClass = outtxtClass + '    Dim strCellString As String\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Col_NumberHeader = 0\n';
        outtxtClass = outtxtClass + '    For i = 1 To 400\n';
        outtxtClass = outtxtClass +
            '        strCellString = ' +
            'Trim(CStr(Sheets(sheetName).Cells(headRow, i)))\n';
        outtxtClass = outtxtClass +
            '        If strCellString = headText Then\n';
        outtxtClass = outtxtClass + '            Col_NumberHeader = i\n';
        outtxtClass = outtxtClass + '            Exit Function\n';
        outtxtClass = outtxtClass + '        End If\n';
        outtxtClass = outtxtClass + '    Next i\n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass + 'Private Sub ColLetterTests()\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Test for Col_Letter, Col_LetterHeader, ' +
            'Col_Number and Col_NumberHeader\n';
        outtxtClass = outtxtClass +
            '    Debug.Print Col_Letter(12)\n';
        outtxtClass = outtxtClass +
            '    Debug.Print Col_LetterHeader("Hogyallunk", "Any.csop.")\n';
        outtxtClass = outtxtClass + '    Debug.Print Col_Number("H")\n';
        outtxtClass = outtxtClass +
            '    Debug.Print Col_NumberHeader("Hogyallunk", "Any.csop.")\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass +
            'Private Function GetLastRow(sheetName As String, ' +
            'checkColumn As Long, _\n';
        outtxtClass = outtxtClass +
            '    Optional firstrow = 2, Optional lastrow = 600000, _\n';
        outtxtClass = outtxtClass +
            '        Optional backwardCheck = True) As Long\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Adott fül utolsó sora\n';
        outtxtClass = outtxtClass + '    Dim i As Long\n';
        outtxtClass = outtxtClass + '    Dim curSheet As Worksheet\n';
        outtxtClass = outtxtClass + '    Dim strCell As String\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass +
            '    Set curSheet = ActiveWorkbook.ActiveSheet\n';
        outtxtClass = outtxtClass + '    Sheets(sheetName).Activate\n';
        outtxtClass = outtxtClass + '    GetLastRow = 0\n';
        outtxtClass = outtxtClass + '    If backwardCheck Then\n';
        outtxtClass = outtxtClass +
            '        For i = lastrow To firstrow Step -1\n';
        outtxtClass = outtxtClass +
            '            strCell = Trim(CStr(Cells(i, checkColumn)))\n';
        outtxtClass = outtxtClass + '            If strCell <> "" Then\n';
        outtxtClass = outtxtClass + '                GetLastRow = i\n';
        outtxtClass = outtxtClass + '                Exit For\n';
        outtxtClass = outtxtClass + '            End If\n';
        outtxtClass = outtxtClass + '        Next i\n';
        outtxtClass = outtxtClass + '    Else\n';
        outtxtClass = outtxtClass + '        For i = firstrow To lastrow\n';
        outtxtClass = outtxtClass +
            '            strCell = Trim(CStr(Cells(i, checkColumn)))\n';
        outtxtClass = outtxtClass + '            If strCell = "" Then\n';
        outtxtClass = outtxtClass + '                GetLastRow = i - 1\n';
        outtxtClass = outtxtClass + '                Exit For\n';
        outtxtClass = outtxtClass + '            End If\n';
        outtxtClass = outtxtClass + '        Next i\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    curSheet.Activate\n';
        outtxtClass = outtxtClass + '    Set curSheet = Nothing\n';
        outtxtClass = outtxtClass +
            '    Debug.Print "LastRow of " & sheetName & ": " &' +
            ' GetLastRow & " ChkCol:" & checkColumn\n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass + '\n';
        outtxtClass = outtxtClass + '\n';

    }
    if ($('#ClassCompCheck2').prop('checked')) {
        outtxtClass = outtxtClass + '' + "'" + '--------------\n';
        outtxtClass = outtxtClass + '' + "'" + 'Refresh ON OFF\n';
        outtxtClass = outtxtClass + '' + "'" + '--------------\n';
        outtxtClass = outtxtClass + 'Private Sub RefreshOFF()\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Screen update OFF\n';
        outtxtClass = outtxtClass + '    With Application\n';
        outtxtClass = outtxtClass + '        .ScreenUpdating = False\n';
        outtxtClass = outtxtClass + '        .EnableEvents = False\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            '.Calculation = xlCalculationManual\n';
        outtxtClass = outtxtClass + '    End With\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass + 'Private Sub RefreshON()\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Screen update ON\n';
        outtxtClass = outtxtClass + '    With Application\n';
        outtxtClass = outtxtClass + '        .ScreenUpdating = True\n';
        outtxtClass = outtxtClass + '        .EnableEvents = True\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            '.Calculation = xlCalculationAutomatic\n';
        outtxtClass = outtxtClass + '    End With\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass + '\n';
    }

    //SAP body
    if ($('#ClassCompCheck3').prop('checked')) {
        outtxtClass = outtxtClass +
            'Private Function GetPathOfFile(FullFilename ' +
            'As String) As String\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Comments  : Get path of a full filename with path\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Parameters:\n';
        outtxtClass = outtxtClass + '    ' + "'" + '' + "'" +
            '           FullFilename\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Dim strRes As String\n';
        outtxtClass = outtxtClass + '    Dim fso\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo FUNC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass +
            '    Set fso = CreateObject("Scripting.FileSystemObject")\n';
        outtxtClass = outtxtClass +
            '    strRes = fso.GetParentFolderName(FullFilename)\n';
        outtxtClass = outtxtClass + '    GetPathOfFile = strRes & "\"\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass +
            '    ' + "'" + '-------------------------------\n';
        outtxtClass = outtxtClass + '    FUNC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    Set fso = Nothing\n';
        outtxtClass = outtxtClass + '    Exit Function\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    FUNC_ERR:\n';
        outtxtClass = outtxtClass +
            '    Debug.Print "Error in Function %PROPNAME%.GetPathOfFile"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass +
            '        ' + "'" + 'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume FUNC_EXIT\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass +
            'Private Function GetFilenameOfFile(FullFilename ' +
            'As String) As String\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Comments  : Get filename of a full filename with path\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Parameters:\n';
        outtxtClass = outtxtClass + '    ' + "'" + '' + "'" +
            '           FullFilename\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Dim strRes As String\n';
        outtxtClass = outtxtClass + '    Dim fso\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo FUNC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass +
            '    Set fso = CreateObject("Scripting.FileSystemObject")\n';
        outtxtClass = outtxtClass +
            '    strRes = fso.GetFileName(FullFilename)\n';
        outtxtClass = outtxtClass + '    GetFilenameOfFile = strRes\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    ' + "'" +
            '-------------------------------\n';
        outtxtClass = outtxtClass + '    FUNC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    Set fso = Nothing\n';
        outtxtClass = outtxtClass + '    Exit Function\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    FUNC_ERR:\n';
        outtxtClass = outtxtClass + '    Debug.Print "Error in Function ' +
            className + '.GetFilenameOfFile"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume FUNC_EXIT\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass + '\n';
        outtxtClass = outtxtClass + 'Sub AddSheet()\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Comments  : Add zhogyallunk sheet\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Parameters:\n';
        outtxtClass = outtxtClass + '    ' + "'" + '' + "'" + '\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Created by: Laszlo Tamas\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    Sheets.Add Before:=Sheets(1)\n';
        outtxtClass = outtxtClass + '    ActiveSheet.Name = m_strSheetname\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Exit Sub\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass + '    Debug.Print "Error in Sub ' +
            className + '.AddSheet"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass + '\n';
        outtxtClass = outtxtClass +
            'Sub DeleteSheet(Optional Alert As Boolean = False)\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Comments  : Remarks\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Parameters:\n';
        outtxtClass = outtxtClass + '    ' + "'" + '' + "'" +
            '           Alert\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Created by: Laszlo Tamas\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    If Alert = False Then\n';
        outtxtClass = outtxtClass +
            '        Application.DisplayAlerts = False\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    On Error Resume Next\n';
        outtxtClass = outtxtClass + '    Sheets(m_strSheetname).Delete\n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    Application.DisplayAlerts = True\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Exit Sub\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass + '    Debug.Print "Error in Sub ' +
            className + '.DeleteSheet"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass +
            'Sub CreateSheet(Optional Alert As Boolean = False)\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Comments  : Remarks\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Parameters:\n';
        outtxtClass = outtxtClass + '    ' + "'" + '' + "'" +
            '           Alert\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Created by: Laszlo Tamas\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    Call DeleteSheet(Alert)\n';
        outtxtClass = outtxtClass + '    Call AddSheet\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Exit Sub\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass + '    Debug.Print "Error in Sub ' +
            className + '.CreateSheet"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass +
            'Sub DuplicateSheet(NewSheetname As String, ' +
            'Optional Alert As Boolean = False)\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Comments  : Duplicate full hogyallunk sheet\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Parameters:\n';
        outtxtClass = outtxtClass + '    ' + "'" + '' + "'" +
            '           Alert\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Created by: Laszlo Tamas\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    If Alert = False Then\n';
        outtxtClass = outtxtClass +
            '        Application.DisplayAlerts = False\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    On Error Resume Next\n';
        outtxtClass = outtxtClass + '    Sheets(NewSheetname).Delete\n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    Application.DisplayAlerts = True\n';
        outtxtClass = outtxtClass + '    Sheets(m_strSheetname).Select\n';
        outtxtClass = outtxtClass +
            '    Sheets(m_strSheetname).Copy Before:=Sheets(m_strSheetname)\n';
        outtxtClass = outtxtClass + '    ActiveSheet.Name = NewSheetname\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Exit Sub\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass + '    Debug.Print "Error in Sub ' +
            className + '.DuplicateSheet"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass +
            'Sub CopySheetContent(NewSheetname As String, ' +
            'PasteSpec As Boolean, Optional Alert As Boolean = False)\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Comments  : Copy the (filtered, viewable) ' +
            'content of hogyallunk sheet\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Parameters:\n';
        outtxtClass = outtxtClass + '    ' + "'" + '' + "'" +
            '           Alert\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Created by: Laszlo Tamas\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    If Alert = False Then\n';
        outtxtClass = outtxtClass +
            '        Application.DisplayAlerts = False\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    On Error Resume Next\n';
        outtxtClass = outtxtClass + '    Sheets(NewSheetname).Delete\n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    Application.DisplayAlerts = True\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass +
            '    Sheets.Add Before:=Sheets(m_strSheetname)\n';
        outtxtClass = outtxtClass + '    ActiveSheet.Name = NewSheetname\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Sheets(m_strSheetname).Select\n';
        outtxtClass = outtxtClass + '    Cells.Select\n';
        outtxtClass = outtxtClass + '    Selection.Copy\n';
        outtxtClass = outtxtClass + '    Sheets(NewSheetname).Select\n';
        outtxtClass = outtxtClass + '    Range("A1").Select\n';
        outtxtClass = outtxtClass + '    If PasteSpec Then\n';
        outtxtClass = outtxtClass +
            '        Selection.PasteSpecial Paste:=xlPasteValues, ' +
            'Operation:=xlNone, SkipBlanks _\n';
        outtxtClass = outtxtClass + '            :=False, Transpose:=False\n';
        outtxtClass = outtxtClass + '    Else\n';
        outtxtClass = outtxtClass + '        ActiveSheet.Paste\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Exit Sub\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass + '    Debug.Print "Error in Sub ' +
            className + '.CopySheetContent"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass +
            'Function GetOpenFilename(Optional Title = "Fájl", ' +
            'Optional CollectionName = ' +
            '"Fájlok", Optional Extensions = "*.*") ' + 'As String\n';
        outtxtClass = outtxtClass + '    ' + "'" + ' Comments:\n';
        outtxtClass = outtxtClass + '    ' + "'" + ' Params  : Title\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            '           CollectionName\n';
        outtxtClass = outtxtClass + '    ' + "'" + '           Extensions\n';
        outtxtClass = outtxtClass + '    ' + "'" + ' Returns : String\n';
        outtxtClass = outtxtClass + '    ' + "'" + ' Modified:\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Dim intChoice As Integer\n';
        outtxtClass = outtxtClass + '    Dim strPath As String\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    intChoice = 0\n';
        outtxtClass = outtxtClass + '    strPath = vbNullString\n';
        outtxtClass = outtxtClass +
            '    With Application.FileDialog(msoFileDialogOpen)\n';
        outtxtClass = outtxtClass +
            '        .Title = Title & " kiválasztása"\n';
        outtxtClass = outtxtClass +
            '        .Filters.Add CollectionName, Extensions\n';
        outtxtClass = outtxtClass + '        .FilterIndex = .Filters.Count\n';
        outtxtClass = outtxtClass + '        .AllowMultiSelect = False\n';
        outtxtClass = outtxtClass + '        \n';
        outtxtClass = outtxtClass + '        intChoice = .Show\n';
        outtxtClass = outtxtClass + '        If intChoice <> 0 Then\n';
        outtxtClass = outtxtClass + '            strPath = .SelectedItems(1)\n';
        outtxtClass = outtxtClass + '        End If\n';
        outtxtClass = outtxtClass + '    End With\n';
        outtxtClass = outtxtClass + '    GetOpenFilename = strPath\n';
        outtxtClass = outtxtClass + '    ' + "'" + '    Debug.Print "' +
            className +
            '.GetOpenFilename: " & GetOpenFilename & " Title:" & Title\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    Exit Function\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass +
            '    Debug.Print Err.Description, vbCritical, "' +
            className + '.GetOpenFilename"\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass +
            'Sub SetFilenameWithDialog(Optional Title = "Fájl", ' +
            'Optional CollectionName = "Fájlok", ' +
            'Optional Extensions = "*.*")\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Comments  : Load data from SAP exported CSV (text)\n';
        outtxtClass = outtxtClass + '    ' + "'" + ' Params  : Title\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            '           CollectionName\n';
        outtxtClass = outtxtClass + '    ' + "'" + '           Extensions\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Created by: Laszlo Tamas\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Dim strRes As String\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass +
            '    strRes = GetOpenFilename(Title, CollectionName, Extensions)\n';
        outtxtClass = outtxtClass + '    m_strFullFilename = strRes\n';
        outtxtClass = outtxtClass + '    Debug.Print "' +
            className + '.SetFilenameWithDialog FullFilename ' +
            'has been set to: " & m_strFullFilename\n';
        outtxtClass = outtxtClass + '    m_strPath = GetPathOfFile(strRes)\n';
        outtxtClass = outtxtClass + '    Debug.Print "' +
            className + '.SetFilenameWithDialog Path ' +
            'has been set to: " & m_strPath\n';
        outtxtClass = outtxtClass + '    m_strFilename = ' +
            'GetFilenameOfFile(strRes)\n';
        outtxtClass = outtxtClass + '    Debug.Print "' +
            className + '.SetFilenameWithDialog Filename ' +
            'has been set to: " & m_strFilename\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Exit Sub\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass + '    Debug.Print "Error in Sub ' +
            className + '.SetFilenameWithDialog"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass +
            'Function GetSaveFilename(Optional DTitle = "Fájl", ' +
            'Optional FFilter = "Excel files , *.xlsx") As String\n';
        outtxtClass = outtxtClass + '    ' + "'" + ' Comments:\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    GetSaveFilename = vbNullString\n';
        outtxtClass = outtxtClass +
            '    GetSaveFilename = Application.GetSaveAsFilename(' +
            'InitialFileName:=m_strOutputFilename, ' +
            'FileFilter:=FFilter, Title:=DTitle)\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            '    Debug.Print "GetOpenFilename: " & ' +
            'GetSaveFilename & " Title:" & Title\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    Exit Function\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass +
            '    Debug.Print Err.Description, vbCritical, "' +
            className + '.GetSaveFilename"\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + 'End Function\n';
        outtxtClass = outtxtClass +
            'Sub SetOutputFilenameWithDialog(Optional DTitle = ' +
            '"Fájl", Optional FFilter = "Excel files , *.xlsx")\n';
        outtxtClass = outtxtClass + '    ' + "'" +
            'Comments  : Load data from SAP exported CSV (text)\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Created by: Laszlo Tamas\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Dim strRes As String\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass +
            '    strRes = GetSaveFilename(DTitle, FFilter)\n';
        outtxtClass = outtxtClass + '    m_strOutputFilename = strRes\n';
        outtxtClass = outtxtClass +
            '    Debug.Print "' + className +
            '.SetOutputFilenameWithDialog FullFilename ' +
            'has been set to: " & m_strOutputFilename\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Exit Sub\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass + '    Debug.Print "Error in Sub ' +
            className + '.SetOutputFilenameWithDialog"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass + '        ' + "'" +
            'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + 'End Sub\n';
        outtxtClass = outtxtClass +
            'Sub LoadText(Optional PasteSpec As Boolean = False)\n';
        outtxtClass = outtxtClass + '    On Error GoTo PROC_ERR\n';
        outtxtClass = outtxtClass + '    Sheets(m_strSheetname).Select\n';
        outtxtClass = outtxtClass + '    Cells.Select\n';
        outtxtClass = outtxtClass + '    Selection.NumberFormat = "General"\n';
        outtxtClass = outtxtClass + '    Selection.Delete Shift:=xlUp\n';
        outtxtClass = outtxtClass + '    Range("A1").Select\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass +
            '    Workbooks.OpenText Filename:=m_strFullFilename, _\n';
        outtxtClass = outtxtClass + '        Origin:=XLOrigin_UTF, _\n';
        outtxtClass = outtxtClass + '            DataType:=xlDelimited, _\n';
        outtxtClass = outtxtClass + '                Semicolon:=True\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass +
            '    mstrOpenWorkbookName = ActiveWorkbook.Name\n';
        outtxtClass = outtxtClass + '    Cells.Select\n';
        outtxtClass = outtxtClass + '    Selection.Copy\n';
        outtxtClass = outtxtClass +
            '    Windows(mstrMainWorkbookName).Activate\n';
        outtxtClass = outtxtClass + '    Sheets(m_strSheetname).Select\n';
        outtxtClass = outtxtClass + '    Range("A1").Select\n';
        outtxtClass = outtxtClass + '    If PasteSpec Then\n';
        outtxtClass = outtxtClass +
            '        Selection.PasteSpecial Paste:=xlPasteValues, ' +
            'Operation:=xlNone, SkipBlanks _\n';
        outtxtClass = outtxtClass + '            :=False, Transpose:=False\n';
        outtxtClass = outtxtClass + '    Else\n';
        outtxtClass = outtxtClass + '        ActiveSheet.Paste\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass +
            '    Windows(mstrOpenWorkbookName).Activate\n';
        outtxtClass = outtxtClass + '    Application.DisplayAlerts = False\n';
        outtxtClass = outtxtClass + '    ActiveWindow.Close (False)\n';
        outtxtClass = outtxtClass + '    Application.DisplayAlerts = True\n';
        outtxtClass = outtxtClass +
            '    Windows(mstrMainWorkbookName).Activate\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_EXIT:\n';
        outtxtClass = outtxtClass + '    On Error GoTo 0\n';
        outtxtClass = outtxtClass + '    ' + "'" + 'Code here\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    Exit Sub\n';
        outtxtClass = outtxtClass + '    \n';
        outtxtClass = outtxtClass + '    PROC_ERR:\n';
        outtxtClass = outtxtClass +
            '    Debug.Print "Error in Sub ' + className + '.LoadText"\n';
        outtxtClass = outtxtClass + '    If Err.Number Then\n';
        outtxtClass = outtxtClass +
            '        ' + "'" + 'MsgBox Err.Description\n';
        outtxtClass = outtxtClass + '        Debug.Print Err.Description\n';
        outtxtClass = outtxtClass + '    End If\n';
        outtxtClass = outtxtClass + '    Resume PROC_EXIT\n';
        outtxtClass = outtxtClass + 'End Sub\n';

    }

    $('#textClass').val(outtxtClass);

}

function createMainModule() {
    var projectName = capitalizeFirstLetter($('#NameProj').val());
    var outtextMain = "";
    outtextMain = outtextMain + 'Option Explicit\n';
    outtextMain = outtextMain + 'Sub ' + projectName + '()\n';
    outtextMain = outtextMain + '    ' + "'" + 'Description\n';
    outtextMain = outtextMain + '    ' + "'" + 'Parameters:\n';
    outtextMain = outtextMain + '    ' + "'" + 'Created by: Laszlo Tamas\n';
    outtextMain = outtextMain + '\n';
    outtextMain = outtextMain + '\n';
    outtextMain = outtextMain + '    On Error GoTo PROC_ERR\n';
    outtextMain = outtextMain + '\n';
    outtextMain = outtextMain + '    ' + "'" + 'Code here\n';
    outtextMain = outtextMain + '\n';
    outtextMain = outtextMain + '    ' + "'" + '---------------\n';
    outtextMain = outtextMain + 'PROC_EXIT:\n';
    outtextMain = outtextMain + '    On Error GoTo 0\n';
    outtextMain = outtextMain + '    Exit Sub\n';
    outtextMain = outtextMain + 'PROC_ERR:\n';
    outtextMain = outtextMain +
        '    Debug.Print  "Error in Procedure ' + projectName + '"\n';
    outtextMain = outtextMain + '    If Err.Number Then\n';
    outtextMain = outtextMain + '        Debug.Print  Err.Description\n';
    outtextMain = outtextMain + '    End If\n';
    outtextMain = outtextMain + '    Resume PROC_EXIT\n';
    outtextMain = outtextMain + 'End Sub\n';
    outtextMain = outtextMain + 'Private Sub ' + projectName + 'Test\n';
    outtextMain = outtextMain +
        '    ' + "'" + 'Test procedure for ' + projectName + '\n';
    outtextMain = outtextMain + '    Dim dtmStartTime As Date\n';
    outtextMain = outtextMain + '\n';
    outtextMain = outtextMain + '\n';
    outtextMain = outtextMain + '\n';
    outtextMain = outtextMain + '    dtmStartTime = Now()\n';
    outtextMain = outtextMain + '    Call ' + projectName + '()\n';
    outtextMain = outtextMain + 'End Sub\n';

    if ($('#CompCheck1').prop('checked')) {
        outtextMain = outtextMain + '' + "'" + '----------------\n';
        outtextMain = outtextMain + '' + "'" + 'Columns and Rows\n';
        outtextMain = outtextMain + '' + "'" + '----------------\n';
        outtextMain = outtextMain +
            'Private Function Col_Letter(lngCol As Long) As String\n';
        outtextMain = outtextMain +
            '    ' + "'" + 'Get letter from column number\n';
        outtextMain = outtextMain + '    Dim vArr\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    ' + "'" + '  On Error Resume Next\n';
        outtextMain = outtextMain +
            '    vArr = Split(Cells(1, lngCol).Address(True, False), "$")\n';
        outtextMain = outtextMain + '    Col_Letter = vArr(0)\n';
        outtextMain = outtextMain + 'End Function\n';
        outtextMain = outtextMain +
            'Private Function Col_LetterHeader(sheetName As String, ' +
            'headText As String, Optional headRow = 1) As String\n';
        outtextMain = outtextMain + '    ' + "'" +
            'Get column letter from header text\n';
        outtextMain = outtextMain + '    Dim lngColNumber As Long\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain +
            '    lngColNumber = Col_NumberHeader(sheetName, ' +
            'headText, headRow)\n';
        outtextMain = outtextMain +
            '    Col_LetterHeader = Col_Letter(lngColNumber)\n';
        outtextMain = outtextMain + 'End Function\n';
        outtextMain = outtextMain +
            'Private Function Col_Number(colLetter) As Long\n';
        outtextMain = outtextMain + '    ' + "'" +
            'Get column number from column letter\n';
        outtextMain = outtextMain +
            '    Col_Number = Range(colLetter & "1").Column\n';
        outtextMain = outtextMain + 'End Function\n';
        outtextMain = outtextMain +
            'Private Function Col_NumberHeader(sheetName As String, ' +
            'headText As String, Optional headRow = 1) As Long\n';
        outtextMain = outtextMain + '    ' + "'" +
            'Get column number from header text\n';
        outtextMain = outtextMain + '    Dim i As Long\n';
        outtextMain = outtextMain + '    Dim strCellString As String\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    Col_NumberHeader = 0\n';
        outtextMain = outtextMain + '    For i = 1 To 400\n';
        outtextMain = outtextMain +
            '        strCellString = ' +
            'Trim(CStr(Sheets(sheetName).Cells(headRow, i)))\n';
        outtextMain = outtextMain +
            '        If strCellString = headText Then\n';
        outtextMain = outtextMain + '            Col_NumberHeader = i\n';
        outtextMain = outtextMain + '            Exit Function\n';
        outtextMain = outtextMain + '        End If\n';
        outtextMain = outtextMain + '    Next i\n';
        outtextMain = outtextMain + 'End Function\n';
        outtextMain = outtextMain + 'Private Sub ColLetterTests()\n';
        outtextMain = outtextMain + '    ' + "'" +
            'Test for Col_Letter, Col_LetterHeader, Col_Number ' +
            'and Col_NumberHeader\n';
        outtextMain = outtextMain + '    Debug.Print Col_Letter(12)\n';
        outtextMain = outtextMain +
            '    Debug.Print Col_LetterHeader("Hogyallunk", "Any.csop.")\n';
        outtextMain = outtextMain + '    Debug.Print Col_Number("H")\n';
        outtextMain = outtextMain +
            '    Debug.Print Col_NumberHeader("Hogyallunk", "Any.csop.")\n';
        outtextMain = outtextMain + 'End Sub\n';
        outtextMain = outtextMain +
            'Private Function GetLastRow(sheetName As String, ' +
            'checkColumn As Long, _\n';
        outtextMain = outtextMain +
            '    Optional firstrow = 2, Optional lastrow = 600000, _\n';
        outtextMain = outtextMain +
            '        Optional backwardCheck = True) As Long\n';
        outtextMain = outtextMain + '    ' + "'" + 'Adott fül utolsó sora\n';
        outtextMain = outtextMain + '    Dim i As Long\n';
        outtextMain = outtextMain + '    Dim curSheet As Worksheet\n';
        outtextMain = outtextMain + '    Dim strCell As String\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain +
            '    Set curSheet = ActiveWorkbook.ActiveSheet\n';
        outtextMain = outtextMain + '    Sheets(sheetName).Activate\n';
        outtextMain = outtextMain + '    GetLastRow = 0\n';
        outtextMain = outtextMain + '    If backwardCheck Then\n';
        outtextMain = outtextMain +
            '        For i = lastrow To firstrow Step -1\n';
        outtextMain = outtextMain +
            '            strCell = Trim(CStr(Cells(i, checkColumn)))\n';
        outtextMain = outtextMain + '            If strCell <> "" Then\n';
        outtextMain = outtextMain + '                GetLastRow = i\n';
        outtextMain = outtextMain + '                Exit For\n';
        outtextMain = outtextMain + '            End If\n';
        outtextMain = outtextMain + '        Next i\n';
        outtextMain = outtextMain + '    Else\n';
        outtextMain = outtextMain + '        For i = firstrow To lastrow\n';
        outtextMain = outtextMain +
            '            strCell = Trim(CStr(Cells(i, checkColumn)))\n';
        outtextMain = outtextMain + '            If strCell = "" Then\n';
        outtextMain = outtextMain + '                GetLastRow = i - 1\n';
        outtextMain = outtextMain + '                Exit For\n';
        outtextMain = outtextMain + '            End If\n';
        outtextMain = outtextMain + '        Next i\n';
        outtextMain = outtextMain + '    End If\n';
        outtextMain = outtextMain + '    curSheet.Activate\n';
        outtextMain = outtextMain + '    Set curSheet = Nothing\n';
        outtextMain = outtextMain +
            '    Debug.Print "LastRow of " & sheetName & ": " & ' +
            'GetLastRow & " ChkCol:" & checkColumn\n';
        outtextMain = outtextMain + 'End Function\n';
        outtextMain = outtextMain + '\n';
        outtextMain = outtextMain + '\n';

    }
    if ($('#CompCheck2').prop('checked')) {
        outtextMain = outtextMain + '' + "'" + '--------------\n';
        outtextMain = outtextMain + '' + "'" + 'Refresh ON OFF\n';
        outtextMain = outtextMain + '' + "'" + '--------------\n';
        outtextMain = outtextMain + 'Private Sub RefreshOFF()\n';
        outtextMain = outtextMain + '    ' + "'" + 'Screen update OFF\n';
        outtextMain = outtextMain + '    With Application\n';
        outtextMain = outtextMain + '        .ScreenUpdating = False\n';
        outtextMain = outtextMain + '        .EnableEvents = False\n';
        outtextMain = outtextMain + '        ' + "'" +
            '.Calculation = xlCalculationManual\n';
        outtextMain = outtextMain + '    End With\n';
        outtextMain = outtextMain + 'End Sub\n';
        outtextMain = outtextMain + 'Private Sub RefreshON()\n';
        outtextMain = outtextMain + '    ' + "'" + 'Screen update ON\n';
        outtextMain = outtextMain + '    With Application\n';
        outtextMain = outtextMain + '        .ScreenUpdating = True\n';
        outtextMain = outtextMain + '        .EnableEvents = True\n';
        outtextMain = outtextMain + '        ' + "'" +
            '.Calculation = xlCalculationAutomatic\n';
        outtextMain = outtextMain + '    End With\n';
        outtextMain = outtextMain + 'End Sub\n';
        outtextMain = outtextMain + '\n';
    }
    if ($('#CompCheck3').prop('checked')) {
        outtextMain = outtextMain + '' + "'" + '----------------------\n';
        outtextMain = outtextMain + '' + "'" + 'Change keyboard layout\n';
        outtextMain = outtextMain + '' + "'" + '----------------------\n';
        outtextMain = outtextMain + 'Private Sub SwitchToENG()\n';
        outtextMain = outtextMain + '    ' + "'" + 'Váltás angolra\n';
        outtextMain = outtextMain +
            '    Call ActivateKeyboardLayout(1033, 0)\n';
        outtextMain = outtextMain + 'End Sub\n';
        outtextMain = outtextMain + 'Private Sub SwitchToHUN()\n';
        outtextMain = outtextMain + '    ' + "'" + 'Váltás magyarra\n';
        outtextMain = outtextMain +
            '    Call ActivateKeyboardLayout(1038, 0)\n';
        outtextMain = outtextMain + 'End Sub\n';
        outtextMain = outtextMain + 'Private Sub SwitchToTUR()\n';
        outtextMain = outtextMain + '    ' + "'" + 'Váltás törökre\n';
        outtextMain = outtextMain +
            '    Call ActivateKeyboardLayout(1055, 0)\n';
        outtextMain = outtextMain + 'End Sub\n';
        outtextMain = outtextMain + '\n';
        outtextMain = outtextMain + '\n';

    }
    if ($('#CompCheck4').prop('checked')) {
        outtextMain = outtextMain + '' + "'" +
            '------------------------------------------\n';
        outtextMain = outtextMain + '' + "'" +
            'Create Outlook Appointment for ' + projectName + '\n';
        outtextMain = outtextMain + '' + "'" +
            '------------------------------------------\n';
        outtextMain = outtextMain + 'Public Sub CreateAppt' +
            projectName + '(sSubject, sBodyText, sDate)\n';
        outtextMain = outtextMain + '    ' + "'" +
            'A CreateObject módszerrel Office verzió független\n';
        outtextMain = outtextMain + '    Dim olApp As Object\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain +
            '    Set olApp = CreateObject("Outlook.Application")\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    ' + "'" +
            '    Dim olApp As Outlook.Application\n';
        outtextMain = outtextMain + '    Dim olAppt As Object\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    Set olAppt = olApp.CreateItem(1) ' +
            "'" + '0, mail, 1 appointment\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    Dim blnCreated As Boolean\n';
        outtextMain = outtextMain + '    Dim olNs As Object\n';
        outtextMain = outtextMain + '    Dim CalFolder As Object\n';
        outtextMain = outtextMain + '    Dim subFolder As Object\n';
        outtextMain = outtextMain + '    Dim arrCal As String\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    Dim i As Long\n';
        outtextMain = outtextMain + '    Dim strCalSubject As String\n';
        outtextMain = outtextMain + '    Dim strCalPlace As String\n';
        outtextMain = outtextMain + '    Dim strCalBody As String\n';
        outtextMain = outtextMain + '    Dim strStopString As String\n';
        outtextMain = outtextMain + '    Dim colNev\n';
        outtextMain = outtextMain + '    Dim colSznap\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    ' + "'" +
            '    Dim oPattern As RecurrencePattern\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    colNev = 1\n';
        outtextMain = outtextMain + '    colSznap = 2\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    strStopString = ""\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    On Error GoTo PROC_ERR\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    strCalPlace = vbNullString\n';
        outtextMain = outtextMain + '    strCalSubject = vbNullString\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    On Error Resume Next\n';
        outtextMain = outtextMain + '    ' + "'" +
            '    Set olApp = Outlook.Application\n';
        outtextMain = outtextMain + '    ' + "'" + '\n';
        outtextMain = outtextMain + '    ' + "'" +
            '    If olApp Is Nothing Then\n';
        outtextMain = outtextMain + '    ' + "'" +
            '        Set olApp = Outlook.Application\n';
        outtextMain = outtextMain + '    ' + "'" +
            '        blnCreated = True\n';
        outtextMain = outtextMain + '    ' + "'" + '        Err.Clear\n';
        outtextMain = outtextMain + '    ' + "'" + '    Else\n';
        outtextMain = outtextMain + '    ' + "'" +
            '        blnCreated = False\n';
        outtextMain = outtextMain + '    ' + "'" + '    End If\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    On Error GoTo 0\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain +
            '    Set olNs = olApp.GetNamespace("MAPI")\n';
        outtextMain = outtextMain +
            '    Set CalFolder = olNs.GetDefaultFolder(9) ' + "'" +
            '9 olFolderCalendar\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    ' + "'" +
            '    arrCal = cmstrCalSzuletesnapok ' + "'" + 'Naptár neve\n';
        outtextMain = outtextMain + '    ' + "'" + '\n';
        outtextMain = outtextMain + '    ' + "'" +
            '    Set subFolder = CalFolder.Folders(arrCal)\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain +
            '    Set olAppt = CalFolder.Items.Add(1) ' + "'" +
            '1 olAppointmentItem\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    ' + "'" +
            'MsgBox subFolder, vbOKCancel, "Folder Name"\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    With olAppt\n';
        outtextMain = outtextMain + '        \n';
        outtextMain = outtextMain + '        ' + "'" +
            'Define calendar item properties\n';
        outtextMain = outtextMain + '        \n';
        outtextMain = outtextMain + '        ' + "'" +
            'Debug.Print DateValue("2016." & ' +
            'Right(Trim(CStr(Cells(i, colSznap))), 4)) + ' +
            'TimeValue("0:00:00")\n';
        outtextMain = outtextMain +
            '        .Start = DateTime.DateSerial(CInt(Left(sDate, 4)), ' +
            'CInt(Mid(sDate, 6, 2)), CInt(Mid(sDate, 9, 2))) + ' +
            'TimeValue("9:00:00") ' + "'" + 'Given date 09:00\n';
        outtextMain = outtextMain + '        .Subject = sSubject\n';
        outtextMain = outtextMain + '        .Body = sBodyText\n';
        outtextMain = outtextMain + '        .ReminderSet = True\n';
        outtextMain = outtextMain +
            '        .ReminderMinutesBeforeStart = 4320 ' + "'" +
            '3 days, 72 hours\n';
        outtextMain = outtextMain + '        .MeetingStatus = 0 ' +
            "'" + '0 olNonMeeting\n';
        outtextMain = outtextMain + '        .Save\n';
        outtextMain = outtextMain + '    End With\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    Set olAppt = Nothing\n';
        outtextMain = outtextMain + '    Set olApp = Nothing\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + 'PROC_EXIT:\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + '    Exit Sub\n';
        outtextMain = outtextMain + '    \n';
        outtextMain = outtextMain + 'PROC_ERR:\n';
        outtextMain = outtextMain +
            '    MsgBox "An error occurred - Exporting items to Calendar." & ' +
            'vbCrLf & Err.Number & " " & Err.Description\n';
        outtextMain = outtextMain + '    Resume PROC_EXIT\n';
        outtextMain = outtextMain + 'End Sub\n';
        outtextMain = outtextMain + 'Private Sub CreateAppt' +
            projectName + '_Test()\n';
        outtextMain = outtextMain + '    Call CreateAppt' + projectName +
            '("Kórház2", "Teszt szöveg2", "2017.10.18")\n';
        outtextMain = outtextMain + 'End Sub\n';
        outtextMain = outtextMain + '\n';
        outtextMain = outtextMain + '\n';
        outtextMain = outtextMain + '\n';

    }

    $('#textMain').val(outtextMain);
}

function createMenu() {
    var projectName = capitalizeFirstLetter($('#NameProj').val());
    var ribbonName = projectName + "Ribbon";

    var outtxtThisWorkbook = "";
    outtxtThisWorkbook = outtxtThisWorkbook +
        'Private Sub Workbook_BeforeClose(Cancel As Boolean)\n';
    outtxtThisWorkbook = outtxtThisWorkbook +
        '    Call DeleteRibbons' + projectName + '\n';
    outtxtThisWorkbook = outtxtThisWorkbook + 'End Sub\n';
    outtxtThisWorkbook = outtxtThisWorkbook + '\n';
    outtxtThisWorkbook = outtxtThisWorkbook + 'Private Sub Workbook_Open()\n';
    outtxtThisWorkbook = outtxtThisWorkbook +
        '    Call AddRibbons' + projectName + '\n';
    outtxtThisWorkbook = outtxtThisWorkbook + 'End Sub\n';
    $('#textThisWb').val(outtxtThisWorkbook);

    var outtxtMenuModule = "";
    outtxtMenuModule = outtxtMenuModule + 'Option Explicit\n';
    outtxtMenuModule = outtxtMenuModule +
        'Global Const sToolbar' + projectName + ' As String = "' +
        ribbonName + '"\n';
    outtxtMenuModule = outtxtMenuModule + '\n';
    outtxtMenuModule = outtxtMenuModule +
        'Public Sub AddRibbons' + projectName + '()\n';
    outtxtMenuModule = outtxtMenuModule +
        '    ' + "'" + 'Add user ribbons, call it from Workbook_Open\n';
    outtxtMenuModule = outtxtMenuModule +
        '    Call AddRibbonLine' + projectName + '\n';
    outtxtMenuModule = outtxtMenuModule + 'End Sub\n';
    outtxtMenuModule = outtxtMenuModule +
        'Public Sub DeleteRibbons' + projectName + '()\n';
    outtxtMenuModule = outtxtMenuModule +
        '    ' + "'" + 'Delete ribbons, call it from Workbook_BeforeClose\n';
    outtxtMenuModule = outtxtMenuModule + '    On Error Resume Next\n';
    outtxtMenuModule = outtxtMenuModule +
        '    Application.CommandBars(sToolbar' + projectName + ').Delete\n';
    outtxtMenuModule = outtxtMenuModule + 'End Sub\n';
    outtxtMenuModule = outtxtMenuModule + '\n';

    outtxtMenuModule = outtxtMenuModule +
        'Sub AddRibbonLine' + projectName + '()\n';
    outtxtMenuModule = outtxtMenuModule + '    ' + "'" + ribbonName + '\n';
    outtxtMenuModule = outtxtMenuModule + '    Dim cbToolBar\n';
    outtxtMenuModule = outtxtMenuModule + '    \n';

    for (var i = 1; i < menuNum + 1; i++) {
        if ($('#MenuCheck' + i).prop('checked')) {
            outtxtMenuModule = outtxtMenuModule +
                '	Dim ctButton' + i.toString() + '\n';
        }
    }


    outtxtMenuModule = outtxtMenuModule + '    \n';
    outtxtMenuModule = outtxtMenuModule + '    On Error Resume Next\n';
    outtxtMenuModule = outtxtMenuModule +
        '    Set cbToolBar = Application.CommandBars.Add(sToolbar' +
        projectName + ', msoBarTop, False, True)\n';
    outtxtMenuModule = outtxtMenuModule + '    With cbToolBar\n';

    for (i = 1; i < menuNum + 1; i++) {
        if ($('#MenuCheck' + i).prop('checked')) {
            outtxtMenuModule = outtxtMenuModule +
                '        Set ctButton' + i.toString() +
                ' = .Controls.Add(Type:=msoControlButton, ID:=2950)\n';
        }
    }


    outtxtMenuModule = outtxtMenuModule + '    End With\n';
    outtxtMenuModule = outtxtMenuModule + '    \n';

    for (i = 1; i < menuNum + 1; i++) {
        if ($('#MenuCheck' + i).prop('checked')) {
            outtxtMenuModule = outtxtMenuModule +
                '    With ctButton' + i.toString() + '\n';
            outtxtMenuModule = outtxtMenuModule +
                '        .Caption = "' + $('#Cap' + i).val() + '\n';
            outtxtMenuModule = outtxtMenuModule +
                '        .FaceId = ' + $('#Face' + i).val() + '\n';
            outtxtMenuModule = outtxtMenuModule +
                '        .OnAction = "' + $('#Onaction' + i).val() + '"\n';
            outtxtMenuModule = outtxtMenuModule +
                '        .TooltipText = "' + $('#TTip' + i).val() + '"\n';
            outtxtMenuModule = outtxtMenuModule +
                '        .Style = msoButtonIconAndCaption\n';
            outtxtMenuModule = outtxtMenuModule + '    End With\n';
            outtxtMenuModule = outtxtMenuModule + '    \n';
        }
    }

    outtxtMenuModule = outtxtMenuModule + '    \n';
    outtxtMenuModule = outtxtMenuModule + '    \n';
    outtxtMenuModule = outtxtMenuModule + '    \n';
    outtxtMenuModule = outtxtMenuModule + '    With cbToolBar\n';
    outtxtMenuModule = outtxtMenuModule + '        .Visible = True\n';
    outtxtMenuModule = outtxtMenuModule +
        '        .Protection = msoBarNoChangeVisible\n';
    outtxtMenuModule = outtxtMenuModule + '    End With\n';
    outtxtMenuModule = outtxtMenuModule + 'End Sub\n';


    outtxtMenuModule = outtxtMenuModule + 'Sub MenuNULL()\n';
    outtxtMenuModule = outtxtMenuModule +
        '    ' + "'" + 'Empty dummy subroutine\n';
    outtxtMenuModule = outtxtMenuModule + 'End Sub\n';



    $('#textMenu').val(outtxtMenuModule);

}

function setPredef() {
    var predef = $('#ProjPredef').val();
    //TODO jQuery and formatting 
    if (predef == "SAP") {
        $('#CompCheck1').prop('checked', true);
        $('#ClassCompCheck1').prop('checked', true);
        $('#ClassCompCheck2').prop('checked', true);
        $('#ClassCompCheck3').prop('checked', true);

        $('#ClassCheck1').prop('checked', true);
        $('#Clprop1').val('FullFilename');
        $('#Clpar1').val('parFullFilename');
        $('#Type1').val('String');
        $('#Mode1').val('Let and Get');


        $('#ClassCheck2').prop('checked', true);
        $('#Clprop2').val('Filename');
        $('#Clpar2').val('parFilename');
        $('#Type2').val('String');
        $('#Mode2').val('Let and Get');


        $('#ClassCheck3').prop('checked', true);
        $('#Clprop3').val('Path');
        $('#Clpar3').val('parPath');
        $('#Type3').val('String');
        $('#Mode3').val('Let and Get');

        $('#ClassCheck4').prop('checked', true);
        $('#Clprop4').val('Sheetname');
        $('#Clpar4').val('parSheetname');
        $('#Type4').val('String');
        $('#Mode4').val('Let and Get');

        $('#ClassCheck5').prop('checked', true);
        $('#Clprop5').val('OutputFilename');
        $('#Clpar5').val('parOutputFilename');
        $('#Type5').val('String');
        $('#Mode5').val('Let and Get');

    }

}

function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}

function getPrefix(varType) {

    if (varType == "String") {
        return "str";
    }
    if (varType == "Long") {
        return "lng";
    }
    if (varType == "Integer") {
        return "int";
    }
    if (varType == "Boolean") {
        return "bln";
    }
    if (varType == "Double") {
        return "dbl";
    }
    if (varType == "Date") {
        return "dat";
    }
    if (varType == "Variant") {
        return "vnt";
    }
    if (varType == "Object") {
        return "obj";
    }
    if (varType == "SheetName") {
        return "sh";
    }
    if (varType == "Worksheet") {
        return "wst";
    }
    if (varType == "Outlook") {
        return "ol";
    }

}

function getConstInitValue(varType) {
    if (varType == "String") {
        return ' = "text"';
    }
    if (varType == "Long") {
        return " = 0";
    }
    if (varType == "Integer") {
        return " = 0";
    }
    if (varType == "Boolean") {
        return " = True";
    }
    if (varType == "Double") {
        return " = 0.5";
    }
    if (varType == "Date") {
        return ' = CDate("04/22/2016 12:00 AM")';
    }
    if (varType == "Variant") {
        return " = True";
    }
    if (varType == "Object") {
        return " = ";
    }
    if (varType == "SheetName") {
        return ' = "Munka1"';
    }
    if (varType == "Worksheet") {
        return " = ActiveSheet";
    }
    if (varType == "Outlook") {
        return " = ";
    }

}

function getDeclareType(varType) {
    if (varType == "String") {
        return "String";
    }
    if (varType == "Long") {
        return "Long";
    }
    if (varType == "Integer") {
        return "Integer";
    }
    if (varType == "Boolean") {
        return "Boolean";
    }
    if (varType == "Double") {
        return "Double";
    }
    if (varType == "Date") {
        return "Date";
    }
    if (varType == "Variant") {
        return "Variant";
    }
    if (varType == "Object") {
        return "Object";
    }
    if (varType == "SheetName") {
        return "String";
    }
    if (varType == "Worksheet") {
        return "Worksheet";
    }
    if (varType == "Outlook") {
        return "Outlook";
    }

}