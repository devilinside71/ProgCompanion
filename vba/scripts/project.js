var menuNum = 8;
// Min 5 properties
var propNum = 8;

$(document).ready(function() {
  var i = 1;
  var typeOptions = new Array(
    'String',
    'Long',
    'Integer',
    'Boolean',
    'Double',
    'Date',
    'Variant',
    'Object',
    'SheetName',
    'Worksheet',
    'Outlook'
  );
  for (i = 1; i < menuNum + 1; i++) {
    $('#tableMenu')
      .find('tbody')
      .append(
        $('<tr>')
          .append($('<td class="Col1">Menu' + i + '</td>'))
          .append(
            $(
              '<td class="checkBoxCol">' +
                '<input type="checkbox" id="MenuCheck' +
                i +
                '" /></td>'
            )
          )
          .append(
            $(
              '<td class="Col2"><input id="Cap' +
                i +
                '" type="text" value="Menu' +
                i +
                '" class="textBox" /></td>'
            )
          )
          .append(
            $(
              '<td class="Col3"><input id="Face' +
                i +
                '" type="text" value="1244" class="textBox" /></td>'
            )
          )
          .append(
            $(
              '<td class="Col4"><input id="Onaction' +
                i +
                '" type="text" value="MenuNULL" class="textBox" /></td>'
            )
          )
          .append(
            $(
              '<td class="Col5"><input id="TTip' +
                i +
                '" type="text" value="Menu' +
                i +
                ' végrehajtása" class="textBox" /></td>'
            )
          )
      );
  }
  for (i = 1; i < propNum + 1; i++) {
    $('#tableProps')
      .find('tbody')
      .append(
        $('<tr>')
          .append($('<td class="Col1">Property' + i + '</td>'))
          .append(
            $(
              '<td class="checkBoxCol">' +
                '<input type="checkbox" id="ClassCheck' +
                i +
                '" /></td>'
            )
          )
          .append(
            $(
              '<td class="Col2"><input id="Clprop' +
                i +
                '" type="text" value="Prop' +
                i +
                '" class="textBox" /></td>'
            )
          )
          .append(
            $(
              '<td class="Col3"><input id="Clpar' +
                i +
                '" type="text" value="Par' +
                i +
                '" class="textBox" /></td>'
            )
          )
          .append($('<td class="Col4"><select id="Type' + i + '">'))
          .append($('<td class="Col5"><select id="Mode' + i + '">'))
      );
  }
  for (i = 1; i < propNum + 1; i++) {
    $('#Type' + i)
      .append(
        $(
          '<option value="' +
            typeOptions[0] +
            '">' +
            typeOptions[0] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[1] +
            '">' +
            typeOptions[1] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[2] +
            '">' +
            typeOptions[2] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[3] +
            '">' +
            typeOptions[3] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[4] +
            '">' +
            typeOptions[4] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[5] +
            '">' +
            typeOptions[5] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[6] +
            '">' +
            typeOptions[6] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[7] +
            '">' +
            typeOptions[7] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[8] +
            '">' +
            typeOptions[8] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[9] +
            '">' +
            typeOptions[9] +
            '</option>'
        )
      )
      .append(
        $(
          '<option value="' +
            typeOptions[10] +
            '">' +
            typeOptions[10] +
            '</option>'
        )
      );
  }
  for (i = 1; i < propNum + 1; i++) {
    $('#Mode' + i)
      .append($('<option value="Let and Get">Let and Get</option>'))
      .append($('<option value="Let">Let</option>'))
      .append($('<option value="Get">Get</option>'));
  }

  $('#createproject').click(function() {
    createProject();
  });
  $('#clear').click(function() {
    resetNames();
  });

  $('#createclass').click(function() {
    createClass();
  });
  $('#setpredef').click(function() {
    setPredef();
  });
  $('#NameProj').on('input', function() {
    $('#NameClass').val($('#NameProj').val() + 'Class');
  });
});

function createProject() {
  createMenu();
  createMainModule();
  createClass();
}

function resetNames() {
  for (i = 1; i < menuNum + 1; i++) {
    $('#Cap' + i).val('Button' + i);
    $('#Face' + i).val('1244');
    $('#Onaction' + i).val('MenuNULL');
    $('#TTip' + i).val('Button' + i + ' végrahajtása');
  }
}

function createClass() {
  var className = capitalizeFirstLetter($('#NameClass').val());
  // Classtest
  var outtextClasstest = '';
  var i = 1;
  var outtxtClass;
  outtextClasstest =
    outtextClasstest + 'Private Sub ' + className + '_ClassTest()\n';
  outtextClasstest =
    outtextClasstest + '    Dim cl' + className + ' As New ' + className + '\n';
  outtextClasstest += '    \n';
  for (i = 1; i < propNum + 1; i++) {
    if ($('#ClassCheck' + 1).prop('checked')) {
      outtextClasstest =
        outtextClasstest +
        '    cl' +
        className +
        '.' +
        $('#Clprop' + i).val() +
        getConstInitValue($('#Type' + i).val()) +
        '\n';
      outtextClasstest =
        outtextClasstest +
        '    Debug.Print "cl' +
        className +
        '.' +
        $('#Clprop' + i).val() +
        ': " & cl' +
        className +
        '.' +
        $('#Clprop' + i).val() +
        '\n';
    }
  }

  outtextClasstest =
    outtextClasstest + '    Set cl' + className + ' = Nothing\n';
  outtextClasstest += 'End Sub\n';
  $('#textMain').val($('#textMain').val() + outtextClasstest);

  // Definition
  outtxtClass = '';
  outtxtClass += 'Option Explicit\n';
  outtxtClass += '\n';

  for (i = 1; i < propNum + 1; i++) {
    if ($('#ClassCheck' + 1).prop('checked')) {
      outtxtClass =
        outtxtClass +
        'Private m_' +
        getPrefix($('#Type' + i).val()) +
        $('#Clprop' + i).val() +
        ' As ' +
        getDeclareType($('#Type' + i).val()) +
        '\n';
      outtxtClass =
        outtxtClass +
        'Private Const cm' +
        getPrefix($('#Type' + i).val()) +
        $('#Clprop' + i).val() +
        ' As ' +
        getDeclareType($('#Type' + i).val()) +
        getConstInitValue($('#Type' + i).val()) +
        '\n';
    }
  }

  // Properties
  for (i = 1; i < propNum + 1; i++) {
    if ($('#ClassCheck' + i).prop('checked')) {
      if (
        $('#Mode' + i).val() === 'Let and Get' ||
        $('#Mode' + i).val() === 'Let'
      ) {
        outtxtClass =
          outtxtClass +
          'Public Property Let ' +
          $('#Clprop' + i).val() +
          '(' +
          $('#Clpar' + i).val() +
          ' As ' +
          getDeclareType($('#Type' + i).val()) +
          ')\n';
        outtxtClass += '    \n';
        outtxtClass += '    On Error GoTo PROC_ERR\n';
        outtxtClass += '    \n';
        outtxtClass =
          outtxtClass +
          '    m_' +
          getPrefix($('#Type' + i).val()) +
          $('#Clprop' + i).val() +
          ' = ' +
          $('#Clpar' + i).val() +
          '\n';
        outtxtClass =
          outtxtClass +
          '    Debug.Print "' +
          className +
          '.' +
          $('#Clprop' + i).val() +
          ' has been set to: " & m_' +
          getPrefix($('#Type' + i).val()) +
          $('#Clprop' + i).val() +
          '\n';
        outtxtClass += '    \n';
        outtxtClass += 'PROC_EXIT:\n';
        outtxtClass += '    Exit Property\n';
        outtxtClass += '    \n';
        outtxtClass += 'PROC_ERR:\n';
        outtxtClass += '    Err.Raise Err.Number\n';
        outtxtClass += '    Resume PROC_EXIT\n';
        outtxtClass += 'End Property\n';
      }
      if (
        $('#Mode' + i).val() === 'Let and Get' ||
        $('#Mode' + i).val() === 'Get'
      ) {
        outtxtClass += '\n';
        outtxtClass =
          outtxtClass +
          'Public Property Get ' +
          $('#Clprop' + i).val() +
          '() As ' +
          getDeclareType($('#Type' + i).val()) +
          '\n';
        outtxtClass += '    \n';
        outtxtClass += '    On Error GoTo PROC_ERR\n';
        outtxtClass += '    \n';
        outtxtClass =
          outtxtClass +
          '    ' +
          $('#Clprop' + i).val() +
          ' = m_' +
          getPrefix($('#Type' + i).val()) +
          $('#Clprop' + i).val() +
          '\n';
        outtxtClass += '    \n';
        outtxtClass += 'PROC_EXIT:\n';
        outtxtClass += '    Exit Property\n';
        outtxtClass += '    \n';
        outtxtClass += 'PROC_ERR:\n';
        outtxtClass += '    Err.Raise Err.Number\n';
        outtxtClass += '    Resume PROC_EXIT\n';
        outtxtClass += 'End Property\n';
      }
    }
  }

  // Class body
  outtxtClass += 'Private Sub Class_Initialize()\n';
  outtxtClass =
    outtxtClass + '    Debug.Print "Class ' + className + ' initialized"\n';
  outtxtClass += '    \n';
  for (i = 1; i < propNum + 1; i++) {
    if ($('#ClassCheck' + i).prop('checked')) {
      outtxtClass =
        outtxtClass +
        '    m_' +
        getPrefix($('#Type' + i).val()) +
        $('#Clprop' + i).val() +
        ' = cm' +
        getPrefix($('#Type' + i).val()) +
        $('#Clprop' + i).val() +
        '\n';
      outtxtClass =
        outtxtClass +
        '    Debug.Print "' +
        className +
        ' Default value for ' +
        $('#Clprop' + i).val() +
        ': " & m_' +
        getPrefix($('#Type' + i).val()) +
        $('#Clprop' + i).val() +
        '\n';
    }
  }

  outtxtClass += 'End Sub\n';
  outtxtClass += 'Private Sub Class_Terminate()\n';
  outtxtClass =
    outtxtClass + '    Debug.Print "Class ' + className + ' terminated"\n';
  outtxtClass += 'End Sub\n';
  outtxtClass += 'Sub Reset()\n';
  outtxtClass += '    \n';

  for (i = 1; i < propNum + 1; i++) {
    if ($('#ClassCheck' + i).prop('checked')) {
      outtxtClass =
        outtxtClass +
        '    m_' +
        getPrefix($('#Type' + i).val()) +
        $('#Clprop' + i).val() +
        ' = cm' +
        getPrefix($('#Type' + i).val()) +
        $('#Clprop' + i).val() +
        '\n';
      outtxtClass =
        outtxtClass +
        '    Debug.Print "' +
        className +
        ' Default value for ' +
        $('#Clprop' + i).val() +
        ': " & m_' +
        getPrefix($('#Type' + i).val()) +
        $('#Clprop' + i).val() +
        '\n';
    }
  }

  outtxtClass += 'End Sub\n';

  // Others
  if ($('#ClassCompCheck1').prop('checked')) {
    outtxtClass += '\'----------------\n';
    outtxtClass += '\'Columns and Rows\n';
    outtxtClass += '\'----------------\n';
    outtxtClass += 'Private Function Col_Letter(lngCol As Long) As String\n';
    outtxtClass += '    \'Get letter from column number\n';
    outtxtClass += '    Dim vArr\n';
    outtxtClass += '    \n';
    outtxtClass += '    \'  On Error Resume Next\n';
    outtxtClass +=
      '    vArr = Split(Cells(1, lngCol).Address(True, False), "$")\n';
    outtxtClass += '    Col_Letter = vArr(0)\n';
    outtxtClass += 'End Function\n';
    outtxtClass =
      outtxtClass +
      'Private Function Col_LetterHeader(sheetName As String,' +
      ' headText As String, Optional headRow = 1) As String\n';
    outtxtClass += '    \'Get column letter from header text\n';
    outtxtClass += '    Dim lngColNumber As Long\n';
    outtxtClass += '    \n';
    outtxtClass =
      outtxtClass +
      '    lngColNumber = ' +
      'Col_NumberHeader(sheetName, headText, headRow)\n';
    outtxtClass += '    Col_LetterHeader = Col_Letter(lngColNumber)\n';
    outtxtClass += 'End Function\n';
    outtxtClass += 'Private Function Col_Number(colLetter) As Long\n';
    outtxtClass += '    \'Get column number from column letter\n';
    outtxtClass += '    Col_Number = Range(colLetter & "1").Column\n';
    outtxtClass += 'End Function\n';
    outtxtClass =
      outtxtClass +
      'Private Function Col_NumberHeader(sheetName As String, ' +
      'headText As String, Optional headRow = 1) As Long\n';
    outtxtClass += '    \'Get column number from header text\n';
    outtxtClass += '    Dim i As Long\n';
    outtxtClass += '    Dim strCellString As String\n';
    outtxtClass += '    \n';
    outtxtClass += '    Col_NumberHeader = 0\n';
    outtxtClass += '    For i = 1 To 400\n';
    outtxtClass =
      outtxtClass +
      '        strCellString = ' +
      'Trim(CStr(Sheets(sheetName).Cells(headRow, i)))\n';
    outtxtClass += '        If strCellString = headText Then\n';
    outtxtClass += '            Col_NumberHeader = i\n';
    outtxtClass += '            Exit Function\n';
    outtxtClass += '        End If\n';
    outtxtClass += '    Next i\n';
    outtxtClass += 'End Function\n';
    outtxtClass += 'Private Sub ColLetterTests()\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      'Test for Col_Letter, Col_LetterHeader, ' +
      'Col_Number and Col_NumberHeader\n';
    outtxtClass += '    Debug.Print Col_Letter(12)\n';
    outtxtClass +=
      '    Debug.Print Col_LetterHeader("Hogyallunk", "Any.csop.")\n';
    outtxtClass += '    Debug.Print Col_Number("H")\n';
    outtxtClass +=
      '    Debug.Print Col_NumberHeader("Hogyallunk", "Any.csop.")\n';
    outtxtClass += 'End Sub\n';
    outtxtClass =
      outtxtClass +
      'Private Function GetLastRow(sheetName As String, ' +
      'checkColumn As Long, _\n';
    outtxtClass += '    Optional firstrow = 2, Optional lastrow = 600000, _\n';
    outtxtClass += '        Optional backwardCheck = True) As Long\n';
    outtxtClass += '    \'Adott fül utolsó sora\n';
    outtxtClass += '    Dim i As Long\n';
    outtxtClass += '    Dim curSheet As Worksheet\n';
    outtxtClass += '    Dim strCell As String\n';
    outtxtClass += '    \n';
    outtxtClass += '    Set curSheet = ActiveWorkbook.ActiveSheet\n';
    outtxtClass += '    Sheets(sheetName).Activate\n';
    outtxtClass += '    GetLastRow = 0\n';
    outtxtClass += '    If backwardCheck Then\n';
    outtxtClass += '        For i = lastrow To firstrow Step -1\n';
    outtxtClass += '            strCell = Trim(CStr(Cells(i, checkColumn)))\n';
    outtxtClass += '            If strCell <> "" Then\n';
    outtxtClass += '                GetLastRow = i\n';
    outtxtClass += '                Exit For\n';
    outtxtClass += '            End If\n';
    outtxtClass += '        Next i\n';
    outtxtClass += '    Else\n';
    outtxtClass += '        For i = firstrow To lastrow\n';
    outtxtClass += '            strCell = Trim(CStr(Cells(i, checkColumn)))\n';
    outtxtClass += '            If strCell = "" Then\n';
    outtxtClass += '                GetLastRow = i - 1\n';
    outtxtClass += '                Exit For\n';
    outtxtClass += '            End If\n';
    outtxtClass += '        Next i\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    curSheet.Activate\n';
    outtxtClass += '    Set curSheet = Nothing\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "LastRow of " & sheetName & ": " &' +
      ' GetLastRow & " ChkCol:" & checkColumn\n';
    outtxtClass += 'End Function\n';
    outtxtClass += '\n';
    outtxtClass += '\n';
  }
  if ($('#ClassCompCheck2').prop('checked')) {
    outtxtClass += '\'--------------\n';
    outtxtClass += '\'Refresh ON OFF\n';
    outtxtClass += '\'--------------\n';
    outtxtClass += 'Private Sub RefreshOFF()\n';
    outtxtClass += '    \'Screen update OFF\n';
    outtxtClass += '    With Application\n';
    outtxtClass += '        .ScreenUpdating = False\n';
    outtxtClass += '        .EnableEvents = False\n';
    outtxtClass += '        \'.Calculation = xlCalculationManual\n';
    outtxtClass += '    End With\n';
    outtxtClass += 'End Sub\n';
    outtxtClass += 'Private Sub RefreshON()\n';
    outtxtClass += '    \'Screen update ON\n';
    outtxtClass += '    With Application\n';
    outtxtClass += '        .ScreenUpdating = True\n';
    outtxtClass += '        .EnableEvents = True\n';
    outtxtClass =
      outtxtClass +
      '        ' +
      '\'' +
      '.Calculation = xlCalculationAutomatic\n';
    outtxtClass += '    End With\n';
    outtxtClass += 'End Sub\n';
    outtxtClass += '\n';
  }

  // SAP body
  if ($('#ClassCompCheck3').prop('checked')) {
    outtxtClass =
      outtxtClass +
      'Private Function GetPathOfFile(FullFilename ' +
      'As String) As String\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      'Comments  : Get path of a full filename with path\n';
    outtxtClass += '    \'Parameters:\n';
    outtxtClass += '    \'\'           FullFilename\n';
    outtxtClass += '    \n';
    outtxtClass += '    Dim strRes As String\n';
    outtxtClass += '    Dim fso\n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo FUNC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    Set fso = CreateObject("Scripting.FileSystemObject")\n';
    outtxtClass += '    strRes = fso.GetParentFolderName(FullFilename)\n';
    outtxtClass += '    GetPathOfFile = strRes & ""\n';
    outtxtClass += '    \n';
    outtxtClass += '    \'-------------------------------\n';
    outtxtClass += '    FUNC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    Set fso = Nothing\n';
    outtxtClass += '    Exit Function\n';
    outtxtClass += '    \n';
    outtxtClass += '    FUNC_ERR:\n';
    outtxtClass +=
      '    Debug.Print "Error in Function %PROPNAME%.GetPathOfFile"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume FUNC_EXIT\n';
    outtxtClass += '    \n';
    outtxtClass += 'End Function\n';
    outtxtClass =
      outtxtClass +
      'Private Function GetFilenameOfFile(FullFilename ' +
      'As String) As String\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      'Comments  : Get filename of a full filename with path\n';
    outtxtClass += '    \'Parameters:\n';
    outtxtClass += '    \'\'           FullFilename\n';
    outtxtClass += '    \n';
    outtxtClass += '    Dim strRes As String\n';
    outtxtClass += '    Dim fso\n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo FUNC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    Set fso = CreateObject("Scripting.FileSystemObject")\n';
    outtxtClass += '    strRes = fso.GetFileName(FullFilename)\n';
    outtxtClass += '    GetFilenameOfFile = strRes\n';
    outtxtClass += '    \n';
    outtxtClass += '    \'-------------------------------\n';
    outtxtClass += '    FUNC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    Set fso = Nothing\n';
    outtxtClass += '    Exit Function\n';
    outtxtClass += '    \n';
    outtxtClass += '    FUNC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Function ' +
      className +
      '.GetFilenameOfFile"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume FUNC_EXIT\n';
    outtxtClass += '    \n';
    outtxtClass += 'End Function\n';
    outtxtClass += '\n';
    outtxtClass += 'Sub AddSheet()\n';
    outtxtClass += '    \'Comments  : Add zhogyallunk sheet\n';
    outtxtClass += '    \'Parameters:\n';
    outtxtClass += '    \'\'\n';
    outtxtClass += '    \n';
    outtxtClass += '    \'Created by: Laszlo Tamas\n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    Sheets.Add Before:=Sheets(1)\n';
    outtxtClass += '    ActiveSheet.Name = m_strSheetname\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    Exit Sub\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Sub ' +
      className +
      '.AddSheet"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += 'End Sub\n';
    outtxtClass += '\n';
    outtxtClass += 'Sub DeleteSheet(Optional Alert As Boolean = False)\n';
    outtxtClass += '    \'Comments  : Remarks\n';
    outtxtClass += '    \'Parameters:\n';
    outtxtClass += '    \'\'           Alert\n';
    outtxtClass += '    \n';
    outtxtClass += '    \'Created by: Laszlo Tamas\n';
    outtxtClass += '    \n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    If Alert = False Then\n';
    outtxtClass += '        Application.DisplayAlerts = False\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    On Error Resume Next\n';
    outtxtClass += '    Sheets(m_strSheetname).Delete\n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    Application.DisplayAlerts = True\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    Exit Sub\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Sub ' +
      className +
      '.DeleteSheet"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += 'End Sub\n';
    outtxtClass += 'Sub CreateSheet(Optional Alert As Boolean = False)\n';
    outtxtClass += '    \'Comments  : Remarks\n';
    outtxtClass += '    \'Parameters:\n';
    outtxtClass += '    \'\'           Alert\n';
    outtxtClass += '    \n';
    outtxtClass += '    \'Created by: Laszlo Tamas\n';
    outtxtClass += '    \n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    Call DeleteSheet(Alert)\n';
    outtxtClass += '    Call AddSheet\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    Exit Sub\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Sub ' +
      className +
      '.CreateSheet"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += 'End Sub\n';
    outtxtClass =
      outtxtClass +
      'Sub DuplicateSheet(NewSheetname As String, ' +
      'Optional Alert As Boolean = False)\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      'Comments  : Duplicate full hogyallunk sheet\n';
    outtxtClass += '    \'Parameters:\n';
    outtxtClass += '    \'\'           Alert\n';
    outtxtClass += '    \n';
    outtxtClass += '    \'Created by: Laszlo Tamas\n';
    outtxtClass += '    \n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    If Alert = False Then\n';
    outtxtClass += '        Application.DisplayAlerts = False\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    On Error Resume Next\n';
    outtxtClass += '    Sheets(NewSheetname).Delete\n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    Application.DisplayAlerts = True\n';
    outtxtClass += '    Sheets(m_strSheetname).Select\n';
    outtxtClass +=
      '    Sheets(m_strSheetname).Copy Before:=Sheets(m_strSheetname)\n';
    outtxtClass += '    ActiveSheet.Name = NewSheetname\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    Exit Sub\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Sub ' +
      className +
      '.DuplicateSheet"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += 'End Sub\n';
    outtxtClass =
      outtxtClass +
      'Sub CopySheetContent(NewSheetname As String, ' +
      'PasteSpec As Boolean, Optional Alert As Boolean = False)\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      'Comments  : Copy the (filtered, viewable) ' +
      'content of hogyallunk sheet\n';
    outtxtClass += '    \'Parameters:\n';
    outtxtClass += '    \'\'           Alert\n';
    outtxtClass += '    \n';
    outtxtClass += '    \'Created by: Laszlo Tamas\n';
    outtxtClass += '    \n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    If Alert = False Then\n';
    outtxtClass += '        Application.DisplayAlerts = False\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    On Error Resume Next\n';
    outtxtClass += '    Sheets(NewSheetname).Delete\n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    Application.DisplayAlerts = True\n';
    outtxtClass += '    \n';
    outtxtClass += '    Sheets.Add Before:=Sheets(m_strSheetname)\n';
    outtxtClass += '    ActiveSheet.Name = NewSheetname\n';
    outtxtClass += '    \n';
    outtxtClass += '    Sheets(m_strSheetname).Select\n';
    outtxtClass += '    Cells.Select\n';
    outtxtClass += '    Selection.Copy\n';
    outtxtClass += '    Sheets(NewSheetname).Select\n';
    outtxtClass += '    Range("A1").Select\n';
    outtxtClass += '    If PasteSpec Then\n';
    outtxtClass =
      outtxtClass +
      '        Selection.PasteSpecial Paste:=xlPasteValues, ' +
      'Operation:=xlNone, SkipBlanks _\n';
    outtxtClass += '            :=False, Transpose:=False\n';
    outtxtClass += '    Else\n';
    outtxtClass += '        ActiveSheet.Paste\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    Exit Sub\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Sub ' +
      className +
      '.CopySheetContent"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += 'End Sub\n';
    outtxtClass =
      outtxtClass +
      'Function GetOpenFilename(Optional Title = "Fájl", ' +
      'Optional CollectionName = ' +
      '"Fájlok", Optional Extensions = "*.*") ' +
      'As String\n';
    outtxtClass += '    \' Comments:\n';
    outtxtClass += '    \' Params  : Title\n';
    outtxtClass += '    \'           CollectionName\n';
    outtxtClass += '    \'           Extensions\n';
    outtxtClass += '    \' Returns : String\n';
    outtxtClass += '    \' Modified:\n';
    outtxtClass += '    \n';
    outtxtClass += '    \n';
    outtxtClass += '    Dim intChoice As Integer\n';
    outtxtClass += '    Dim strPath As String\n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    \n';
    outtxtClass += '    intChoice = 0\n';
    outtxtClass += '    strPath = vbNullString\n';
    outtxtClass += '    With Application.FileDialog(msoFileDialogOpen)\n';
    outtxtClass += '        .Title = Title & " kiválasztása"\n';
    outtxtClass += '        .Filters.Add CollectionName, Extensions\n';
    outtxtClass += '        .FilterIndex = .Filters.Count\n';
    outtxtClass += '        .AllowMultiSelect = False\n';
    outtxtClass += '        \n';
    outtxtClass += '        intChoice = .Show\n';
    outtxtClass += '        If intChoice <> 0 Then\n';
    outtxtClass += '            strPath = .SelectedItems(1)\n';
    outtxtClass += '        End If\n';
    outtxtClass += '    End With\n';
    outtxtClass += '    GetOpenFilename = strPath\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      '    Debug.Print "' +
      className +
      '.GetOpenFilename: " & GetOpenFilename & " Title:" & Title\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    Exit Function\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print Err.Description, vbCritical, "' +
      className +
      '.GetOpenFilename"\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += '    \n';
    outtxtClass += 'End Function\n';
    outtxtClass =
      outtxtClass +
      'Sub SetFilenameWithDialog(Optional Title = "Fájl", ' +
      'Optional CollectionName = "Fájlok", ' +
      'Optional Extensions = "*.*")\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      'Comments  : Load data from SAP exported CSV (text)\n';
    outtxtClass += '    \' Params  : Title\n';
    outtxtClass += '    \'           CollectionName\n';
    outtxtClass += '    \'           Extensions\n';
    outtxtClass += '    \'Created by: Laszlo Tamas\n';
    outtxtClass += '    \n';
    outtxtClass += '    Dim strRes As String\n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass +=
      '    strRes = GetOpenFilename(Title, CollectionName, Extensions)\n';
    outtxtClass += '    m_strFullFilename = strRes\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "' +
      className +
      '.SetFilenameWithDialog FullFilename ' +
      'has been set to: " & m_strFullFilename\n';
    outtxtClass += '    m_strPath = GetPathOfFile(strRes)\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "' +
      className +
      '.SetFilenameWithDialog Path ' +
      'has been set to: " & m_strPath\n';
    outtxtClass += '    m_strFilename = GetFilenameOfFile(strRes)\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "' +
      className +
      '.SetFilenameWithDialog Filename ' +
      'has been set to: " & m_strFilename\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    Exit Sub\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Sub ' +
      className +
      '.SetFilenameWithDialog"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += 'End Sub\n';
    outtxtClass =
      outtxtClass +
      'Function GetSaveFilename(Optional DTitle = "Fájl", ' +
      'Optional FFilter = "Excel files , *.xlsx") As String\n';
    outtxtClass += '    \' Comments:\n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    GetSaveFilename = vbNullString\n';
    outtxtClass =
      outtxtClass +
      '    GetSaveFilename = Application.GetSaveAsFilename(' +
      'InitialFileName:=m_strOutputFilename, ' +
      'FileFilter:=FFilter, Title:=DTitle)\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      '    Debug.Print "GetOpenFilename: " & ' +
      'GetSaveFilename & " Title:" & Title\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    Exit Function\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print Err.Description, vbCritical, "' +
      className +
      '.GetSaveFilename"\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += '    \n';
    outtxtClass += 'End Function\n';
    outtxtClass =
      outtxtClass +
      'Sub SetOutputFilenameWithDialog(Optional DTitle = ' +
      '"Fájl", Optional FFilter = "Excel files , *.xlsx")\n';
    outtxtClass =
      outtxtClass +
      '    ' +
      '\'' +
      'Comments  : Load data from SAP exported CSV (text)\n';
    outtxtClass += '    \'Created by: Laszlo Tamas\n';
    outtxtClass += '    \n';
    outtxtClass += '    Dim strRes As String\n';
    outtxtClass += '    \n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    strRes = GetSaveFilename(DTitle, FFilter)\n';
    outtxtClass += '    m_strOutputFilename = strRes\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "' +
      className +
      '.SetOutputFilenameWithDialog FullFilename ' +
      'has been set to: " & m_strOutputFilename\n';
    outtxtClass += '    \n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    Exit Sub\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Sub ' +
      className +
      '.SetOutputFilenameWithDialog"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += 'End Sub\n';
    outtxtClass += 'Sub LoadText(Optional PasteSpec As Boolean = False)\n';
    outtxtClass += '    On Error GoTo PROC_ERR\n';
    outtxtClass += '    Sheets(m_strSheetname).Select\n';
    outtxtClass += '    Cells.Select\n';
    outtxtClass += '    Selection.NumberFormat = "General"\n';
    outtxtClass += '    Selection.Delete Shift:=xlUp\n';
    outtxtClass += '    Range("A1").Select\n';
    outtxtClass += '    \n';
    outtxtClass += '    Workbooks.OpenText Filename:=m_strFullFilename, _\n';
    outtxtClass += '        Origin:=XLOrigin_UTF, _\n';
    outtxtClass += '            DataType:=xlDelimited, _\n';
    outtxtClass += '                Semicolon:=True\n';
    outtxtClass += '    \n';
    outtxtClass += '    mstrOpenWorkbookName = ActiveWorkbook.Name\n';
    outtxtClass += '    Cells.Select\n';
    outtxtClass += '    Selection.Copy\n';
    outtxtClass += '    Windows(mstrMainWorkbookName).Activate\n';
    outtxtClass += '    Sheets(m_strSheetname).Select\n';
    outtxtClass += '    Range("A1").Select\n';
    outtxtClass += '    If PasteSpec Then\n';
    outtxtClass =
      outtxtClass +
      '        Selection.PasteSpecial Paste:=xlPasteValues, ' +
      'Operation:=xlNone, SkipBlanks _\n';
    outtxtClass += '            :=False, Transpose:=False\n';
    outtxtClass += '    Else\n';
    outtxtClass += '        ActiveSheet.Paste\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Windows(mstrOpenWorkbookName).Activate\n';
    outtxtClass += '    Application.DisplayAlerts = False\n';
    outtxtClass += '    ActiveWindow.Close (False)\n';
    outtxtClass += '    Application.DisplayAlerts = True\n';
    outtxtClass += '    Windows(mstrMainWorkbookName).Activate\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_EXIT:\n';
    outtxtClass += '    On Error GoTo 0\n';
    outtxtClass += '    \'Code here\n';
    outtxtClass += '    \n';
    outtxtClass += '    Exit Sub\n';
    outtxtClass += '    \n';
    outtxtClass += '    PROC_ERR:\n';
    outtxtClass =
      outtxtClass +
      '    Debug.Print "Error in Sub ' +
      className +
      '.LoadText"\n';
    outtxtClass += '    If Err.Number Then\n';
    outtxtClass += '        \'MsgBox Err.Description\n';
    outtxtClass += '        Debug.Print Err.Description\n';
    outtxtClass += '    End If\n';
    outtxtClass += '    Resume PROC_EXIT\n';
    outtxtClass += 'End Sub\n';
  }

  $('#textClass').val(outtxtClass);
}

function createMainModule() {
  var projectName = capitalizeFirstLetter($('#NameProj').val());
  var outtextMain = '';
  outtextMain += 'Option Explicit\n';
  outtextMain = outtextMain + 'Sub ' + projectName + '()\n';
  outtextMain += '    \'Description\n';
  outtextMain += '    \'Parameters:\n';
  outtextMain += '    \'Created by: Laszlo Tamas\n';
  outtextMain += '\n';
  outtextMain += '\n';
  outtextMain += '    On Error GoTo PROC_ERR\n';
  outtextMain += '\n';
  outtextMain += '    \'Code here\n';
  outtextMain += '\n';
  outtextMain += '    \'---------------\n';
  outtextMain += 'PROC_EXIT:\n';
  outtextMain += '    On Error GoTo 0\n';
  outtextMain += '    Exit Sub\n';
  outtextMain += 'PROC_ERR:\n';
  outtextMain =
    outtextMain + '    Debug.Print  "Error in Procedure ' + projectName + '"\n';
  outtextMain += '    If Err.Number Then\n';
  outtextMain += '        Debug.Print  Err.Description\n';
  outtextMain += '    End If\n';
  outtextMain += '    Resume PROC_EXIT\n';
  outtextMain += 'End Sub\n';
  outtextMain = outtextMain + 'Private Sub ' + projectName + 'Test\n';
  outtextMain = outtextMain + '    \'Test procedure for ' + projectName + '\n';
  outtextMain += '    Dim dtmStartTime As Date\n';
  outtextMain += '\n';
  outtextMain += '\n';
  outtextMain += '\n';
  outtextMain += '    dtmStartTime = Now()\n';
  outtextMain = outtextMain + '    Call ' + projectName + '()\n';
  outtextMain += 'End Sub\n';

  if ($('#CompCheck1').prop('checked')) {
    outtextMain += '\'----------------\n';
    outtextMain += '\'Columns and Rows\n';
    outtextMain += '\'----------------\n';
    outtextMain += 'Private Function Col_Letter(lngCol As Long) As String\n';
    outtextMain += '    \'Get letter from column number\n';
    outtextMain += '    Dim vArr\n';
    outtextMain += '    \n';
    outtextMain += '    \'  On Error Resume Next\n';
    outtextMain +=
      '    vArr = Split(Cells(1, lngCol).Address(True, False), "$")\n';
    outtextMain += '    Col_Letter = vArr(0)\n';
    outtextMain += 'End Function\n';
    outtextMain =
      outtextMain +
      'Private Function Col_LetterHeader(sheetName As String, ' +
      'headText As String, Optional headRow = 1) As String\n';
    outtextMain += '    \'Get column letter from header text\n';
    outtextMain += '    Dim lngColNumber As Long\n';
    outtextMain += '    \n';
    outtextMain =
      outtextMain +
      '    lngColNumber = Col_NumberHeader(sheetName, ' +
      'headText, headRow)\n';
    outtextMain += '    Col_LetterHeader = Col_Letter(lngColNumber)\n';
    outtextMain += 'End Function\n';
    outtextMain += 'Private Function Col_Number(colLetter) As Long\n';
    outtextMain += '    \'Get column number from column letter\n';
    outtextMain += '    Col_Number = Range(colLetter & "1").Column\n';
    outtextMain += 'End Function\n';
    outtextMain =
      outtextMain +
      'Private Function Col_NumberHeader(sheetName As String, ' +
      'headText As String, Optional headRow = 1) As Long\n';
    outtextMain += '    \'Get column number from header text\n';
    outtextMain += '    Dim i As Long\n';
    outtextMain += '    Dim strCellString As String\n';
    outtextMain += '    \n';
    outtextMain += '    Col_NumberHeader = 0\n';
    outtextMain += '    For i = 1 To 400\n';
    outtextMain =
      outtextMain +
      '        strCellString = ' +
      'Trim(CStr(Sheets(sheetName).Cells(headRow, i)))\n';
    outtextMain += '        If strCellString = headText Then\n';
    outtextMain += '            Col_NumberHeader = i\n';
    outtextMain += '            Exit Function\n';
    outtextMain += '        End If\n';
    outtextMain += '    Next i\n';
    outtextMain += 'End Function\n';
    outtextMain += 'Private Sub ColLetterTests()\n';
    outtextMain =
      outtextMain +
      '    ' +
      '\'' +
      'Test for Col_Letter, Col_LetterHeader, Col_Number ' +
      'and Col_NumberHeader\n';
    outtextMain += '    Debug.Print Col_Letter(12)\n';
    outtextMain +=
      '    Debug.Print Col_LetterHeader("Hogyallunk", "Any.csop.")\n';
    outtextMain += '    Debug.Print Col_Number("H")\n';
    outtextMain +=
      '    Debug.Print Col_NumberHeader("Hogyallunk", "Any.csop.")\n';
    outtextMain += 'End Sub\n';
    outtextMain =
      outtextMain +
      'Private Function GetLastRow(sheetName As String, ' +
      'checkColumn As Long, _\n';
    outtextMain += '    Optional firstrow = 2, Optional lastrow = 600000, _\n';
    outtextMain += '        Optional backwardCheck = True) As Long\n';
    outtextMain += '    \'Adott fül utolsó sora\n';
    outtextMain += '    Dim i As Long\n';
    outtextMain += '    Dim curSheet As Worksheet\n';
    outtextMain += '    Dim strCell As String\n';
    outtextMain += '    \n';
    outtextMain += '    Set curSheet = ActiveWorkbook.ActiveSheet\n';
    outtextMain += '    Sheets(sheetName).Activate\n';
    outtextMain += '    GetLastRow = 0\n';
    outtextMain += '    If backwardCheck Then\n';
    outtextMain += '        For i = lastrow To firstrow Step -1\n';
    outtextMain += '            strCell = Trim(CStr(Cells(i, checkColumn)))\n';
    outtextMain += '            If strCell <> "" Then\n';
    outtextMain += '                GetLastRow = i\n';
    outtextMain += '                Exit For\n';
    outtextMain += '            End If\n';
    outtextMain += '        Next i\n';
    outtextMain += '    Else\n';
    outtextMain += '        For i = firstrow To lastrow\n';
    outtextMain += '            strCell = Trim(CStr(Cells(i, checkColumn)))\n';
    outtextMain += '            If strCell = "" Then\n';
    outtextMain += '                GetLastRow = i - 1\n';
    outtextMain += '                Exit For\n';
    outtextMain += '            End If\n';
    outtextMain += '        Next i\n';
    outtextMain += '    End If\n';
    outtextMain += '    curSheet.Activate\n';
    outtextMain += '    Set curSheet = Nothing\n';
    outtextMain =
      outtextMain +
      '    Debug.Print "LastRow of " & sheetName & ": " & ' +
      'GetLastRow & " ChkCol:" & checkColumn\n';
    outtextMain += 'End Function\n';
    outtextMain += '\n';
    outtextMain += '\n';
  }
  if ($('#CompCheck2').prop('checked')) {
    outtextMain += '\'--------------\n';
    outtextMain += '\'Refresh ON OFF\n';
    outtextMain += '\'--------------\n';
    outtextMain += 'Private Sub RefreshOFF()\n';
    outtextMain += '    \'Screen update OFF\n';
    outtextMain += '    With Application\n';
    outtextMain += '        .ScreenUpdating = False\n';
    outtextMain += '        .EnableEvents = False\n';
    outtextMain += '        \'.Calculation = xlCalculationManual\n';
    outtextMain += '    End With\n';
    outtextMain += 'End Sub\n';
    outtextMain += 'Private Sub RefreshON()\n';
    outtextMain += '    \'Screen update ON\n';
    outtextMain += '    With Application\n';
    outtextMain += '        .ScreenUpdating = True\n';
    outtextMain += '        .EnableEvents = True\n';
    outtextMain =
      outtextMain +
      '        ' +
      '\'' +
      '.Calculation = xlCalculationAutomatic\n';
    outtextMain += '    End With\n';
    outtextMain += 'End Sub\n';
    outtextMain += '\n';
  }
  if ($('#CompCheck3').prop('checked')) {
    outtextMain += '\'----------------------\n';
    outtextMain += '\'Change keyboard layout\n';
    outtextMain += '\'----------------------\n';
    outtextMain += 'Private Sub SwitchToENG()\n';
    outtextMain += '    \'Váltás angolra\n';
    outtextMain += '    Call ActivateKeyboardLayout(1033, 0)\n';
    outtextMain += 'End Sub\n';
    outtextMain += 'Private Sub SwitchToHUN()\n';
    outtextMain += '    \'Váltás magyarra\n';
    outtextMain += '    Call ActivateKeyboardLayout(1038, 0)\n';
    outtextMain += 'End Sub\n';
    outtextMain += 'Private Sub SwitchToTUR()\n';
    outtextMain += '    \'Váltás törökre\n';
    outtextMain += '    Call ActivateKeyboardLayout(1055, 0)\n';
    outtextMain += 'End Sub\n';
    outtextMain += '\n';
    outtextMain += '\n';
  }
  if ($('#CompCheck4').prop('checked')) {
    outtextMain += '\'------------------------------------------\n';
    outtextMain =
      outtextMain +
      '' +
      '\'' +
      'Create Outlook Appointment for ' +
      projectName +
      '\n';
    outtextMain += '\'------------------------------------------\n';
    outtextMain =
      outtextMain +
      'Public Sub CreateAppt' +
      projectName +
      '(sSubject, sBodyText, sDate)\n';
    outtextMain =
      outtextMain +
      '    ' +
      '\'' +
      'A CreateObject módszerrel Office verzió független\n';
    outtextMain += '    Dim olApp As Object\n';
    outtextMain += '    \n';
    outtextMain += '    Set olApp = CreateObject("Outlook.Application")\n';
    outtextMain += '    \n';
    outtextMain += '    \'    Dim olApp As Outlook.Application\n';
    outtextMain += '    Dim olAppt As Object\n';
    outtextMain += '    \n';
    outtextMain =
      outtextMain +
      '    Set olAppt = olApp.CreateItem(1) ' +
      '\'' +
      '0, mail, 1 appointment\n';
    outtextMain += '    \n';
    outtextMain += '    Dim blnCreated As Boolean\n';
    outtextMain += '    Dim olNs As Object\n';
    outtextMain += '    Dim CalFolder As Object\n';
    outtextMain += '    Dim subFolder As Object\n';
    outtextMain += '    Dim arrCal As String\n';
    outtextMain += '    \n';
    outtextMain += '    Dim i As Long\n';
    outtextMain += '    Dim strCalSubject As String\n';
    outtextMain += '    Dim strCalPlace As String\n';
    outtextMain += '    Dim strCalBody As String\n';
    outtextMain += '    Dim strStopString As String\n';
    outtextMain += '    Dim colNev\n';
    outtextMain += '    Dim colSznap\n';
    outtextMain += '    \n';
    outtextMain += '    \'    Dim oPattern As RecurrencePattern\n';
    outtextMain += '    \n';
    outtextMain += '    colNev = 1\n';
    outtextMain += '    colSznap = 2\n';
    outtextMain += '    \n';
    outtextMain += '    strStopString = ""\n';
    outtextMain += '    \n';
    outtextMain += '    On Error GoTo PROC_ERR\n';
    outtextMain += '    \n';
    outtextMain += '    strCalPlace = vbNullString\n';
    outtextMain += '    strCalSubject = vbNullString\n';
    outtextMain += '    \n';
    outtextMain += '    On Error Resume Next\n';
    outtextMain += '    \'    Set olApp = Outlook.Application\n';
    outtextMain += '    \'\n';
    outtextMain += '    \'    If olApp Is Nothing Then\n';
    outtextMain += '    \'        Set olApp = Outlook.Application\n';
    outtextMain += '    \'        blnCreated = True\n';
    outtextMain += '    \'        Err.Clear\n';
    outtextMain += '    \'    Else\n';
    outtextMain += '    \'        blnCreated = False\n';
    outtextMain += '    \'    End If\n';
    outtextMain += '    \n';
    outtextMain += '    On Error GoTo 0\n';
    outtextMain += '    \n';
    outtextMain += '    Set olNs = olApp.GetNamespace("MAPI")\n';
    outtextMain =
      outtextMain +
      '    Set CalFolder = olNs.GetDefaultFolder(9) ' +
      '\'' +
      '9 olFolderCalendar\n';
    outtextMain += '    \n';
    outtextMain += '    \n';
    outtextMain += '    \n';
    outtextMain += '    \n';
    outtextMain =
      outtextMain +
      '    ' +
      '\'' +
      '    arrCal = cmstrCalSzuletesnapok ' +
      '\'' +
      'Naptár neve\n';
    outtextMain += '    \'\n';
    outtextMain =
      outtextMain +
      '    ' +
      '\'' +
      '    Set subFolder = CalFolder.Folders(arrCal)\n';
    outtextMain += '    \n';
    outtextMain =
      outtextMain +
      '    Set olAppt = CalFolder.Items.Add(1) ' +
      '\'' +
      '1 olAppointmentItem\n';
    outtextMain += '    \n';
    outtextMain =
      outtextMain +
      '    ' +
      '\'' +
      'MsgBox subFolder, vbOKCancel, "Folder Name"\n';
    outtextMain += '    \n';
    outtextMain += '    With olAppt\n';
    outtextMain += '        \n';
    outtextMain += '        \'Define calendar item properties\n';
    outtextMain += '        \n';
    outtextMain =
      outtextMain +
      '        ' +
      '\'' +
      'Debug.Print DateValue("2016." & ' +
      'Right(Trim(CStr(Cells(i, colSznap))), 4)) + ' +
      'TimeValue("0:00:00")\n';
    outtextMain =
      outtextMain +
      '        .Start = DateTime.DateSerial(CInt(Left(sDate, 4)), ' +
      'CInt(Mid(sDate, 6, 2)), CInt(Mid(sDate, 9, 2))) + ' +
      'TimeValue("9:00:00") ' +
      '\'' +
      'Given date 09:00\n';
    outtextMain += '        .Subject = sSubject\n';
    outtextMain += '        .Body = sBodyText\n';
    outtextMain += '        .ReminderSet = True\n';
    outtextMain =
      outtextMain +
      '        .ReminderMinutesBeforeStart = 4320 ' +
      '\'' +
      '3 days, 72 hours\n';
    outtextMain += '        .MeetingStatus = 0 \'0 olNonMeeting\n';
    outtextMain += '        .Save\n';
    outtextMain += '    End With\n';
    outtextMain += '    \n';
    outtextMain += '    Set olAppt = Nothing\n';
    outtextMain += '    Set olApp = Nothing\n';
    outtextMain += '    \n';
    outtextMain += 'PROC_EXIT:\n';
    outtextMain += '    \n';
    outtextMain += '    Exit Sub\n';
    outtextMain += '    \n';
    outtextMain += 'PROC_ERR:\n';
    outtextMain =
      outtextMain +
      '    MsgBox "An error occurred - Exporting items to Calendar." & ' +
      'vbCrLf & Err.Number & " " & Err.Description\n';
    outtextMain += '    Resume PROC_EXIT\n';
    outtextMain += 'End Sub\n';
    outtextMain =
      outtextMain + 'Private Sub CreateAppt' + projectName + '_Test()\n';
    outtextMain =
      outtextMain +
      '    Call CreateAppt' +
      projectName +
      '("Kórház2", "Teszt szöveg2", "2017.10.18")\n';
    outtextMain += 'End Sub\n';
    outtextMain += '\n';
    outtextMain += '\n';
    outtextMain += '\n';
  }

  $('#textMain').val(outtextMain);
}

function createMenu() {
  var projectName = capitalizeFirstLetter($('#NameProj').val());
  var ribbonName = projectName + 'Ribbon';

  var outtxtThisWorkbook = '';
  var outtxtMenuModule;
  outtxtThisWorkbook += 'Private Sub Workbook_BeforeClose(Cancel As Boolean)\n';
  outtxtThisWorkbook =
    outtxtThisWorkbook + '    Call DeleteRibbons' + projectName + '\n';
  outtxtThisWorkbook += 'End Sub\n';
  outtxtThisWorkbook += '\n';
  outtxtThisWorkbook += 'Private Sub Workbook_Open()\n';
  outtxtThisWorkbook =
    outtxtThisWorkbook + '    Call AddRibbons' + projectName + '\n';
  outtxtThisWorkbook += 'End Sub\n';
  $('#textThisWb').val(outtxtThisWorkbook);

  outtxtMenuModule = '';
  outtxtMenuModule += 'Option Explicit\n';
  outtxtMenuModule =
    outtxtMenuModule +
    'Global Const sToolbar' +
    projectName +
    ' As String = "' +
    ribbonName +
    '"\n';
  outtxtMenuModule += '\n';
  outtxtMenuModule =
    outtxtMenuModule + 'Public Sub AddRibbons' + projectName + '()\n';
  outtxtMenuModule =
    outtxtMenuModule +
    '    ' +
    '\'' +
    'Add user ribbons, call it from Workbook_Open\n';
  outtxtMenuModule =
    outtxtMenuModule + '    Call AddRibbonLine' + projectName + '\n';
  outtxtMenuModule += 'End Sub\n';
  outtxtMenuModule =
    outtxtMenuModule + 'Public Sub DeleteRibbons' + projectName + '()\n';
  outtxtMenuModule =
    outtxtMenuModule +
    '    ' +
    '\'' +
    'Delete ribbons, call it from Workbook_BeforeClose\n';
  outtxtMenuModule += '    On Error Resume Next\n';
  outtxtMenuModule =
    outtxtMenuModule +
    '    Application.CommandBars(sToolbar' +
    projectName +
    ').Delete\n';
  outtxtMenuModule += 'End Sub\n';
  outtxtMenuModule += '\n';

  outtxtMenuModule =
    outtxtMenuModule + 'Sub AddRibbonLine' + projectName + '()\n';
  outtxtMenuModule = outtxtMenuModule + '    \'' + ribbonName + '\n';
  outtxtMenuModule += '    Dim cbToolBar\n';
  outtxtMenuModule += '    \n';

  for (i = 1; i < menuNum + 1; i++) {
    if ($('#MenuCheck' + i).prop('checked')) {
      outtxtMenuModule =
        outtxtMenuModule + '    Dim ctButton' + i.toString() + '\n';
    }
  }

  outtxtMenuModule += '    \n';
  outtxtMenuModule += '    On Error Resume Next\n';
  outtxtMenuModule =
    outtxtMenuModule +
    '    Set cbToolBar = Application.CommandBars.Add(sToolbar' +
    projectName +
    ', msoBarTop, False, True)\n';
  outtxtMenuModule += '    With cbToolBar\n';

  for (i = 1; i < menuNum + 1; i++) {
    if ($('#MenuCheck' + i).prop('checked')) {
      outtxtMenuModule =
        outtxtMenuModule +
        '        Set ctButton' +
        i.toString() +
        ' = .Controls.Add(Type:=msoControlButton, ID:=2950)\n';
    }
  }

  outtxtMenuModule += '    End With\n';
  outtxtMenuModule += '    \n';

  for (i = 1; i < menuNum + 1; i++) {
    if ($('#MenuCheck' + i).prop('checked')) {
      outtxtMenuModule =
        outtxtMenuModule + '    With ctButton' + i.toString() + '\n';
      outtxtMenuModule =
        outtxtMenuModule + '        .Caption = "' + $('#Cap' + i).val() + '\n';
      outtxtMenuModule =
        outtxtMenuModule + '        .FaceId = ' + $('#Face' + i).val() + '\n';
      outtxtMenuModule =
        outtxtMenuModule +
        '        .OnAction = "' +
        $('#Onaction' + i).val() +
        '"\n';
      outtxtMenuModule =
        outtxtMenuModule +
        '        .TooltipText = "' +
        $('#TTip' + i).val() +
        '"\n';
      outtxtMenuModule += '        .Style = msoButtonIconAndCaption\n';
      outtxtMenuModule += '    End With\n';
      outtxtMenuModule += '    \n';
    }
  }

  outtxtMenuModule += '    \n';
  outtxtMenuModule += '    \n';
  outtxtMenuModule += '    \n';
  outtxtMenuModule += '    With cbToolBar\n';
  outtxtMenuModule += '        .Visible = True\n';
  outtxtMenuModule += '        .Protection = msoBarNoChangeVisible\n';
  outtxtMenuModule += '    End With\n';
  outtxtMenuModule += 'End Sub\n';

  outtxtMenuModule += 'Sub MenuNULL()\n';
  outtxtMenuModule += '    \'Empty dummy subroutine\n';
  outtxtMenuModule += 'End Sub\n';

  $('#textMenu').val(outtxtMenuModule);
}

function setPredef() {
  var predef = $('#ProjPredef').val();
  // TODO jQuery and formatting
  if (predef === 'SAP') {
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
  var ret;
  switch (varType) {
    case 'String':
      ret = 'str';
      break;
    case 'Long':
      ret = 'lng';
      break;
    case 'Integer':
      ret = 'int';
      break;
    case 'Boolean':
      ret = 'bln';
      break;
    case 'Double':
      ret = 'dbl';
      break;
    case 'Date':
      ret = 'dat';
      break;
    case 'Variant':
      ret = 'vnt';
      break;
    case 'Object':
      ret = 'obj';
      break;
    case 'SheetName':
      ret = 'sh';
      break;
    case 'Worksheet':
      ret = 'wst';
      break;
    case 'Outlook':
      ret = 'ol';
      break;
    default:
      break;
  }
  return ret;
}

function getConstInitValue(varType) {
  if (varType === 'String') {
    return ' = "text"';
  }
  if (varType === 'Long') {
    return ' = 0';
  }
  if (varType === 'Integer') {
    return ' = 0';
  }
  if (varType === 'Boolean') {
    return ' = True';
  }
  if (varType === 'Double') {
    return ' = 0.5';
  }
  if (varType === 'Date') {
    return ' = CDate("04/22/2016 12:00 AM")';
  }
  if (varType === 'Variant') {
    return ' = True';
  }
  if (varType === 'Object') {
    return ' = ';
  }
  if (varType === 'SheetName') {
    return ' = "Munka1"';
  }
  if (varType === 'Worksheet') {
    return ' = ActiveSheet';
  }
  if (varType === 'Outlook') {
    return ' = ';
  }
}

function getDeclareType(varType) {
  if (varType === 'String') {
    return 'String';
  }
  if (varType === 'Long') {
    return 'Long';
  }
  if (varType === 'Integer') {
    return 'Integer';
  }
  if (varType === 'Boolean') {
    return 'Boolean';
  }
  if (varType === 'Double') {
    return 'Double';
  }
  if (varType === 'Date') {
    return 'Date';
  }
  if (varType === 'Variant') {
    return 'Variant';
  }
  if (varType === 'Object') {
    return 'Object';
  }
  if (varType === 'SheetName') {
    return 'String';
  }
  if (varType === 'Worksheet') {
    return 'Worksheet';
  }
  if (varType === 'Outlook') {
    return 'Outlook';
  }
}
