/// <reference path="../../typings/globals/jquery/index.d.ts" />
var elemNum = 6;
$(document).ready(function () {
    var i = 1;
    var typeOptions = new Array(
        "String", "Long", "Integer", "Boolean", "Double", "Date", "Variant",
        "Object", "SheetName", "Worksheet", "Outlook"
    );
    for (i = 1; i < elemNum + 1; i++) {
        $("#tabla").find('tbody')
            .append($('<tr>')
                .append($('<td class="nameColumn">Parameter' + i + ':</td>'))
                .append($(
                    '<td class="otherColumns"><input type="text" id="Name' +
                    i + '" />'))
                .append($(
                    '<td class="otherColumns"><select id="TypePar' + i + '">'))
                .append($('<td class="otherColumns">'))
            );
    }
    for (i = 1; i < elemNum + 1; i++) {
        $('#TypePar' + i)
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



    $('#generate').click(function () {
        createSub();
    });
    $('#clear').click(function () {
        resetNames();
    });
});

function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}

function resetNames() {
    var i = 1;
    for (i = 1; i < elemNum + 1; i++) {
        $('#Name' + i).val("");
    }
}

function createSub() {
    var funcText = "";
    var inputPars = "";
    var dimPars = "";
    var testPars = "";
    var i = 1;

    for (i = 1; i < elemNum + 1; i++) {
        if ($('#Name' + i).val() != "") {
            inputPars = inputPars +
                capitalizeFirstLetter($('#Name' + i).val()) +
                " As " + getDeclareType($('#TypePar' + i).val()) + ", ";
        }
    }

    if (inputPars != "") {
        inputPars = inputPars.slice(0, inputPars.length - 2);
    }
    if ($('#classFunction').prop('checked')) {
        funcText = "Sub ";
    } else {
        funcText = "Private Sub ";
    }
    funcText = funcText + capitalizeFirstLetter($('#NameFunc').val()) + "(" +
        inputPars +
        ")\n";
    funcText = funcText + "    '" + $('#remarkText').val() + "\n";
    funcText = funcText + "    'Parameters:\n";

    for (i = 1; i < elemNum + 1; i++) {
        if ($('#Name' + i).val() != "") {
            funcText = funcText + "    '           " +
                capitalizeFirstLetter($('#Name' + i).val()) + "\n";
        }
    }

    funcText = funcText + "    'Created by: Laszlo Tamas\n\n";

    funcText = funcText + "\n    On Error GoTo PROC_ERR\n\n";
    funcText = funcText + "    'Code here\n\n";
    funcText = funcText + "    '---------------\n";
    funcText = funcText + "PROC_EXIT:\n";
    funcText = funcText + "    On Error GoTo 0\n";
    funcText = funcText + "    Exit Sub\n";
    funcText = funcText + "PROC_ERR:\n";
    funcText = funcText + '    Debug.Print  "Error in Procedure ' +
        capitalizeFirstLetter($('#NameFunc').val()) + '"\n';
    funcText = funcText + "    If Err.Number Then\n";
    funcText = funcText + "        Debug.Print  Err.Description\n";
    funcText = funcText + "    End If\n";
    funcText = funcText + "    Resume PROC_EXIT\n";

    funcText = funcText + "End Sub\n";
    $('#Code').val(funcText);

    //Test code

    for (i = 1; i < elemNum + 1; i++) {
        if ($('#Name' + i).val() != "") {
            dimPars = dimPars + "    Dim " +
                getPrefix($('#TypePar' + i).val()) +
                capitalizeFirstLetter($('#Name' + i).val()) +
                " As " + getDeclareType($('#TypePar' + i).val()) + "\n";
        }
    }


    dimPars = dimPars + "\n";
    for (i = 1; i < elemNum + 1; i++) {
        if ($('#Name' + i).val() != "") {
            dimPars = dimPars + "    " + getPrefix($('#TypePar' + i).val()) +
                capitalizeFirstLetter($('#Name' + i).val()) +
                getConstInitValue($('#TypePar' + i).val()) + "\n";
        }
    }



    for (i = 1; i < elemNum + 1; i++) {
        if ($('#Name' + i).val() != "") {
            testPars = testPars + getPrefix($('#TypePar' + i).val()) +
                capitalizeFirstLetter($('#Name' + i).val()) + ", ";
        }
    }

    if (testPars != "") {
        testPars = testPars.slice(0, testPars.length - 2);
    }


    funcText = "Private Sub ";
    funcText = funcText +
        capitalizeFirstLetter($('#NameFunc').val()) + "Test\n";
    if ($('#classFunction').prop('checked')) {
        funcText = funcText + "    'Test procedure for " +
            capitalizeFirstLetter($('#className').val()) + "." +
            capitalizeFirstLetter($('#NameFunc').val()) + "\n";
    } else {
        funcText = funcText + "    'Test procedure for " +
            capitalizeFirstLetter($('#NameFunc').val()) + "\n";
    }
    if ($('#classFunction').prop('checked')) {
        funcText = funcText + "    Dim cl" +
            capitalizeFirstLetter($('#className').val()) + " As New " +
            capitalizeFirstLetter($('#className').val()) + "\n";
    }
    funcText = funcText + "    Dim dtmStartTime As Date\n";
    funcText = funcText + dimPars + "\n\n";
    funcText = funcText + "    dtmStartTime = Now()\n";
    if ($('#classFunction').prop('checked')) {
        funcText = funcText + '    Call cl' +
            capitalizeFirstLetter($('#className').val()) + '.' +
            capitalizeFirstLetter($('#NameFunc').val()) +
            '(' + testPars + ')\n';
        funcText = funcText + "    Set cl" +
            capitalizeFirstLetter($('#className').val()) + " = Nothing\n";
    } else {
        funcText = funcText + '    Call ' +
            capitalizeFirstLetter($('#NameFunc').val()) +
            '(' + testPars + ')\n';
    }
    funcText = funcText + "End Sub\n";

    $('#CodeTest').val(funcText);

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