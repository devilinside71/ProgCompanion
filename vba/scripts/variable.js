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
                .append(
                    $('<td class="nameColumn"><input type="text" id="Name' + i +
                        '" />'))
                .append(
                    $('<td class="otherColumns"><select id="Dimension' + i +
                        '">'))
                .append(
                    $('<td class="otherColumns"><select id="Scope' + i +
                        '">'))
                .append(
                    $('<td class="otherColumns"><select id="Type' + i +
                        '">'))
            );
    }
    for (i = 1; i < elemNum + 1; i++) {
        $('#Dimension' + i)
            .append($('<option value="Normal">Normal</option>'))
            .append($('<option value="Constant">Constant</option>'))
            .append($('<option value="Array">Array</option>'));
    }
    for (i = 1; i < elemNum + 1; i++) {
        $('#Scope' + i)
            .append($('<option value="Procedure">Procedure</option>'))
            .append($('<option value="Module">Module</option>'))
            .append($('<option value="Global">Global</option>'));
    }
    for (i = 1; i < elemNum + 1; i++) {
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

    $('#generate').click(function () {
        declareVars();
    });
    $('#clear').click(function () {
        resetNames();
    });
});

function declareVars() {
    var declarations = "";
    var initvals = "";
    var i = 1;
    for (i = 1; i < elemNum + 1; i++) {
        // var name = document.getElementById("Name" + i.toString()).value;
        var name = $('#Name' + i).val();
        var dimension = $('#Dimension' + i).val()
        var scope = $('#Scope' + i).val()
        var type = $('#Type' + i).val()
        if (name != "") {
            declarations = declarations +
                getDeclars(name, dimension, scope, type) + "\n";
            initvals = initvals +
                getInitValues(name, dimension, scope, type) + "\n";
        }
    }
    declarations = declarations + "--------------------------\n"
    declarations = declarations + initvals;

    $('#Code').val(declarations);
}

function getDeclars(name, dimension, scope, type) {
    var declaration = "";

    if (dimension == "Normal") {
        declaration = declaration +
            getDeclareCommand(scope) +
            " " + getScopePrefix(scope) + getPrefix(type) +
            capitalizeFirstLetter(name) +
            " As " + getDeclareType(type) +
            "\n";
    }
    if (dimension == "Constant") {
        declaration = declaration +
            getDeclareConstCommand(scope) +
            "Const c" + getScopePrefix(scope) + getPrefix(type) +
            capitalizeFirstLetter(name) +
            " As " + getDeclareType(type) +
            getConstInitValue(type) +
            "\n";
    }
    return declaration;
}

function getInitValues(name, dimension, scope, type) {
    var declaration = "";
    if (dimension == "Normal") {
        declaration = declaration +
            getScopePrefix(scope) + getPrefix(type) +
            capitalizeFirstLetter(name) +
            getInitValue(type) +
            "\n";
    }
    return declaration;
}

function getInitValue(varType) {
    if (varType == "String") {
        return " = vbNullString";
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
        return " = False";
    }
    if (varType == "Object") {
        return " = Nothing";
    }
    if (varType == "SheetName") {
        return ' = "Sheet"';
    }
    if (varType == "Worksheet") {
        return " = ActiveSheet";
    }
    if (varType == "Outlook") {
        return " = Nothing";
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


function getScopePrefix(scope) {
    if (scope == "Procedure") {
        return "";
    }
    if (scope == "Module") {
        return "m";
    }
    if (scope == "Global") {
        return "g";
    }

}


function getDeclareCommand(scope) {
    if (scope == "Procedure") {
        return "Dim";
    }
    if (scope == "Module") {
        return "Private";
    }
    if (scope == "Global") {
        return "Global";
    }
}

function getDeclareConstCommand(scope) {
    if (scope == "Procedure") {
        return "";
    }
    if (scope == "Module") {
        return "Private ";
    }
    if (scope == "Global") {
        return "Global ";
    }
}

function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}

function resetNames() {
    for (var i = 1; i < elemNum + 1; i++) {
        $('#Name' + i).val("");
    }
}