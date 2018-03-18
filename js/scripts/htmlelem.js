/// <reference path="../../typings/globals/jquery/index.d.ts" />

$(document).ready(function () {
    $('#generate').click(function () {
        generateCodes();
    });
});

function generateCodes() {
    generateHTMLCode();
    generateJSCode();
}

/**
 * Generate HTML code
 * @constructor
 */
function generateHTMLCode() {
    var textvar = "";
    if ($('#htmlElemType').val() == "Select") {
        textvar = textvar + '<select id="' + $('#htmlElemId').val() +
            '" name="' + $('#htmlElemId').val() + '">\n';
        textvar = textvar + '    <option value="-1"></option>\n';
        textvar = textvar + '</select>\n';
    }

    if ($('#htmlElemType').val() == "Button") {
        textvar = textvar + '<input id="' + $('#htmlElemId').val() +
            '" name="' + $('#htmlElemId').val() +
            '" type="button" value="button" />\n';
    }
    $('#htmlElemHtml').val(textvar)
}

/**
 * Generate Javascript code
 * @constructor
 */
function generateJSCode() {
    var textVar = "";
    if ($('#htmlElemType').val() == "Select") {
        textVar = textVar + '//Values for ' + $('#htmlElemId').val() + '\n';
        textVar = textVar + 'var ' + $('#htmlElemId').val() +
            'Values = ["One", "Two", "Three"];\n';
        textVar = textVar + '\n';
    }
    textVar = textVar + '/**\n';
    textVar = textVar + ' * Init\n';
    textVar = textVar + ' * @constructor\n';
    textVar = textVar + ' */\n'
    textVar = textVar + 'function start() {\n';
    textVar = textVar + '    document.getElementById(' + "'" +
        $('#htmlElemId').val() +
        "'" + ').addEventListener(' + "'" + $('#htmtlElemEventType').val() +
        "'" + ', ' + $('#htmlElemId').val() +
        capitalizeFirstLetter($('#htmtlElemEventType').val()) + ', false);\n';
    if ($('#htmlElemType').val() == "Select") {
        textVar = textVar + '    ' + $('#htmlElemId').val() + 'Fill(true);\n';
    }
    textVar = textVar + '}\n';
    textVar = textVar + '\n';
    textVar = textVar + '/**\n';
    textVar = textVar + ' * ' +
        capitalizeFirstLetter($('#htmtlElemEventType').val()) +
        ' eventhandler for ' + $('#htmlElemId').val() + '\n';
    textVar = textVar + ' * @constructor\n';
    textVar = textVar + ' */\n';
    textVar = textVar + 'function ' + $('#htmlElemId').val() +
        capitalizeFirstLetter($('#htmtlElemEventType').val()) + '() {\n';
    textVar = textVar + '    \n';
    textVar = textVar + '}\n';
    textVar = textVar + '\n';
    if ($('#htmlElemType').val() == "Select") {
        textVar = textVar + '/**\n';
        textVar = textVar + ' * Fill up ' + $('#htmlElemId').val() + '\n';
        textVar = textVar + ' * @constructor\n';
        textVar = textVar + ' */\n';
        textVar = textVar + 'function ' + $('#htmlElemId').val() +
            'Fill(startWithEmpty) {\n';
        textVar = textVar + '    var selectElem = document.getElementById("' +
            $('#htmlElemId').val() + '");\n';
        textVar = textVar + '    var i;\n';
        textVar = textVar +
            '    for (i = selectElem.options.length - 1; i >= 0; i--) {\n';
        textVar = textVar + '        selectElem.remove(i);\n';
        textVar = textVar + '    }\n';
        textVar = textVar + '    if (startWithEmpty){\n';
        textVar = textVar + '	     var emptyOpt = document.createElement(' +
            "'" + 'option' + "'" + ');\n';
        textVar = textVar + '		 emptyOpt.value = -1;\n';
        textVar = textVar + '		 emptyOpt.innerHTML = "";\n';
        textVar = textVar + '        selectElem.appendChild(emptyOpt);\n';
        textVar = textVar + '    }\n';
        textVar = textVar + '    for (i = 0; i < ' + $('#htmlElemId').val() +
            'Values.length; i++) {\n';
        textVar = textVar + '        var ' + $('#htmlElemId').val() + 'Val = ' +
            $('#htmlElemId').val() + 'Values[i];\n';
        textVar = textVar + '        var opt = document.createElement(' + "'" +
            'option' + "'" + ');\n';
        textVar = textVar + '        opt.value = i;\n';
        textVar = textVar + '        opt.innerHTML = ' +
            $('#htmlElemId').val() + 'Val;\n';
        textVar = textVar + '        selectElem.appendChild(opt);\n';
        textVar = textVar + '    }\n';
        textVar = textVar + '}\n';
    }
    $('#htmlElemJavascript').val(textVar);
}

function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}