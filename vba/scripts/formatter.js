/// <reference path="../../typings/globals/jquery/index.d.ts" />
$(document).ready(function () {


    $('#format').click(function () {
        formatVBA();
    });
    $('#clear').click(function () {
        clearCode();
    });
});

function formatVBA() {
    var commands = new Array("if", "else", "else if", "end if",
        "sub", "private sub", "public sub", "function",
        "private function", "public function",
        "end sub", "end function",
        "enum", "private enum", "public enum", "property",
        "private property", "public property",
        "end enum", "end property",
        "for", "next", "with", "end with", "do", "loop");
    var commandsUp = new Array("AppActivate",
        "Beep",
        "Call",
        "ChDir",
        "ChDrive",
        "Close",
        "Const",
        "Date",
        "Declare",
        "DeleteSetting",
        "Dim",
        "Do-Loop",
        "End",
        "Erase",
        "Error",
        "Exit Do",
        "Exit For",
        "Exit Function",
        "Exit Property",
        "Exit Sub",
        "FileCopy",
        "For",
        "Each",
        "Next",
        "Function",
        "Get",
        "GoSub",
        "Return",
        "GoTo",
        "If",
        "Then",
        "Else",
        "Input #",
        "Kill",
        "Let",
        "Line Input #",
        "Load",
        "Lock",
        "Unlock",
        "Mid",
        "MkDir",
        "Name",
        "On Error",
        "On",
        "Open",
        "Option Base",
        "Option Compare",
        "Option Explicit",
        "Option Private",
        "Print #",
        "Private",
        "Property Get",
        "Property Let",
        "Property Set",
        "Public",
        "Put",
        "RaiseEvent",
        "Randomize",
        "ReDim",
        "Rem",
        "Reset",
        "Resume",
        "RmDir",
        "SaveSetting",
        "Seek",
        "Select Case",
        "SendKeys",
        "Set",
        "SetAttr",
        "Static",
        "Stop",
        "Sub",
        "Time",
        "Type",
        "Unload",
        "While",
        "Wend",
        "Width #",
        "With",
        "Write #"

    );
    var commandsBefore = new Array("", "-", "-", "-",
        "0", "0", "0", "0", "0", "0",
        "0", "0",
        "0", "0", "0", "0", "0", "0",
        "0", "0",
        "", "-", "", "-", "", "-");
    var commandsAfter = new Array("+", "+", "+", "",
        "+", "+", "+", "+", "+", "+",
        "0", "0",
        "+", "+", "+", "+", "+", "+",
        "0", "0",
        "+", "", "+", "", "+", "");
    //sorokra bontás
    var lines = $('#Code').val().split("\n");
    var outText = "";
    var line = "";
    var beforeIndent = 0;
    var afterIndent = 0;
    var i = 0;
    var k = 0;
    for (i = 0; i < lines.length; i++) {
        line = lines[i].trim();

        for (k = 0; k < commands.length; k++) {
            if ((commands[k] + " ").toLowerCase() ==
                line.slice(0, commands[k].length + 1).toLowerCase().trim() +
                " ") {
                //outText = outText + getIndent(beforeIndent) + line + "QQQ\n";
                if (commandsAfter[k] == "+") {
                    //beforeIndent++;
                    afterIndent = 1;
                }
                if (commandsAfter[k] == "0") {
                    beforeIndent = 0;
                }
                if (commandsBefore[k] == "-") {
                    beforeIndent--;
                }
                if (commandsBefore[k] == "0") {
                    beforeIndent = 0;
                }
            }

        }
        //make uppercase
        for (k = 0; k < commandsUp.length; k++) {
            if ((commandsUp[k] + " ").toLowerCase() ==
                line.slice(0, commandsUp[k].length + 1).toLowerCase().trim() +
                " ") {
                line = line.replace(line.slice(0, commandsUp[k].length), commandsUp[k])
            }
        }
        outText = outText + getIndent(beforeIndent) + line + "\n";
        if (afterIndent == 1) {
            afterIndent = 0;
            beforeIndent++;
        }
    }
    $('#CodeFormat').val(outText);
}

function getIndent(num) {
    var res = "";
    var indent = "    ";
    var i = 0;
    for (i = 0; i < num; i++) {
        res = res + indent;
    }
    return res;
}

function clearCode() {
    $('#Code').val("");
    $('#CodeFormat').val("");
}