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
    var commandsUp = new Array("If", "Else", "Else If", "End If",
        "Sub", "Private Sub", "Public Sub", "Function",
        "Private Function", "Public Function",
        "End Sub", "End Function",
        "Enum", "Private enum", "Public Enum", "Property",
        "Private Property", "Public Property",
        "End Enum", "End Property",
        "For", "Next", "With", "End With", "Do", "Loop");
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
            if ((commands[k] + " ").toLowerCase() ==
                line.slice(0, commands[k].length + 1).toLowerCase().trim() +
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