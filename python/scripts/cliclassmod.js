/// <reference path="../../typings/globals/jquery/index.d.ts" />
var argNum = 6;
var funcNum = 5;
$(document).ready(function () {
    var i = 1;
    for (i = 1; i < argNum + 1; i++) {
        $("#tableargs").find('tbody')
            .append($('<tr>')
                .append($('<td class="Col1">Argument' + i + ':</td>'))
                .append($('<td class="Col2"><input id="argsSh' + i + '" type="text" class="full" /></td>'))
                .append($('<td class="Col1"><input id="args' + i + '" type="text" class="full" /></td>'))
                .append($('<td class="half"><input id="help' + i + '" type="text" class="full" /></td>'))
                .append($('<td class="Col2"><input id="argscheck' + i + '" type="checkbox" />Boolean</td>'))
            );
    }
    for (i = 1; i < funcNum + 1; i++) {
        $("#tablefuncs").find('tbody')
            .append($('<tr>')
                .append($('<td class="titleCol2">Function' + i + ':</td>'))
                .append($('<td class="half"><input id="funcs' + i +
                    '" type="text" class="full" /></td>'))
                .append($('<td>&nbsp;</td>'))
            );
    }

    $('#generate').click(function () {
        generateModules();
    });
    $('#clear').click(function () {
        resetNames();
    });
});

function generateModules() {
    generateMain();
}

function resetNames() {

}

function generateMain() {
    var textString = "";
    textString = textString + '# -*- coding: utf-8 -*-\n';
    textString = textString + '"""\n';
    textString = textString + 'This module deals with code.\n';
    textString = textString + '"""\n';
    textString = textString + '\n';
    textString = textString + 'import sys\n';
    textString = textString + 'import os\n';
    textString = textString + 'import argparse\n';
    textString = textString + '\n';
    textString = textString + '\n';
    textString = textString + '__author__ = "Laszlo Tamas"\n';
    textString = textString + '__copyright__ = "Copyright (c) 2048 Laszlo Tamas"\n';
    textString = textString + '__licence__ = "MIT"\n';
    textString = textString + '__version__ = "1.0"\n';
    textString = textString + '\n';
    textString = textString + '\n';

    textString = textString + 'class ' + $('#ClassName').val() + '(object):\n';
    textString = textString + '    """Class to deal with code\n';
    textString = textString + '    """\n';
    textString = textString + '\n';
    textString = textString + '    def __init__(self):\n';
    textString = textString + '        pass\n';
    textString = textString + '\n';
    textString = textString + '    def sample_func(self):\n';
    textString = textString + '        pass\n';    
    textString = textString + '\n';
    textString = textString + '\n';
    textString = textString + 'def parse_arguments():\n';
    textString = textString + '    """\n';
    textString = textString + '    Parse program arguments.\n';
    textString = textString + '\n';
    textString = textString + '    @return arguments\n';
    textString = textString + '    """\n';
    textString = textString + '    parser = argparse.ArgumentParser()\n';

    textString = textString + '    parser.add_argument(' + "'" + '-d' + "'" + ', ' + "'" + '--directory' + "'" + ', help=' + "'" + 'ZPL template folder' + "'" + ')\n';
    textString = textString + '    parser.add_argument(' + "'" + '-l' + "'" + ', ' + "'" + '--label' + "'" + ', help=' + "'" + 'ZPL label template' + "'" + ')\n';
    textString = textString + '    parser.add_argument(' + "'" + '-t' + "'" + ', ' + "'" + '--text' + "'" + ', help=' + "'" + 'input text' + "'" + ')\n';

    textString = textString + '    parser.add_argument(' + "'" + '-f' + "'" + ', ' + "'" + '--function' + "'" + ', type=str,\n';
    textString = textString + '                        choices=[' + "'" + 'convert_utf8' + "'" + '],\n';
    textString = textString + '                        help=' + "'" + 'function to execute' + "'" + ')\n';

    textString = textString + '    parser.add_argument(' + "'" + '-v' + "'" + ', ' + "'" + '--verbose' + "'" + ', action=' + "'" + 'store_true' + "'" + ',\n';
    textString = textString + '                        help=' + "'" + 'increase output verbosity' + "'" + ')\n';
    textString = textString + '    return parser.parse_args()\n';

    textString = textString + '\n';
    textString = textString + '\n';
    textString = textString + 'def execute_program():\n';
    textString = textString + '    """Execute the program by arguments.\n';
    textString = textString + '    """\n';
    textString = textString + '    args = parse_arguments()\n';

    textString = textString + '    if args.function == ' + "'" + 'convert_utf8' + "'" + ':\n';
    textString = textString + '        pass\n';

    textString = textString + '\n';
    textString = textString + 'if __name__ == ' + "'" + '__main__' + "'" + ':\n';
    textString = textString + '    ' + $('#ClassName').val().toUpperCase() + ' = ' + $('#ClassName').val() + '()\n';
    textString = textString + '    execute_program()\n';
    textString = textString + '    sys.exit()\n';

    $('#TextAreaMain').val(textString);

}