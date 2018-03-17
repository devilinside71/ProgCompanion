$(document).ready(function () {
    var elemNum = 7;
    for (var i = 1; i < elemNum + 1; i++) {
        $("#tabla").find('tbody')
            .append($('<tr>')
                .append($('<td class="nameColumn"><input name="Text1" type="text" id="Name' + i + '" />'))
                .append($('<td class="otherColumns"><select id="Dimension' + i + '">'))
                .append($('<td class="otherColumns"><select id="Scope' + i + '">'))
                .append($('<td class="otherColumns"><select id="Type' + i + '">'))
            );
    }
    for (var i = 1; i < elemNum + 1; i++) {
        $('#Dimension' + i)
            .append($('<option value="Normal">Normal</option>'))
            .append($('<option value="Constant">Constant</option>'))
            .append($('<option value="Array">Array</option>'));
    }
    for (var i = 1; i < elemNum + 1; i++) {
        $('#Scope' + i)
            .append($('<option value="Procedure">Procedure</option>'))
            .append($('<option value="Module">Module</option>'))
            .append($('<option value="Global">Global</option>'));
    }
    for (var i = 1; i < elemNum + 1; i++) {
        $('#Type' + i)
            .append($('<option value="String">String</option>'))
            .append($('<option value="Long">Long</option>'))
            .append($('<option value="Integer">Integer</option>'))
            .append($('<option value="Boolean">Boolean</option>'))
            .append($('<option value="Double">Double</option>'))
            .append($('<option value="Date">Date</option>'))
            .append($('<option value="Variant">Variant</option>'))
            .append($('<option value="Object">Object</option>'))
            .append($('<option value="SheetName">Sheet name</option>'))
            .append($('<option value="Worksheet">Worksheet</option>'))
            .append($('<option value="Outlook">Outlook</option>'));
    }

});