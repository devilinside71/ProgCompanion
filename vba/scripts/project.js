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
            .append($('<td class="Col1">Menu'+i+'</td>'))
            .append($('<td class="checkBoxCol"><input type="checkbox" id="MenuCheck'+i+'" /></td>'))
            .append($('<td class="Col2"><input id="Cap'+i+'" type="text" value="Menu'+i+'" class="textBox" /></td>'))
            .append($('<td class="Col3"><input id="Face'+i+'" type="text" value="1244" class="textBox" /></td>'))
            .append($('<td class="Col4"><input id="Onaction'+i+'" type="text" value="MenuNULL" class="textBox" /></td>'))
            .append($('<td class="Col5"><input id="TTip'+i+'" type="text" value="Menu'+i+' végrehajtása" class="textBox" /></td>'))
        );
    }
    for (i = 1; i < propNum + 1; i++) {
        $("#tableProps").find('tbody')
            .append($('<tr>')
            .append($('<td class="Col1">Property'+i+'</td>'))
            .append($('<td class="checkBoxCol"><input type="checkbox" id="ClassCheck'+i+'" /></td>'))
            .append($('<td class="Col2"><input id="Clprop'+i+'" type="text" value="Prop'+i+'" class="textBox" /></td>'))
            .append($('<td class="Col3"><input id="Clpar'+i+'" type="text" value="Par'+i+'" class="textBox" /></td>'))
            .append($('<td class="Col4"><select id="Type'+i+'">'))
            .append($('<td class="Col5"><select id="Mode'+i+'">'))
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
});

function createProject() {
    createMenu();
    createMainModule();
    createClass();
}

function resetNames(params) {

}

function createClass(params) {

}

function setPredef(params) {

}

function createMenu(params) {

}

function createMainModule(params) {

}