$(".open-for-table").click(function () {

    var fromtable = $(this).attr("table");

    var table = JSON.parse(mainWeb.openReadForm(fromtable));

    for (var i = 0; i < table.Data.length; i++) {

        AddRow(fromtable ,table.Data[i], "");
    }
})

function AddRow(table, stroka, classstyle) {
    $(".excel-table" + table).append("<tr class=" + classstyle + "><td> " + stroka[0] + stroka[1] + "</td><td>" + stroka[2] + "</td></tr> ");
}

$(".equals-table").click(function () {

    var otchet = JSON.parse(mainWeb.equalsTable());
    //mainWeb.showDevTools();

    //console.log(otchet);

    $(".excel-table1").html("");
    $(".excel-table2").html("");


    for (var i = 0; i < otchet.ErrorTable1.length; i++) {

        AddRow("1", otchet.ErrorTable1[i], "red-row");
    }

    for (var i = 0; i < otchet.ErrorTable2.length; i++) {

        AddRow("2", otchet.ErrorTable2[i], "red-row");
    }

    for (var i = 0; i < otchet.Contains.length; i++) {

        AddRow("1", otchet.Contains[i], "green-row");
        AddRow("2", otchet.Contains[i], "green-row");
    }

    for (var i = 0; i < otchet.RestTable1.length; i++) {

        AddRow("1", otchet.RestTable1[i], "orange-row");
    }
    for (var i = 0; i < otchet.RestTable2.length; i++) {

        AddRow("2", otchet.RestTable2[i], "orange-row");
    }
});

