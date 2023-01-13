var focussedCells = []
var refedCells = []

function removeFocussed() {
    for (var i = 0; i < focussedCells.length; i++) {
        focussedCells[i].removeClass("cellselected")
    }
    focussedCells = []
    for (const el of refedCells) 
        el.removeClass("cellrefed")
    refedCells = []
}

function cell_click(cell) {
    $.ajax({
        url: $("#root").val() + "/getCellText",
        data: {
            cell: cell
        },
        success: function (result) {
            removeFocussed();
            $("#formula-bar").val(result.txt);
            $("#cell-name").html(cell.substring(1));
            var c = $("#" + cell)
            c.addClass("cellselected");
            focussedCells.push(c);
            for (const el of result.refedCells) {
                c = $(`#c${el}`)
                c.addClass("cellrefed")
                refedCells.push(c)
            }
            $("#formula-bar").focus();
            $("#message").html(result.errorMessage)
        }
    });
}


function formulaBar_keyUp(e) {
    if (e.code == "Enter" || e.code == "NumpadEnter") {
        $.ajax({
            url: $("#root").val() + "/setCellText",
            data: {
                cell: $("#cell-name").html(),
                text: $("#formula-bar").val()
            },
            success: function (result) {
                if (result.success)
                    location.reload(true)
                else
                    $("#message").html(result.message)
            }

        })
    }
}


$(window).on('load', function () {
    $("#formula-bar").keyup(formulaBar_keyUp);
});

var refedGridCells = []

function gridcell_click(cell) {
    const cellname = cell.header.name + (cell.rowIndex + 1);
    $.ajax({
        url: $("#root").val() + "/getCellText",
        data: {
            cell: 'c' + cellname
        },
        success: function (result) {
            $("#formula-bar").val(result.txt);
            $("#cell-name").html(cellname);
            var c = $("#c" + cellname)
            c.addClass("cellselected");
            focussedCells.push(c);
            refedGridCells = result.refedCells
            $("#formula-bar").focus();
            $("#message").html(result.errorMessage);
        }
    });
}

grid.style.width = '100%';

grid.addEventListener('click', function (e) {
    if (!e.cell) { return; }
    gridcell_click(e.cell);
});

grid.addEventListener('rendercell', function (e) {
    const cellname = e.cell.header.name + (e.cell.rowIndex + 1);
    if (refedGridCells.indexOf(cellname) != -1)
        e.ctx.fillStyle = '#E0FFE0';
});

