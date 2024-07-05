let table = document.getElementsByClassName("sheet-body")[0],
rows = document.getElementsByClassName("rows")[0],
columns = document.getElementsByClassName("columns")[0]
tableExists = false

const generateTable = () => {
    let rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value);
    //sweet alert
    if (isNaN(rowsNumber) || isNaN(columnsNumber) || rowsNumber <= 0 || columnsNumber <= 0) {
        Swal.fire({
            icon: 'error',
            title: 'Invalid Input',
            text: 'Please enter valid numbers for rows and columns.',
        });
        return;
    }
    table.innerHTML = ""
    for(let i=0; i<rowsNumber; i++){
        var tableRow = ""
        for(let j=0; j<columnsNumber; j++){
            tableRow += `<td contenteditable></td>`
        }
        table.innerHTML += tableRow
    }
    if(rowsNumber>0 && columnsNumber>0){
        tableExists = true
    }
}

const ExportToExcel = (type, fn, dl) => {
    if(!tableExists){
        Swal.fire({
            icon: 'error',
            title: 'No Table Found',
            text: 'Please generate a table before exporting.',
        });
        return;
    }
    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
}

// Event listeners for generate and export buttons
document.getElementById('generateButton').addEventListener('click', generateTable);
document.getElementById('exportButton').addEventListener('click', () => ExportToExcel('xlsx'));