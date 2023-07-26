const table = document.getElementsByClassName("sheet-body")[0]
const rows = document.getElementsByClassName("rows")[0]
const columns = document.getElementsByClassName("columns")[0]

let tableExists = false
// const validateRowsInput = (input) => {
// }
const generateTable = () => {
    const rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value)
    if (rows.value === "" || columns.value === "") {
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'Invalid rows or columns count!',
        })
    }else{
        table.innerHTML = ""
        for (let i = 0; i < rowsNumber; i++) {
            let tableRow = ""
            for (let j = 0; j < columnsNumber; j++) {
                tableRow += `<td contenteditable></td>`
            }
            table.innerHTML += tableRow
        }
        if (rowsNumber > 0 && columnsNumber > 0) {
            tableExists = true
        }
    }
}

const ExportToExcel = (type, fn, dl) => {
    if (!tableExists) {
        return
    }
    const elt = table
    const wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
}