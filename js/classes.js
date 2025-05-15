const xlsx = require('xlsx');
const prompt = require("prompt-sync")();
const fs = require('fs');
const path = require('path');

class Menu{

    constructor(){
        this.workbook = null;
    }
    crear_archivo(){ // FUNCIONA
        
        // Crear libro nuevo
        const workbook = xlsx.utils.book_new();
        const nombre_libro_nuevo = prompt("Introduce el nombre del nuevo libro de excel: ");
        
        // Contenido celda A1
        xlsx.utils.book_append_sheet(workbook,xlsx.utils.aoa_to_sheet([['Header']]),"Hoja1")

        // Creación del archivo xlsx
        xlsx.writeFile(workbook, `../xlsx/${nombre_libro_nuevo}.xlsx`);
        console.clear()
    }
    listar_archivos(){
          fs.readdir('../xlsx/', (documentos) => {

            console.log("Archivos XLSX:");
            documentos.forEach(documento => {
                if (path.extname(documento) === ".xlsx")
                console.log(documento);
            });
        });
    }
    leer_archivo(){ // FUNCIONA
        const excel = xlsx.readFile("./xlsx/libro2.xlsx");
        const nombreHoja = excel.SheetNames;

        let datos = xlsx.utils.sheet_to_csv(excel.Sheets[nombreHoja[0]],{
            cellDates : true
        })
        console.log(datos);
       
    }
    borrar_archivo(){
        fs.unlinkSync("../xlsx/libro3.xlsx");
        console.clear()
    }
    editar_archivo(){
        this.workbook = xlsx.readFile("../xlsx/libro1.xlsx")
        const sheetName = this.workbook.SheetNames[0];
        const sheet = this.workbook.Sheets[sheetName];

        // Modificar la celda A1
        sheet['A1'] = {v:'Nuevo valor'};

        xlsx.utils.sheet_add_aoa(sheet, [['nuevaFila','valor1','valor2']], {origin: 'B3'});

        // Guardar los cambios
        xlsx.writeFile(this.workbook, "../xlsx/libro1.xlsx");
    }
}

module.exports = Menu;
