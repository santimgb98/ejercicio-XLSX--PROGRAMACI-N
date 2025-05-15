const xlsx = require('xlsx');
const prompt = require("prompt-sync")();
const fs = require('fs');
const path = require('path')
const { log } = require('console');
const workbook = xlsx.utils.book_new();

class Menu{

    crear_archivo(){
        
        // Crear libro nuevo
        const workbook = xlsx.utils.book_new();
        const nombre_libro_nuevo = prompt("Introduce el nombre del nuevo libro de excel: ");
        
        // Contenido celda A1
        xlsx.utils.book_append_sheet(workbook,xlsx.utils.aoa_to_sheet([['Header']]))

        // Creación del archivo xlsx
        xlsx.writeFile(workbook, `./xlsx/${nombre_libro_nuevo}.xlsx`);
        console.clear()
    }
    listar_archivos(){
          fs.readdir('./xlsx', (files) => {
            console.log("Archivos XLSX:");
            files.forEach(file => {
                if (path.extname(file) == ".xlsx")
                console.log(file);
            });
        });
    }
    leer_archivo(){
        const excel = xlsx.readFile("./xlsx/libro2.xlsx");
        const nombreHoja = excel.SheetNames;

        let datos = xlsx.utils.sheet_to_csv(excel.Sheets[nombreHoja[0]],{
            cellDates : true
        })
        log(datos);
       
    }
    borrar_archivo(){
        fs.unlinkSync("./xlsx/libro3.xlsx");
        console.clear()
    }
    editar_archivo(){
        var sheet = workbook.Sheets['libro1'];

        // Modificar la celda A1
        sheet[0] = 'Nuevo valor';

        // Agregar una nueva fila
        var nuevaFila = ['Nueva fila', 'valor 1', 'valor 2'];
        xlsx.utils.sheet_add_aoa(sheet, [nuevaFila], {start: 'B3'});

        // Guardar los cambios
        xlsx.writeFile(workbook, "./xlsx/libro1.xlsx");
    }
}

module.exports = Menu;
