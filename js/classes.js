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
        const xlsxToRead = prompt('Selecciona un libro: ');
        const excel = xlsx.readFile(`./xlsx/${xlsxToRead}.xlsx`);
        const nombreHoja = excel.SheetNames;

        let datos = xlsx.utils.sheet_to_csv(excel.Sheets[nombreHoja[0]],{
            cellDates : true
        })
        console.log(`Contenido de ${xlsxToRead}: ${datos}`);
       
    }
    borrar_archivo(){
        // Se introduce el libro que se desea eliminar
        const xlsxToRemove = prompt('Selecciona un libro: ');

        // Try-Catch para eliminar el libro/mostrar el fallo si no se completa la eliminación
        try{
            fs.unlinkSync(`./xlsx/${xlsxToRemove}.xlsx`);
            console.clear()
            console.log(`${xlsxToRemove} ha sido eliminado con éxito!`);
        }catch(err){
            console.log("No se ha podido completar correctamente la eliminación del documento")
            console.log(`Se ha producido un ${err}`)
        }
    }
    editar_archivo(){
        const xlsxToEdit = prompt("Seleciona libro a editar: ");
        const cellStartToEdit = prompt("En qué celda empieza la edición: ");
        const newContent = prompt("Introduce el nuevo contenido: ");
        try{
            this.workbook = xlsx.readFile(`./xlsx/${xlsxToEdit}.xlsx`)
            const sheetName = this.workbook.SheetNames[0];
            const sheet = this.workbook.Sheets[sheetName];

            xlsx.utils.sheet_add_aoa(sheet, [[`${newContent}`]], {origin: `${cellStartToEdit}`});

            // Guardar los cambios
            xlsx.writeFile(this.workbook, `./xlsx/${xlsxToEdit}.xlsx`);
            console.log(`El ${xlsxToEdit} se ha editado correctamente!`);
        }catch(err){
            console.log("No se ha podido completar correctamente la edición del documento")
            console.log(`Se ha producido un ${err}`)
        }
    }
}

module.exports = Menu;
