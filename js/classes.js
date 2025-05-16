const xlsx = require('xlsx');
const prompt = require("prompt-sync")();
const fs = require('fs');
const path = require('path');

class Menu{

    constructor(){
        this.workbook = null;
    }
    crear_archivo(){ // CREAR ARCHIVOS
        console.clear();
        // Crear libro nuevo
        const workbook = xlsx.utils.book_new();
        const nombre_libro_nuevo = prompt("Introduce el nombre del nuevo libro de excel: ");
        
        xlsx.utils.book_append_sheet(workbook,xlsx.utils.aoa_to_sheet([['Header']]),"Hoja1")

        // Creación del archivo xlsx
        xlsx.writeFile(workbook, `./xlsx/${nombre_libro_nuevo}.xlsx`);
        console.clear()
    }
    listar_archivos() { // LISTAR ARCHIVOS
        console.clear();

        // Leer un directorio y cada uno se los subdirectorios que tiene
        fs.readdir('./xlsx',{recursive:true}, (err, documentos) => {
            if (err) {
                console.error("** Error al leer el directorio:", err);
                return;
            }
            console.log("---------------------------")
            console.log("Archivos XLSX:");
            documentos.forEach(documento => {
                // path para detectar que es un .xlsx, si lo es lo imprime
                // .extname para recoger unicamente el nombre del archivo 
                if (path.extname(documento) === ".xlsx") {
                    console.log(documento);
                }
                
            });
            console.log("---------------------------")
        });
}
    leer_archivo(){ // LEER ARCHIVO
        console.clear() ;
        fs.readdir('./xlsx',{recursive:true}, (err, documentos) => {
            if (err) {
                console.error("** Error al leer el directorio:", err);
                return;
            }
            console.log("---------------------------")
            console.log("Archivos XLSX:");
            documentos.forEach(documento => {
                if (path.extname(documento) === ".xlsx") {
                    console.log(documento);
                }
            });
            console.log("---------------------------")
        });
        const xlsxToRead = prompt('Selecciona un libro: ');
        

        try{
            const excel = xlsx.readFile(`./xlsx/${xlsxToRead}.xlsx`);
            const nombreHoja = excel.SheetNames;
            let datos = xlsx.utils.sheet_to_csv(excel.Sheets[nombreHoja[0]],{
                cellDates : true
            })
            console.log(`Contenido de ${xlsxToRead}: ${datos}`);
        }catch(err){
            console.log(`** No se ha leido el archivo debido a ${err}`)
        }
       
    }
    borrar_archivo(){
        console.clear();
        fs.readdir('./xlsx',{recursive:true}, (err, documentos) => {
            if (err) {
                console.error("** Error al leer el directorio:", err);
                return;
            }
            console.log("---------------------------")
            console.log("Archivos XLSX:");
            documentos.forEach(documento => {
                if (path.extname(documento) === ".xlsx") {
                    console.log(documento);
                }
            });
            console.log("---------------------------")
        });
        // Se introduce el libro que se desea eliminar
        const xlsxToRemove = prompt('Selecciona un libro: ');

        // Try-Catch para eliminar el libro/mostrar el fallo si no se completa la eliminación
        try{
            fs.unlinkSync(`./xlsx/${xlsxToRemove}.xlsx`);
            console.clear()
            console.log(`${xlsxToRemove} ha sido eliminado con éxito!`);
        }catch(err){
            console.log("No se ha podido completar correctamente la eliminación del documento")
            console.log(`** Se ha producido un ${err}`)
        }
    }
    editar_archivo(){
        console.clear();
        fs.readdir('./xlsx',{recursive:true}, (err, documentos) => {
            if (err) {
                console.error("** Error al leer el directorio:", err);
                return;
            }
            console.log("---------------------------")
            console.log("Archivos XLSX:");
            documentos.forEach(documento => {
                if (path.extname(documento) === ".xlsx") {
                    console.log(documento);
                }
            });
            console.log("---------------------------")
        });
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
