const xlsx = require( 'xlsx-populate');
const prompt = require("prompt-sync")();

class Menu{
    constructor(){}
    crear_archivo(){
        const nombre_libro_nuevo = prompt("Introduce el nombre del nuevo libro de excel: ");
        xlsx.fromBlankAsync()
            .then(workbook =>{
            workbook.sheet(0).cell('A1').value("hello");
            workbook.sheet(0).cell('A2').value("bye");
            workbook.sheet(0).cell('A3').value("hello");

            // devuelve un nuevo libro de excel
            //return workbook.toFileAsync(`./xlsx/${nombre_libro_nuevo}.xlsx`);
            console.log(`../xlsx/${nombre_libro_nuevo}.xlsx`)
    })
    }
    leer_archivo(){

    }
    borrar_archivo(){

    }
    editar_archivo(){

    }
}

module.exports = Menu;

