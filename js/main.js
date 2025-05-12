const prompt = require("prompt-sync")();
const xlsx = require('xlsx-populate');
//const Menu = require( "./prueba.js");

class Menu{
    constructor(){}
    crear_archivo(){
        //const nombre_libro_nuevo = prompt("Introduce el nombre del nuevo libro de excel: ");
        xlsx.fromBlankAsync()
            .then(workbook =>{
            workbook.sheet(0).cell('A1').value("hello");
            workbook.sheet(0).cell('A2').value("bye");
            workbook.sheet(0).cell('A3').value("hello");

            // devuelve un nuevo libro de excel
            return workbook.toFileAsync("../nuevoLibro.xlsx");
            //console.log(`../xlsx/${nombre_libro_nuevo}.xlsx`)
    })
    }
    leer_archivo(){

    }
    borrar_archivo(){

    }
    editar_archivo(){

    }
}
const menu_excel = new Menu();
while(true){
    console.log(`MENÚ EXCEL`);
    console.log(`
    1. Crear archivo excel
    2. Leer archivo excel
    3. Borrar archivo excel
    4. Editar archivo excel
    5. Salir"
    -----------------------`);

    var option = prompt("Selecciona una opción del menú: ");
    while(isNaN(option)===true || (option < 1 && option > 5)){
        option = prompt("Selecciona una opción del menú: ");
    }

    switch(option){
        case "1":{
            menu_excel.crear_archivo();
        }break;
        case "2":{

        }break;
        case "3":{

        }break;
        case "4":{
            
        }break;
        case "5":{
            process.exit();
        }break;
    }
    
}