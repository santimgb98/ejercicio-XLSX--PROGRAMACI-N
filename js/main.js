const prompt = require("prompt-sync")();
const Menu = require( "./classes.js");
const { brotliCompress } = require("zlib");

const menu_excel = new Menu();
while(true){
    console.log(`MENÚ EXCEL`);
    console.log(`
    1. Crear archivo excel
    2. Listar archivos
    3. Leer archivo excel
    4. Borrar archivo excel
    5. Editar archivo excel
    6. Salir"
    -----------------------`);

    var option = prompt("Selecciona una opción del menú: ");
    switch(option){
        case "1":{
            menu_excel.crear_archivo();
        }break;
        case "2":{
            menu_excel.listar_archivos();
        }break;
        case "3":{
            menu_excel.leer_archivo();
        }break;
        case "4":{
            menu_excel.borrar_archivo();
        }break;
        case "5":{
            menu_excel.editar_archivo()
        }break;
        case "6":{
            process.exit();
        }
    }
    
}