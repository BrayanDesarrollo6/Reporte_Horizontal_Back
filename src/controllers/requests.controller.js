// Almacenar objeto de funciones de enrutamiento - pertenece a router
const requestsController = {}
const { log } = require('console');
// Variables compartidas
const fs = require('fs');
const spawn = require("child_process").spawn;
const utf8 = require('utf8');
// Variables Reporte Horizontal
const XLSX = require('xlsx');
var Nombre_Horizontal = "";
var Lista_nombres = [];
// Constantes Archivo TxtSS
var Nombre_txt = "";

// Reporte Horizontal
requestsController.ReporteHorizontalResponse = (req, res) => {
    const ID_received = req.body.Data;
    // Obtener IDPeriodos
    let process;
    let estado = 0;
    let data_0 = ID_received.idproceso;
    let data_1 = ID_received.idperiodo;
    let data_2 = ID_received.idperiodo2;
    let data_3 = ID_received.idperiodo3;

    //console.log(data_0,data_1,data_2,data_3);
    
    if(data_1 !== undefined && data_1 !== null && data_2 === undefined || data_2 === null || data_2 === "" && data_3 === undefined)
    {
        estado = 1;
        process = spawn('python',["./src/python/ReporteHorizontal.py",estado,data_0,data_1]);
    }
    else if(data_1 !== undefined && data_1 !== null && data_2 !== undefined && data_2 !== null && data_3 === undefined || data_3 === null || data_3 === "")
    {
        estado = 2;
        process = spawn('python',["./src/python/ReporteHorizontal.py",estado,data_0,data_1,data_2]);
    }   
    else if(data_1 !== undefined && data_1 !== null && data_2 !== undefined && data_2 !== null && data_3 !== undefined && data_3 !== null)
    {
        estado = 3;
        process = spawn('python',["./src/python/ReporteHorizontal.py",estado,data_0,data_1,data_2,data_3]);
    }
    if(estado != 0){
        // const process = spawn('python',["./src/python/ReporteHorizontal.py",data]);
        process.stderr.on("data",(data)=>{
            console.error('stderr:',data.toString());
        })
        process.stdout.on('data', (data) => {
            Nombre_Horizontal = data.toString();
            Nombre_Horizontal = Nombre_Horizontal.split("\r\n").join("");
            Nombre_Horizontal = Nombre_Horizontal.split("\n").join("");
            if(Nombre_Horizontal == "No existe registro"){res.json({process: '0',result: 'No hay registro'});}
            else
            {
                Lista_nombres = Nombre_Horizontal.split(',');
                if(Lista_nombres.length === 1)
                {res.json({process: '1',result: Lista_nombres[0]});}
                else if(Lista_nombres.length === 2)
                {res.json({process: '2',result: Lista_nombres});}
                else
                {res.json({process: '3',result: Lista_nombres});}
            }
            // else if(Nombre_Horizontal == "Las empresas no son iguales"){res.json({process: '2',result: 'No hay registro'});}
            // else{process.stdout.on('end', function(data) {res.json({process: '1',result: Nombre_Horizontal});})}
        });
    }
    else{
        res.json({process: '0',result: 'No hay ID'});
    }
};

// Txt
requestsController.ReporteTxtResponse = (req, res) => {

    let data_1 = req.body.Data.empresa;
    let data_2 = req.body.Data.anio;
    let data_3 = req.body.Data.mes;
    let data_4 = req.body.Data.groups;
    
    if(data_1 != "" && data_2 != "" && data_3 != "" && data_4 != null){
        const process = spawn('python',["./src/python/TXTSS.py",data_1,data_2,data_3,data_4]);
        process.stderr.on("data",(data)=>{
            console.error('stderr:',data.toString());
        })
        process.stdout.on('data', (data) => {
            Nombre_txt = data.toString();
            Nombre_txt = Nombre_txt.split("\r\n").join("");
            Nombre_txt = Nombre_txt.split("\n").join("");
            if(Nombre_txt == "No existe registro"){res.json({process: '0', result: 'No hay Reporte TXTSS'});}
            else{process.stdout.on('end', function(data) {res.json({process: '1', result: Nombre_txt});})}
        });
    }
    else if(data_1 != "" && data_2 != "" && data_3 != "" && data_4 == null){
        data_4 = "Search_group";
        const process = spawn('python',["./src/python/TXTSS.py",data_1,data_2,data_3,data_4]);
        process.stderr.on("data",(data)=>{
            console.error('stderr:',data.toString());
        })
        process.stdout.on('data', (data) => {
            Group_List_ = data.toString();
            Group_List_ = Group_List_.split("\r\n").join("");
            Group_List_ = Group_List_.split("\n").join("");
            if(Group_List_ == "No existe registro"){res.json({process: '0', result: 'No hay Reporte TXTSS'});}
            else{process.stdout.on('end', function(data) {res.json({process: '2', result: Group_List_});})}
        }); 
    }
    else{
        res.json({process: '0',result: 'No hay datos TXTSS'});
    }
}   

// Funcion de leer excel
function leerexcel(Lista_nombre){
    const workBook = XLSX.readFile('./src/database/'+Lista_nombre);
    const workbooksheet = workBook.SheetNames;
    const sheet = workbooksheet[0];
    const dataExcel = XLSX.utils.sheet_to_json(workBook.Sheets[sheet]);
}

// Función descargar Archivo reporte horizontal
requestsController.ReporteHorizontalResponseDocument = (req, res) => {
    // leerexcel(Lista_nombres[0]);
    res.download('./src/database/'+Lista_nombres[0]); 
    setTimeout(() => {fs.unlinkSync('./src/database/'+Lista_nombres[0]);},"100")
}

// Función descargar Archivo reporte horizontal 2
requestsController.ReporteHorizontalResponseDocument2 = (req, res) => {
    // leerexcel(Lista_nombres[1]);
    res.download('./src/database/'+Lista_nombres[1]); 
    setTimeout(() => {fs.unlinkSync('./src/database/'+Lista_nombres[1]);},"100")
}

// Función descargar Archivo reporte horizontal 3
requestsController.ReporteHorizontalResponseDocument3 = (req, res) => {
    // leerexcel(Lista_nombres[2]);
    res.download('./src/database/'+Lista_nombres[2]); 
    setTimeout(() => {fs.unlinkSync('./src/database/'+Lista_nombres[2]);},"100")
}

// Funcion descargar Archivo TxtSS
requestsController.ReporteTxtResponseDocument = (req, res) => {
    res.download('./src/database/'+Nombre_txt); 
    setTimeout(() => {fs.unlinkSync('./src/database/'+Nombre_txt);}, "100")
}

// Funcion descargar plantilla Excel
requestsController.PlantillaExcelProcess = (req, res) => {
    const { body } = req;
    const jsonString = JSON.stringify(body);
    // Condicionales
    if(jsonString != "")
    {
        const process = spawn('python',["./src/python/Plantillas.py",jsonString]);
        process.stderr.on("data",(data)=>{
            console.error('stderr:',data.toString());
        })
        process.stdout.on('data', (data) => {
            Respuesta = data.toString();
            Respuesta = Respuesta.split("\r\n").join("");
            Respuesta = Respuesta.split("\n").join("");
            console.log(Respuesta);
            if(Respuesta == "No existe registro"){
                res.json({ message:'Solicitud recibida correctamente, sin registro'});
            }
            else{
                process.stdout.on('end', function(data) {
                    setTimeout(() => {fs.unlinkSync(Respuesta);}, "2000")
                    res.json({ message:'Solicitud recibida correctamente'});
                })
            }
        });
    }
}

// Exportar módulo
module.exports = requestsController;