// Almacenar objeto de funciones de enrutamiento - pertenece a router
const especialesController = {}
const { log } = require('console');
// Variables compartidas
const fs = require('fs');
const spawn = require("child_process").spawn;
const utf8 = require('utf8');
// Variables Reporte Nomina, LQ y ReLQ
const XLSX = require('xlsx');
var Nombre_Horizontal = "";
var Lista_nombres = [];
// Constantes Archivo TxtSS
var Nombre_txt = "";
// Directories Horizontal
PythonReporteHorizontal = "./src/python/ReporteEspecialDHL.py";
DescargaReporteHorizontal = './src/database/';


// Reporte Horizontal
especialesController.reporteNominaDHL = (req, res) => {

    const ID_received = req.body.Data;
    let process;
    let estado = 0;
    let data_0 = ID_received.idproceso;
    let data_1 = ID_received.idperiodo;
    let data_2 = ID_received.idperiodo2;
    let data_3 = ID_received.idperiodo3;
    if(data_1 !== undefined && data_1 !== null && data_2 === undefined || data_2 === null || data_2 === "" && data_3 === undefined)
    {
        console.log(ID_received)
        estado = 1;
        process = spawn('python',[PythonReporteHorizontal,estado,data_0,data_1]);
    }
    else if(data_1 !== undefined && data_1 !== null && data_2 !== undefined && data_2 !== null && data_3 === undefined || data_3 === null || data_3 === "")
    {
        estado = 2;
        process = spawn('python',[PythonReporteHorizontal,estado,data_0,data_1,data_2]);
    }   
    else if(data_1 !== undefined && data_1 !== null && data_2 !== undefined && data_2 !== null && data_3 !== undefined && data_3 !== null)
    {
        estado = 3;
        process = spawn('python',[PythonReporteHorizontal,estado,data_0,data_1,data_2,data_3]);
    }
    if(estado != 0){
        process.stderr.on("data",(data)=>{
            console.error('stderr:',data.toString());
        })
        process.stdout.on('data', (data) => {
            console.log(data.toString())
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
        });
    }
    else{
        res.json({process: '0',result: 'No hay ID'});
    }
};

// Función descargar Archivo reporte horizontal
especialesController.reporteNominaResponseDocument = (req, res) => {
    res.download(DescargaReporteHorizontal+Lista_nombres[0]); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteHorizontal+Lista_nombres[0]);},"100")
}
// Función descargar Archivo reporte horizontal
especialesController.reporteNominaResponseDocument2 = (req, res) => {
    res.download(DescargaReporteHorizontal+Lista_nombres[1]); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteHorizontal+Lista_nombres[1]);},"100")
}
// Función descargar Archivo reporte horizontal
especialesController.reporteNominaResponseDocument3 = (req, res) => {
    res.download(DescargaReporteHorizontal+Lista_nombres[2]); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteHorizontal+Lista_nombres[2]);},"100")
}
// Exportar módulo
module.exports = especialesController;