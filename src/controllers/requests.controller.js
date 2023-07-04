// Almacenar objeto de funciones de enrutamiento - pertenece a router
const requestsController = {}
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
PythonReporteHorizontal = "./src/python/ReporteHorizontal.py";
DescargaReporteHorizontal = './src/database/';
// Directories Txt
PythonReporteTxt = "./src/python/TXTSS.py";
DescargaReporteTxt = './src/database/';
// Directories Plantillas
PythonPlantillasHorizontal = "./src/python/Plantillas.py";
PythonObtenerEmpresas = "./src/python/obtenerEmpresas.py";
// Directories Json
DescargaJsonLiquidaciones = "./src/database/VariablesEntornoLQ.json";
DescargaJsonReLiquidaciones = "./src/database/VariablesEntornoReLQ.json";
// Directories Liquidaciones
PythonReporteLiquidaciones = "./src/python/reporteLiquidaciones.py";
DescargaReporteLiquidaciones = './src/database/';
// Directories ReLiquidaciones
PythonReporteReLiquidaciones = "./src/python/reporteReliquidaciones.py";
DescargaReporteReLiquidaciones = './src/database/';

// Reporte Horizontal
requestsController.ReporteHorizontalResponse = (req, res) => {

    const ID_received = req.body.Data;
    let process;
    let estado = 0;
    let data_0 = ID_received.idproceso;
    let data_1 = ID_received.idperiodo;
    let data_2 = ID_received.idperiodo2;
    let data_3 = ID_received.idperiodo3;

    if(data_1 !== undefined && data_1 !== null && data_2 === undefined || data_2 === null || data_2 === "" && data_3 === undefined)
    {
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
    res.download(DescargaReporteHorizontal+Lista_nombres[0]); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteHorizontal+Lista_nombres[0]);},"100")
}

// Función descargar Archivo reporte horizontal 2
requestsController.ReporteHorizontalResponseDocument2 = (req, res) => {
    // leerexcel(Lista_nombres[1]);
    res.download(DescargaReporteHorizontal+Lista_nombres[1]); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteHorizontal+Lista_nombres[1]);},"100")
}

// Función descargar Archivo reporte horizontal 3
requestsController.ReporteHorizontalResponseDocument3 = (req, res) => {
    // leerexcel(Lista_nombres[2]);
    res.download(DescargaReporteHorizontal+Lista_nombres[2]); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteHorizontal+Lista_nombres[2]);},"100")
}

// Txt
requestsController.ReporteTxtResponse = (req, res) => {

    let data_1 = req.body.Data.empresa;
    let data_2 = req.body.Data.anio;
    let data_3 = req.body.Data.mes;
    let data_4 = req.body.Data.groups;
    
    if(data_1 != "" && data_2 != "" && data_3 != "" && data_4 != null){
        const process = spawn('python',[PythonReporteTxt,data_1,data_2,data_3,data_4]);
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
        const process = spawn('python',[PythonReporteTxt,data_1,data_2,data_3,data_4]);
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

// Funcion descargar Archivo TxtSS
requestsController.ReporteTxtResponseDocument = (req, res) => {
    res.download(DescargaReporteTxt+Nombre_txt); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteTxt+Nombre_txt);}, "100")
}

// Funcion descargar plantilla Excel
requestsController.PlantillaExcelProcess = (req, res) => {
    const { body } = req;
    const jsonString = JSON.stringify(body);
    // Condicionales
    if(jsonString != "")
    {
        const process = spawn('python',[PythonPlantillasHorizontal,jsonString]);
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

// ObtenerEmpresas
requestsController.obtenerEmpresas = (req, res) => {
    const ID_received = req.body.Data;
    // Obtener IDPeriodos
    let process;
    let estado = ID_received;
    process = spawn('python',[PythonObtenerEmpresas,estado]);
    process.stderr.on("data",(data)=>{
        console.error('stderr:',data.toString());
    })
    process.stdout.on('data', (data) => {
        Respuesta = data.toString();
        Respuesta = Respuesta.split("\r\n").join("");
        Respuesta = Respuesta.split("\n").join("");
        console.log(Respuesta);
        if(Respuesta == "No existe registro"){res.json({process: '0', result: 'No hay Empresas'});}
        else{process.stdout.on('end', function(data) {res.json({process: '1', result: Respuesta});})}
    });
};

// Funcion enviar Archivo Json Empresas
requestsController.enviarEmpresas = (req, res) => {
    fs.readFile(DescargaJsonLiquidaciones, "utf8", (err, jsonString) => {
    if (err) {
        console.log("File read failed:", err);
        return;
    }
    console.log("Enviando datos de consulta json");
    res.json(JSON.parse(jsonString));
    });
}

// Funcion enviar Archivo Json Empresas
requestsController.enviarEmpresasrelq = (req, res) => {
    fs.readFile(DescargaJsonReLiquidaciones, "utf8", (err, jsonString) => {
    if (err) {
        console.log("File read failed:", err);
        return;
    }
    console.log("Enviando datos de consulta json");
    res.json(JSON.parse(jsonString));
    });
}

// Reporte Liquidaciones
requestsController.ReporteLiquidacionResponse = (req, res) => {
    let data_1 = req.body.Data.empresa;
    let data_2 = req.body.Data.estados;
    let data_3 = req.body.Data.anio;
    let data_4 = req.body.Data.mes;
    const process = spawn('python',[PythonReporteLiquidaciones,data_1,data_2,data_3,data_4]);
    process.stderr.on("data",(data)=>{
        console.error('stderr:',data.toString());
    })
    process.stdout.on('data', (data) => {
        Nombre_Horizontal = data.toString();
        console.log(Nombre_Horizontal);
        Nombre_Horizontal = Nombre_Horizontal.split("\r\n").join("");
        Nombre_Horizontal = Nombre_Horizontal.split("\n").join("");
        if(Nombre_Horizontal == "No existe registro"){res.json({process: '0', result: 'No hay registro'});}
        else{process.stdout.on('end', function(data) {res.json({process: '1', result: Nombre_Horizontal});})}
    });
}  

// Funcion descargar Archivo liquidaciones
requestsController.ReporteLiquidacionResponseDocument = (req, res) => {
    res.download(DescargaReporteLiquidaciones+Nombre_Horizontal); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteLiquidaciones+Nombre_Horizontal);}, "100")
}

// Reporte ReLiquidaciones
requestsController.ReporteReLiquidacionResponse = (req, res) => {
    let data_1 = req.body.Data.empresa;
    let data_2 = req.body.Data.estados;
    let data_3 = req.body.Data.anio;
    let data_4 = req.body.Data.mes;
    const process = spawn('python',[PythonReporteReLiquidaciones,data_1,data_2,data_3,data_4]);
    process.stderr.on("data",(data)=>{
        console.error('stderr:',data.toString());
    })
    process.stdout.on('data', (data) => {
        Nombre_Horizontal = data.toString();
        Nombre_Horizontal = Nombre_Horizontal.split("\r\n").join("");
        Nombre_Horizontal = Nombre_Horizontal.split("\n").join("");
        console.log(Nombre_Horizontal);
        if(Nombre_Horizontal == "No existe registro"){res.json({process: '0', result: 'No hay registro'});}
        else{process.stdout.on('end', function(data) {res.json({process: '1', result: Nombre_Horizontal});})}
    });
}  

// Funcion descargar Archivo Reliquidaciones
requestsController.ReporteReLiquidacionResponseDocument = (req, res) => {
    res.download(DescargaReporteReLiquidaciones+Nombre_Horizontal); 
    setTimeout(() => {fs.unlinkSync(DescargaReporteReLiquidaciones+Nombre_Horizontal);}, "100")
}

// Exportar módulo
module.exports = requestsController;