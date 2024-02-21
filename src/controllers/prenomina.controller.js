// Almacenar objeto de funciones de enrutamiento - pertenece a router
const prenominaController = {}
const { dirname,join } = require('path');
const { fileURLToPath } = require('url');
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
prenominaController.generarPrenomina = (req, res) => {
    
    const __dirname = dirname(require.main.filename);
    const { body } = req;
    const jsonString = JSON.stringify(body);
    var Nombre_Horizontal = "";
    const process = spawn('python',[join(__dirname,'/python/Prenomina/prenomina.py'),jsonString]);
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
};

// Exportar módulo
module.exports = prenominaController;