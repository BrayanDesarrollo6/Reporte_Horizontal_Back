// Almacenar objeto de funciones de enrutamiento - pertenece a router
const prenominaController = {}
const { dirname,join } = require('path');
// Variables compartidas
const spawn = require("child_process").spawn;

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