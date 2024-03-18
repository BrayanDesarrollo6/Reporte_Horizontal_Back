const facturacionExamenesController = {}
const { dirname,join } = require('path');
const spawn = require("child_process").spawn;
const fs = require('fs');

facturacionExamenesController.generarFormato = (req, res) => {
    
    const __dirname = dirname(require.main.filename);
    const { body } = req;
    const jsonString = JSON.stringify(body);
    let file_name = "";
    const process = spawn('python',[join(__dirname,'/python/reporteFacturacionExamenes.py'),jsonString]);
    process.stderr.on("data",(data)=>{
        console.error('stderr:',data.toString());
    })
    process.stdout.on('data', (data) => {
        file_name = data.toString();
        file_name = file_name.split("\r\n").join("");
        file_name = file_name.split("\n").join("");
        if(file_name.includes("Error")){
            res.json({result: file_name});
        }
        else{
            process.stdout.on('end', function(data) {
                fs.unlinkSync(file_name);
                res.json({ message:'Solicitud recibida correctamente'});
            })
        }
    });
};

// Exportar módulo
module.exports = facturacionExamenesController;