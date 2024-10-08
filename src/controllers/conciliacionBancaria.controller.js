const { dirname,join } = require('path');
const spawn = require("child_process").spawn;
const fs = require('fs');
const conciliacionBancariaController = {};

const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

conciliacionBancariaController.calcular = (req, res) => {
    const { app, report, recordId, fieldName, fileName, valueMin, iterations } = req.body;
    if (!app || !report || !recordId || !fieldName || !fileName || valueMin === undefined || valueMin === null || !iterations) {
        res.status(400).json({ message: 'Las keys app, report, recordId, fieldName, fileName, valueMin, iterations en el cuerpo de la solicitud son obligatorias'});
        console.log('El cuerpo de la solicitud esta incompleto');
        return;
    }

    const jsonString = JSON.stringify(req.body);
    const __dirname = dirname(require.main.filename);

    res.status(200).json({ message: 'Solicitud recibida correctamente' });

    delay(5);

    const process = spawn('python',[join(__dirname,'/python/ConciliacionBancaria/conciliacionBancaria.py'), jsonString]);

    process.stderr.on("data", (data) => {
        console.error('stderr:',data.toString());
    });

    process.stdout.on('data', (data) => {
        file_path = data.toString();
        file_path = file_path.split("\r\n").join("");
        file_path = file_path.split("\n").join("");
        
        if (file_path.includes("Error")) {
            // res.status(400).json({ message: file_path });
            console.log(file_path);
        
        } else {
            process.stdout.on('end', function(data) {
                fs.unlinkSync(file_path);
                console.log(file_path);
                // res.status(200).json({ message: 'Solicitud recibida correctamente' });
            })
        }
    });
};

module.exports = conciliacionBancariaController;