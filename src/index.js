// Traer el modulo app
const app = require('./app');

// Respuesta para conocer que el puerto se encuentra corriendo
app.listen(app.get('port'), ()=>{
    console.log("Servidor en puerto",app.get('port'));
})