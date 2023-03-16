// Objeto que almacena las rutas (ROUTER)
const { Router } = require('express');
const router = Router();
const requestController = require('../controllers/requests.controller.js')
//API APORTES EN LINEA
const ApiAportesController = require('../controllers/APIaportes.controller.js')

// Ruta 1 - Reporte Horizontal
router.post('/procesar', requestController.ReporteHorizontalResponse)
router.get('/procesar', requestController.ReporteHorizontalResponseDocument)
router.get('/procesar2', requestController.ReporteHorizontalResponseDocument2)
router.get('/procesar3', requestController.ReporteHorizontalResponseDocument3)

// Ruta 2 - Reporte TXT
router.post('/procesarTXTSS', requestController.ReporteTxtResponse)
router.get('/procesarTXTSS', requestController.ReporteTxtResponseDocument)

// Ruta 3 - API Aportes
router.post('/certificadoAportes', ApiAportesController.certificadoAportes)
router.post('/consultaIndividual', ApiAportesController.consultaIndividual)
router.post('/validacionCargue', ApiAportesController.Validacion_Cargue)
router.post('/validar', ApiAportesController.validar)
// router.post('/consultaMasiva', ApiAportesController.certificadoAportes)
module.exports = router;