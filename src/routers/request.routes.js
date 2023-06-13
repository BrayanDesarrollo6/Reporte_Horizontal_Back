// Objeto que almacena las rutas (ROUTER)
const { Router } = require('express');
const router = Router();
const requestController = require('../controllers/requests.controller.js')

// Ruta 1 - Reporte Horizontal
router.post('/procesar', requestController.ReporteHorizontalResponse)
router.get('/procesar', requestController.ReporteHorizontalResponseDocument)
router.get('/procesar2', requestController.ReporteHorizontalResponseDocument2)
router.get('/procesar3', requestController.ReporteHorizontalResponseDocument3)

// Ruta 2 - Reporte TXT
router.post('/procesarTXTSS', requestController.ReporteTxtResponse)
router.get('/procesarTXTSS', requestController.ReporteTxtResponseDocument)

// Ruta 3 - Plantillas Excel
router.post('/plantilla', requestController.PlantillaExcelProcess)

module.exports = router;