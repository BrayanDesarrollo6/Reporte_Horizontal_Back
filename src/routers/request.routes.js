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

// Ruta 4 - Reporte liquidaciones
router.post('/getEmpresas', requestController.obtenerEmpresas)
router.get('/getEmpresas', requestController.enviarEmpresas)
router.post('/procesarlq', requestController.ReporteLiquidacionResponse)
router.get('/procesarlq', requestController.ReporteLiquidacionResponseDocument)

// Ruta 5 - Reporte Reliquidaciones
router.post('/getEmpresasrelq', requestController.obtenerEmpresas)
router.get('/getEmpresasrelq', requestController.enviarEmpresasrelq)
router.post('/procesarrelq', requestController.ReporteReLiquidacionResponse)
router.get('/procesarrelq', requestController.ReporteReLiquidacionResponseDocument)

//RUTA 6 - API sin frontexterno, directo zoho
router.post('/api/v1/formatoOrdenIngreso', requestController.formatoOrdenIngreso)

module.exports = router;