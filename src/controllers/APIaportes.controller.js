const ApiAportesController = {}

const NombreUsuario = "901023218PruebasNOM";
const Password = "n3NR61M_Ngs7zd";
const Aplicacion=  "E2271FA7-0FCA-4293-BF6D-53414286FDB0";

ApiAportesController.certificadoAportes= (req, res) => {
    const url = "https://marketplacepruebas.aportesenlinea.com/Transversales.Servicios.Fachada/api/ControlAcceso/Autenticar";
    const data= {
        "data":[{
            "NombreUsuario":NombreUsuario,
            "Password":Password,
            "Aplicacion":Aplicacion
        }]
    }
    const headers= {
        "Content-Type":"application/json",
        "Anon":"Mareigua.Fanaia"
    }
    //Retorna token para acceso a aportes en linea
    Consultar(url,data,headers).then(
        (response)=>{
            // console.log("resultado")
            // console.log(response.data)
            const url = "https://aplicacionespruebas.aportesenlinea.com/Reportes.ServicioWeb/Reportes.svc/CertificadoAportes";
            var tDate = req.body.Fechainicial
            var tDate2 = req.body.FechaFinal
            var date = new Date(tDate);
            var unixTimeStamp1 = Math. floor(date. getTime() / 1000);
            var date2 = new Date(tDate2);
            var unixTimeStamp2 = Math. floor(date2. getTime() / 1000);

            const data= {
                "certificadoAportes":{
                    "TipoIdentificacionEmpleado":req.body.TipoDoc,
                    "NumeroIdentificacionEmpleado":req.body.Doc,
                    "PeriodoDesde": unixTimeStamp1,
                    "PeriodoHasta": unixTimeStamp2,
                    "FormatoReporte": 1,
                    "LlaveApertura": "eb50-43ba-96f2f2"
                }
            }
            const headers= {
                "Content-Type":"application/json",
                "Anon":"Mareigua.Fanaia",
                "Token":response.data
            // 'Content-Type': 'application/x-www-form-urlencoded',
            }
            Consultar(url,data,headers).then( (response) => {
                console.log(response)
                res.json(response)
                //Usage example:
                // let decodificado = atob(response.Reporte);
                // res.json(decodificado)
                // console.log(decodificado);
            })
        }

    )
}
ApiAportesController.consultaIndividual= (req, res) => {
    const url = "https://marketplacepruebas.aportesenlinea.com/Transversales.Servicios.Fachada/api/ControlAcceso/Autenticar";
    const data= {
        "data":[{
            "NombreUsuario":NombreUsuario,
            "Password":Password,
            "Aplicacion":Aplicacion
        }]
    }
    const headers= {
        "Content-Type":"application/json",
        "Anon":"Mareigua.Fanaia"
    }
    //Retorna token para acceso a aportes en linea
    Consultar(url,data,headers).then(
        (response)=>{
            // console.log("resultado")
            // console.log(response.data)
            const url = "https://marketplacepruebas.aportesenlinea.com/Transversales.Servicios.Fachada/api/Persona/ConsultarPersonaEnBaseDatosReferencia";

            const data= {
                "data":[{
                    "TipoDocumento":req.body.TipoDoc,
                    "NumeroDocumento":req.body.Doc,
                }]
            }
            const headers= {
                "Content-Type":"application/json",
                "Anon":"Mareigua.Fanaia",
                "Token":response.data
            // 'Content-Type': 'application/x-www-form-urlencoded',
            }
            Consultar(url,data,headers).then( (response) => {
                // console.log(response)
                res.json(response)
                //Usage example:
                // let decodificado = atob(response.Reporte);
                // res.json(decodificado)
                // console.log(decodificado);
            })
        }

    )
}

ApiAportesController.Validacion_Cargue= (req, res) => {
    const url = "https://marketplacepruebas.aportesenlinea.com/Transversales.Servicios.Fachada/api/ControlAcceso/Autenticar";
    const data= {
        "data":[{
            "NombreUsuario":NombreUsuario,
            "Password":Password,
            "Aplicacion":Aplicacion
        }]
    }
    const headers= {
        "Content-Type":"application/json",
        "Anon":"Mareigua.Fanaia"
    }
    //Retorna token para acceso a aportes en linea
    Consultar(url,data,headers).then(
        (response)=>{
            // console.log("resultado")
            // console.log(response.data)
            const url = "https://marketplacepruebas.aportesenlinea.com/Fanaia.Servicios.Fachada/api/TransmisorPlanillaIntegrada/recepcionSolicitudPlanillaIntegrada";

            const data= re.body
            const headers= {
                "Content-Type":"application/json",
                "Anon":"Mareigua.Fanaia",
                "Token":response.data
            // 'Content-Type': 'application/x-www-form-urlencoded',
            }
            Consultar(url,data,headers).then( (response) => {
                // console.log(response)
                res.json(response)
                //Usage example:
                // let decodificado = atob(response.Reporte);
                // res.json(decodificado)
                // console.log(decodificado);
            })
        }

    )
}
ApiAportesController.validar= (req, res) => {
    const url = "https://marketplacepruebas.aportesenlinea.com/Transversales.Servicios.Fachada/api/ControlAcceso/Autenticar";
    const data= {
        "data":[{
            "NombreUsuario":NombreUsuario,
            "Password":Password,
            "Aplicacion":Aplicacion
        }]
    }
    const headers= {
        "Content-Type":"application/json",
        "Anon":"Mareigua.Fanaia"
    }
    //Retorna token para acceso a aportes en linea
    Consultar(url,data,headers).then(
        (response)=>{
            // console.log("resultado")
            // console.log(response.data)
            const url = "https://marketplacepruebas.aportesenlinea.com/Fanaia.Servicios.Fachada/api/TransmisorPlanillaIntegrada/consultarEstadoSolicitud";

            const data= re.body
            const headers= {
                "Content-Type":"application/json",
                "Anon":"Mareigua.Fanaia",
                "Token":response.data
            // 'Content-Type': 'application/x-www-form-urlencoded',
            }
            Consultar(url,data,headers).then( (response) => {
                // console.log(response)
                res.json(response)
                //Usage example:
                // let decodificado = atob(response.Reporte);
                // res.json(decodificado)
                // console.log(decodificado);
            })
        }

    )
}

async function Consultar(url = "", data = {}, headers_ = {}) {
    const response = await fetch(url, {
    method: "POST", // *GET, POST, PUT, DELETE, etc.
    mode: "cors", // no-cors, *cors, same-origin
    cache: "no-cache", // *default, no-cache, reload, force-cache, only-if-cached
    credentials: "same-origin", // include, *same-origin, omit
    headers: headers_,
    redirect: "follow", // manual, *follow, error
    referrerPolicy: "no-referrer", // no-referrer, *no-referrer-when-downgrade, origin, origin-when-cross-origin, same-origin, strict-origin, strict-origin-when-cross-origin, unsafe-url
    body: JSON.stringify(data), // body data type must match "Content-Type" header
    });
    // console.log(response)
    return response.json(); // parses JSON response into native JavaScript objects
  }

module.exports = ApiAportesController;
