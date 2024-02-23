import pandas as pd
import xlsxwriter
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
from Directories.Directory import DirectoryResumenDHL

class resumen():
    # reemplazar acentos
    def normalize(self,s):
        replacements = (
            ("á", "a"),("é", "e"),("í", "i"),("ó", "o"),("ú", "u"),
            ("Á", "A"),("É", "E"),("Í", "I"),("Ó", "O"),("Ú", "U")
        )
        for a, b in replacements:
            s = s.replace(a, b).replace(a.upper(), b.upper())
        return s
    # reemplazar caracteres
    def replacement(self,name_company):
        name_company = name_company.replace(".", "")
        name_company = name_company.replace("-", "_")
        name_company = name_company.replace("–", "_")
        name_company = name_company.replace("—", "_")
        return name_company

    def generarrow(self,dfEmpleado, Horizontal):
        FilaAgregar = {}
        FilaAgregar["Empresa"] = dfEmpleado['Empresa a la que se le factura'].values[0]
        FilaAgregar["Periodo"] = dfEmpleado['Mes'].values[0]
        FilaAgregar["Cedula jefe"] = ""
        FilaAgregar["AT"] = dfEmpleado['At (actividad)'].values[0]
        FilaAgregar["CU"] = dfEmpleado['Cu (customer -cliente)'].values[0]
        FilaAgregar["N_Deudor"] = dfEmpleado['Cu name'].values[0]
        FilaAgregar["Cost Center"] = dfEmpleado['Cost center'].values[0]
        FilaAgregar["CC Name"] = dfEmpleado['Cost center name'].values[0]
        FilaAgregar["Cedula"] = dfEmpleado['Cedula'].values[0]
        FilaAgregar["NombreAsociado"] = dfEmpleado['Nombre empleado'].values[0]
        FilaAgregar["Cargo"] = dfEmpleado['Cargo'].values[0]
        FilaAgregar["FechaIngreso"] = dfEmpleado['Fechaingreso'].values[0]
        FilaAgregar["FechaRetiro"] = dfEmpleado['Fecharetiro'].values[0]
        FilaAgregar["Basico"] = dfEmpleado['Salario basico'].values[0]
        # DETALLADO
        FilaAgregar["Dias"] = dfEmpleado['Total días de liquidación (mes o quincena)'].values[0]
        FilaAgregar["Salario"] = dfEmpleado['Valor Salario'].values[0] + dfEmpleado['Reajuste salario'].values[0]
        FilaAgregar["AuxilioTransporte"] = dfEmpleado['Aux.transporte'].values[0] + dfEmpleado['Reajuste aux. transporte'].values[0]
        # EXTRAS
        FilaAgregar["Extras"] = dfEmpleado['Total cantidad HE + Recargos'].values[0]
        FilaAgregar["V_Horas_Extras"] = dfEmpleado['Total valor HE + Recargos'].values[0]
        # INCAPACIDADES
        tipos_incapacidades = [
            "incapacidad enfermedad general (Días 1 y 2)",
            "incapacidad accidente de trabajo",
            "incapacidad enfermedad general 3 Días en adelante hasta el día 90",
            "incapacidad permanente entre día 91 al de 180 Días",
            "incapacidad permanente + de 180 Días",
            "licencia de maternidad",
            "Licencia de Paternidad",
        ]
        diasIncapacidad_ = sum(dfEmpleado[f"Días {tipo}"].values[0] for tipo in tipos_incapacidades)
        valorIncapacidad_ = sum(dfEmpleado[f"Valor {tipo}"].values[0] for tipo in tipos_incapacidades)
        FilaAgregar["Incapacidades"] = diasIncapacidad_
        FilaAgregar["V_Incapacidad"] = valorIncapacidad_
        # VACACIONES
        tipos_vacaciones = [
            "Vacaciones en Disfrute Causadas",
            "Vacaciones en Disfrute Anticipadas",
        ]
        valorVacaciones_ = sum(dfEmpleado[f"Valor {tipo}"].values[0] for tipo in tipos_vacaciones)
        FilaAgregar["V_Vacaciones"] = valorVacaciones_
        # OTROS PAGOS FACTURABLES
        tipos_Permisos = [
            "Valor Hospitalizacion",
            "Valor Inasistencia por Alteración del Orden Publico 14708",
            "Valor Licencia de Matrimonio",
            "Valor Licencia de Luto 5 Días hab (muerte de un familiar *1er grado de consanguinidad 5)",
            "Valor Licencia Remunerada - covid casos vulnerables",
            "Valor Licencia Remunerado (mayor a 2 Días) Aprobación RH",
            "Valor Libre Jurado Votaciones",
            "Valor Calamidad Domestica (de 1 o 2 Días)",
            "Valor Día Compensación por desempeño",
            "Valor Día Familiar",
            "Valor Permiso Justificado - Covid prevención aislamiento",
            "Valor Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)",
        ]
        valorPermisos_ = sum(dfEmpleado[f"{tipo}"].values[0] for tipo in tipos_Permisos)
        tipos_bonos = [
            "Indemnización, Bono por Retiro o Suma transaccional",
            "Auxilio desplazamiento (parametrizado en el sistema por días laborados depende lugar trabajo y lugar residencia)",
            "Gastos de transporte fijo (monto mensual asignado para desempeño de sus funciones ej. Comerciales, mantenimiento, gerentes, area seguridad etc)",
            "Gastos de transporte ocasional (reportado por la operación quincenalmente)",
            "Bonificacion no constitutiva de salario  BNCS",
            "Bonificacion salarial",
        ]
        valorBonos_ = sum(dfEmpleado[f"{tipo}"].values[0] for tipo in tipos_bonos)
        # ASIGNAR
        FilaAgregar["OtrosPagosFacturables"] = valorPermisos_+ valorBonos_
        # DEDUCCIONES
        descuentos_= [
            "Grupo # 3 Valor total ausencias injustificadas, sanciones, dominical",
            "Dcto my v/r pagado salario/Saldo en rojo",
            "Saldo reconocido por el cliente",
            "DEDUCCIONES VARIAS - NC",
            "Deduccion Casino"
        ]
        valorDescuentos_ = sum(dfEmpleado[f"{tipo}"].values[0] for tipo in descuentos_)
        FilaAgregar["DeduccionesEmpledos"] = valorDescuentos_
        # PRESTACIONES
        FilaAgregar["PrestacionesSociales"] = dfEmpleado["Valor prestaciones sociales"].values[0] 
        FilaAgregar["SeguridadSocialyParafiscales"] = dfEmpleado["Valor parafiscales"].values[0]  + dfEmpleado["Total Seguridad Social"].values[0]
        FilaAgregar["Admon"] = dfEmpleado["Administración temporales (el % que tenga cada temporal)"].values[0] 
        # TOTAL PRENOMINA
        valores_ = [
            "Salario",
            "AuxilioTransporte",
            "V_Horas_Extras",
            "V_Incapacidad",
            "V_Vacaciones",
            "OtrosPagosFacturables",
            "PrestacionesSociales",
            "SeguridadSocialyParafiscales",
            "Admon",
        ]
        valorTotal_ = sum(FilaAgregar[valor] for valor in valores_)
        FilaAgregar["TotalPrenomina"] = valorTotal_
        Horizontal = pd.concat([Horizontal,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)
        return Horizontal 
    
    def generarResumen(self,df):    
        Horizontal = pd.DataFrame()
        Contrato  = df['Cedula'].unique().tolist()
        for i in Contrato:
            Valores = df['Cedula'] == i
            ContratoPos = df[Valores]
            Horizontal = self.generarrow(ContratoPos, Horizontal)
        NombreDocumento = "Publicacion " + str(df.iloc[0]['Empresa a la que se le factura']) +"-"+ str(df.iloc[0]['Mes'])
        # Normalizar nombre del documento
        NombreDocumento = self.normalize(NombreDocumento)
        NombreDocumento = self.replacement(NombreDocumento)
        heads = Horizontal.columns.values
        
        wb = Workbook()
        ws = wb.active
        
        for r in dataframe_to_rows(Horizontal, index=False, header=True):
            ws.append(r)

        Horizontal = pd.DataFrame(ws.values)
        writer = pd.ExcelWriter(DirectoryResumenDHL+NombreDocumento+".xlsx", engine='xlsxwriter')
        
        Horizontal.to_excel(writer, sheet_name='RESUMEN',index = False, header = False)
        workbook = writer.book
        worksheet = writer.sheets["RESUMEN"]    
        MaxFilas = len(Horizontal.axes[0])
        Totales = Horizontal.loc[MaxFilas -1]
        # Escribir encabezados en la primera fila
        contador = 0
        for k in heads:
            # Configurar el formato para la primera fila (encabezado)
            header_format = workbook.add_format()
            # header_format.set_pattern(1)
            header_format.set_align('center')
            header_format.set_align('vcenter')
            header_format.set_border(1)
            header_format.set_text_wrap()
            header_format.set_bg_color("6AA7FF")
            worksheet.write_string(0, contador, str(k), header_format)
            contador += 1
        # Ajustar automáticamente el ancho de las columnas según el contenido
        for i, heading in enumerate(heads):
            column_values = Horizontal.iloc[:, i].astype(str)
            max_value = max(column_values, key=len)
            column_width = max(len(str(heading)), len(max_value))
            worksheet.set_column(i, i, column_width + 2)
        contador = 0
        for k in Totales:
            Dato = ""
            if(str(k) != "nan"):
                Dato = str(k)
            worksheet.write_string(MaxFilas, contador,Dato)
            contador += 1
        writer.close()
        return NombreDocumento+".xlsx"