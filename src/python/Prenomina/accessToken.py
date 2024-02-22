import mysql.connector
import time
import requests
import os
import json
class funcionesGenerales():
    def conectarBD(self):
        try:
            conexion = mysql.connector.connect(
                host = "201.184.98.75",
                user = "desarrollo3",
                password = "5cTmZk25f",
                database = "ApiManagePer",
                port="3306"
            )
            return conexion
        except mysql.connector.Error as err:
            return None
    def get_registro(self,tabla,campos, condiciones):
        conexion = self.conectarBD()
        try:
            cursor = conexion.cursor(dictionary=True)
            # Crear la consulta SQL
            consulta = f"SELECT {', '.join(campos)} FROM {tabla}"
            if condiciones:
                consulta += f" WHERE {' AND '.join(condiciones)}"
            # Ejecutar la consulta
            cursor.execute(consulta)
            # Obtener resultados
            resultados = cursor.fetchall()
            return resultados
        except mysql.connector.Error as err:
            print(f"Error: {err}")
            return None
        finally:
            if conexion.is_connected():
                cursor.close()
                conexion.close()
    def get_OneOrder(self,tabla,campos,order, condiciones):
        conexion = self.conectarBD()
        if conexion:
            try:
                cursor = conexion.cursor(dictionary=True)
                # Crear la consulta SQL
                # `SELECT ${select} FROM ${tabla} WHERE ? ORDER BY ${order} DESC LIMIT 1`
                consulta = f"SELECT {', '.join(campos)} FROM {tabla}"
                if condiciones:
                    consulta += f" WHERE {' AND '.join(condiciones)}"
                if order:
                    consulta += f" ORDER BY {order} DESC LIMIT 1"
                # Ejecutar la consulta
                cursor.execute(consulta)
                # Obtener resultados
                resultados = cursor.fetchall()
                return resultados
            except mysql.connector.Error as err:
                print(f"Error: {err}")
                return None
            finally:
                if conexion.is_connected():
                    cursor.close()
                    conexion.close()
    def access_token_zoho(self,id):
        resultados_ = self.get_registro("refreshToken_zoho",["id","refresh"],[f"id ={id}"])
        # https://accounts.zoho.com/oauth/v2/token?client_id=1000.IR8Z2X49OOIDKTKJ6V061D695UZSAA&grant_type=refresh_token&client_secret=304f1af84b8d1cf7e9e13906ef481a7c3864c7c4f0&refresh_token=1000.2a42b5c8b8bde22c60272a117dda6fdf.0043c9f3e25b99a7bc9ccddc18bb677d
        # refresh_ = resultados_[0]['refresh']
        # client_id = '1000.IR8Z2X49OOIDKTKJ6V061D695UZSAA'
        # client_secret = '304f1af84b8d1cf7e9e13906ef481a7c3864c7c4f0'
        # url = 'https://accounts.zoho.com/oauth/v2/token?refresh_token=' + refresh_ + '&client_id='+client_id+'&client_secret='+client_secret+'&grant_type=refresh_token'
        # cabeceras = {"Content-Type": "application/json", "Access-Control-Allow-Origin": "*"} 
        # auth_data = {"answer": "42" }
        # resp = requests.post(url, data=auth_data,headers=cabeceras)
        url = 'https://accounts.zoho.com/oauth/v2/token'
        data = {
            'grant_type': 'refresh_token',
            'client_id': '1000.IR8Z2X49OOIDKTKJ6V061D695UZSAA',
            'client_secret': '304f1af84b8d1cf7e9e13906ef481a7c3864c7c4f0',
            'refresh_token':resultados_[0]['refresh']
        }
        resp = requests.post(url, data=data)
        posts = resp.json()
        if resp.status_code == 200:
            inicio_ = int(time.time() * 1000)
            milisengundos = posts['expires_in'] * 1000;
            end_ =  inicio_ + milisengundos; 
            columnas = ["access_token","api_domain","token_type","expires_in","created_at","expired_at","refresh_token_id"]
            valores = [posts['access_token'],posts['api_domain'],posts['token_type'],posts['expires_in'],inicio_,end_,id]
            
            self.add_register("tokens", columnas, valores)
            return posts['access_token']
        
        return None
    def add_register(self,tabla,columnas,valores):
    
        conexion = self.conectarBD()
        if conexion:
            try:
                cursor = conexion.cursor(dictionary=True)
                # Crear la consulta SQL
                consulta = f"INSERT INTO {tabla} ({', '.join(columnas)}) VALUES ({', '.join(['%s']*len(valores))})"

                # Ejecutar la consulta
                cursor.execute(consulta,valores)
                # Hacer commit para aplicar los cambios en la base de datos
                conexion.commit()
            except mysql.connector.Error as err:
                print(f"Error: {err}")
                return None
            finally:
                if conexion.is_connected():
                    cursor.close()
                    conexion.close()

    def updatedata(self,file_path,record_id):

        resultados_ = self.get_registro("refreshToken_zoho",["id","refresh"],["usuario = 'desarrollo3@hq5.com.co'"])
        resultados = self.get_OneOrder("tokens",["access_token","expired_at"],"id",[f"refresh_token_id = '{resultados_[0]['id']}'"])
        # Obtener el tiempo actual en milisegundos
        inicio_ = int(time.time() * 1000)
        token= resultados[0]['access_token']
        if (resultados[0]['expired_at'] < inicio_):
            token = self.access_token_zoho(resultados_[0]['id'])
        url = f"https://creator.zoho.com/api/v2/hq5colombia/compensacionhq5/report/Generar_pre_nomina_Report/{record_id}/Adjunto1/upload"
        headers = {
            "Authorization": f"Zoho-oauthtoken {token}"
        }
        with open(file_path, "rb") as file:
            files = {"file": file}
            response = requests.post(url, headers=headers, files=files)
            if(response.status_code == 200):
                file.close()
                time.sleep(5)
                os.remove(file_path)   
                #hacer registro para el correo
                # Realizar consulta
                Datos_ = {"data": {"Generar_pre_nomina":record_id,"Proceso":"Prenomina"}}
                Datos_ = json.dumps(Datos_)
                url_ = "https://creator.zoho.com/api/v2/hq5colombia/compensacionhq5/form/templatesVarios"
                header = {"Authorization":"Zoho-oauthtoken "+token, "Access-Control-Allow-Origin": "*"} 
                r = requests.post(url_,data=Datos_,headers=header)
                print(r.json())
    def getAIU(self,nameCliente):

        resultados_ = self.get_registro("refreshToken_zoho",["id","refresh"],["usuario = 'desarrollo3@hq5.com.co'"])
        resultados = self.get_OneOrder("tokens",["access_token","expired_at"],"id",[f"refresh_token_id = '{resultados_[0]['id']}'"])
        # Obtener el tiempo actual en milisegundos
        inicio_ = int(time.time() * 1000)
        token= resultados[0]['access_token']
        if (resultados[0]['expired_at'] < inicio_):
            token = self.access_token_zoho(resultados_[0]['id'])
        url = f"https://creator.zoho.com/api/v2/hq5colombia/hq5/report/Ver_Cliente?EMPRESA_APLICAR_CONVOCATORIA={nameCliente}"
        headers = {
            "Authorization": f"Zoho-oauthtoken {token}"
        }
        response = requests.get(url,headers=headers)
        resp = response.json()
        if(response.status_code == 200):
            res = resp['data']
            aui_ = float(res[0]['aiu_cli']) / 100
            return aui_
        else: return 0.07
