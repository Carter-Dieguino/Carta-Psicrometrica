import math
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from tkinter import ttk
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import griddata
import openpyxl

import math
import numpy as np

class CalculadoraPropiedades:
    def __init__(self):
        self.Ra = 287.055
        pass
    
    def Grados_Kelvin(self, temp):
        return temp + 273.15
    
    def calcular_presion(self, altura):
        p_atm = 101.325 * (math.pow((1 - 2.25577e-5 * altura), 5.2559))
        return p_atm * 1000  

    def calcular_pvs(self, Tbs):
        if Tbs < -100:
            Tbs = -100
        elif Tbs > 200:
            Tbs = 200

        Tbs_K = self.Grados_Kelvin(Tbs)

        if Tbs <= 0:  
            C1 = -5.6745359e3
            C2 = 6.3925247e0
            C3 = -9.6778430e-3
            C4 = 6.2215701e-7
            C5 = 2.0747825e-9
            C6 = -9.4840240e-13
            C7 = 4.1635019e0
            ln_pws = (C1 / Tbs_K) + C2 + C3 * Tbs_K + C4 * Tbs_K ** 2 + C5 * Tbs_K ** 3 + C6 * Tbs_K ** 4 + C7 * math.log(Tbs_K)

        elif 0 < Tbs <= 200:
            C8 = -5.8002206e3
            C9 = 1.3914993e0
            C10 = -4.8640239e-2
            C11 = 4.1764768e-5
            C12 = -1.4452093e-8
            C13 = 6.5459673e0
            ln_pws = (C8 / Tbs_K) + C9 + C10 * Tbs_K + C11 * Tbs_K ** 2 + C12 * Tbs_K ** 3 + C13 * math.log(Tbs_K)

        return math.exp(ln_pws)

    def calcular_pv(self, Hr, pvs2):
        return Hr * pvs2

    def razon_humedad(self, Pv, presionAt):
        return 0.622 * (Pv / (presionAt - Pv))

    def razon_humedad_saturada(self, pvs2, presionAt):
        return 0.622 * (pvs2 / (presionAt - pvs2))

    def grado_saturacion(self, W, Ws):
        return (W / Ws) * 100

    def volumen_especifico(self, Tbs, presionAt, W):
        Tbs_K = self.Grados_Kelvin(Tbs)
        return ((self.Ra * Tbs_K) / presionAt) * (1 + 1.6078 * W) / (1 + W)

    def temperatura_punto_rocio(self, Tbs, Pv):
        Tbs_K = self.Grados_Kelvin(Tbs)
        if Pv <= 0:
            Pv = 0.00001  
        if -60 < Tbs < 0:
            return -60.450 + 7.0322 * math.log(Pv) + 0.3700 * (math.log(Pv)) ** 2
        elif 0 < Tbs < 70:
            return -35.957 - 1.8726 * math.log(Pv) + 1.1689 * (math.log(Pv)) ** 2
        return None

    def entalpia(self, Tbs, W):
        return 1.006 * Tbs + W * (2501 + 1.805 * Tbs)

    def bulbo_humedo(self, presionAt, Tbs, W, iter=20):
        Tpr = Tbs - 1
        x0 = np.array(Tpr)
        tolerancia = 0.00001

        for i in range(iter):
            X = x0 + 273.15
            pvs2 = self.calcular_pvs(X - 273.15)

            if pvs2 <= 0:
                pvs2 = 0.00001  
            
            Ws_2 = 0.62198 * (pvs2 / (presionAt - pvs2))
            fx_tbh = (((2501 - 2.381 * x0) * Ws_2 - 1.006 * (Tbs - x0)) / (2501 + 1.805 * Tbs - 4.186 * x0)) - W

            fx_tbh_d = (((2501 - 2.381 * x0) * Ws_2 + Ws_2 * (-2.381) + 1.006) / (2501 + 1.805 * Tbs - 4.186 * x0))

            Tbh = x0 - (fx_tbh / fx_tbh_d)

            if x0 == 0:
                x0 = 1e-14
                
            error = (Tbh - x0) / x0

            x0 = Tbh

            if np.all(np.abs(error) < tolerancia):
                break
        return Tbh


class ManejoDatos:
    def cargar_archivo(self, ruta_archivo):
        try:
            extension = ruta_archivo.split('.')[-1].lower()

            if extension in ['xlsx', 'xls']:
                try:
                    df_cleaned = pd.read_excel(ruta_archivo, sheet_name=None)
                except Exception as e:
                    raise ValueError(f"Error al leer archivo Excel: {e}. Intentando cargar como CSV...")

            elif extension == 'csv':
                df_cleaned = pd.read_csv(ruta_archivo)
            elif extension == 'txt':
                df_cleaned = pd.read_csv(ruta_archivo, delimiter='\t')
            else:
                raise ValueError(f"Formato no soportado: {extension}")

            return df_cleaned

        except Exception as e:
            raise ValueError(f"Ocurrió un error al procesar el archivo: {e}")

    def procesar_datos(self, df_cleaned):
        try:
            if isinstance(df_cleaned, dict):
                if 'Hoja1' in df_cleaned:
                    df_cleaned = df_cleaned['Hoja1']
                elif 'Sheet1' in df_cleaned:
                    df_cleaned = df_cleaned['Sheet1']
                else:
                    raise ValueError("No se encontró 'Hoja1' ni 'Sheet1'. Intentando con la primera hoja disponible.")
                    df_cleaned = next(iter(df_cleaned.values()))

            df_cleaned.columns = [
                'Fecha Local', 'Fecha UTC', 'Dirección del Viento', 'Dirección de ráfaga (grados)',
                'Rapidez de viento (km/h)', 'Rapidez de ráfaga (km/h)', 'Temperatura', 'Humedad',
                'Presión Atmosférica (hpa)', 'Precipitación (mm)', 'Radiación Solar (W/m²)'
            ]
            
            df_cleaned.drop(columns=['Fecha Local'], inplace=True)

            return df_cleaned

        except Exception as e:
            raise ValueError(f"Error al procesar los datos: {e}")

    def cargar_excel(self, ruta_excel):
        try:
            excel_file = pd.ExcelFile(ruta_excel)
            sheet_name = None

            if 'Hoja1' in excel_file.sheet_names:
                sheet_name = 'Hoja1'
            elif 'Sheet1' in excel_file.sheet_names:
                sheet_name = 'Sheet1'
            else:
                raise ValueError("No se encontró la hoja 'Hoja1' ni 'Sheet1' en el archivo Excel.")

            df_cleaned = pd.read_excel(ruta_excel, sheet_name=sheet_name, header=8, usecols="B:L")

            df_cleaned.columns = [
                'Fecha Local', 'Fecha UTC', 'Dirección del Viento', 'Dirección de ráfaga (grados)',
                'Rapidez de viento (km/h)', 'Rapidez de ráfaga (km/h)', 'Temperatura', 'Humedad',
                'Presión Atmosférica (hpa)', 'Precipitación (mm)', 'Radiación Solar (W/m²)'
            ]
            
            df_cleaned.drop(columns=['Fecha Local'], inplace=True)

            df_initial = pd.read_excel(ruta_excel, sheet_name=sheet_name, header=None, usecols=[11])
            altura = df_initial.iloc[6, 0]
            df_cleaned['Altura'] = altura

            df_cleaned['Fecha UTC'] = pd.to_datetime(df_cleaned['Fecha UTC'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
            df_cleaned.set_index('Fecha UTC', inplace=True)

            df_cleaned = df_cleaned.apply(pd.to_numeric, errors='coerce')

            df_cleaned['Humedad'] = df_cleaned['Humedad'] / 100

            df_hourly_avg = df_cleaned.resample('h').mean()
            df_interpolated = df_hourly_avg.interpolate(method='linear')
            df_interpolated_reset = df_interpolated.reset_index()

            column_order = [
                'Temperatura', 'Humedad', 'Altura', 'Fecha UTC',
                'Dirección del Viento', 'Dirección de ráfaga (grados)', 'Rapidez de viento (km/h)',
                'Rapidez de ráfaga (km/h)', 'Presión Atmosférica (hpa)', 'Precipitación (mm)', 'Radiación Solar (W/m²)'
            ]
            df_interpolated_reset = df_interpolated_reset[column_order]

            df_interpolated_reset['Columna Vacía'] = pd.NA

            df_daily_avg = df_cleaned.resample('D').mean()

            df_daily_avg['fecha_hora'] = df_daily_avg.index.strftime('%Y-%m-%d 00:00:00')

            df_daily_avg['Temperatura_Bs'] = df_daily_avg['Temperatura']
            df_daily_avg['Radiación Solar (W/m²)_Bs'] = df_daily_avg['Radiación Solar (W/m²)']

            df_daily_avg_reset = df_daily_avg.reset_index()[['fecha_hora', 'Temperatura_Bs', 'Radiación Solar (W/m²)_Bs']]

            df_final = pd.concat([df_interpolated_reset, df_daily_avg_reset], axis=1)

            return df_final

        except Exception as e:
            raise ValueError(f"Ocurrió un error al leer el archivo: {e}")

    def guardar_excel(self, tabla, ruta_guardado):
        datos_guardar = []
        for i, item in enumerate(tabla.get_children(), start=1):
            datos_fila = [i]
            datos_fila.extend(tabla.item(item)['values'])
            datos_guardar.append(datos_fila)

        columnas = ["#", "Tbs (°C)", "φ (%)", "Pvs (kPa)", "Pv (kPa)", "Ws (kg_vp/kg_AS)", "W (kg_vp/kg_AS)", "μ [G_sat]", 
                    "Veh (m³/kg_AS)", "h (kJ/kg_AS)", "Tpr (°C)", "Tbh (°C)"]
        df = pd.DataFrame(datos_guardar, columns=columnas)
        df.to_excel(ruta_guardado, index=False)

class InterfazGrafica:
    def __init__(self, calculadora, manejo_datos):
        self.calculadora = calculadora
        self.manejo_datos = manejo_datos
        self.ruta_excel = ''
        self.datos_excel = None

    def iniciar_interfaz(self):
        self.ventana_resultados = tk.Tk()
        self.ventana_resultados.title("Calculadora de Propiedades")
        self.ventana_resultados.geometry("1450x550")
        
        self.crear_botones()
        self.crear_tabla()
        self.ventana_resultados.mainloop()

    def crear_botones(self):
        boton_cargar = tk.Button(self.ventana_resultados, text="Cargar Excel", command=self.cargar_excel, padx=20, pady=10, font=("Arial", 12))
        boton_cargar.pack(pady=20)

        boton_guardar = tk.Button(self.ventana_resultados, text="Guardar Excel", command=self.guardar_excel, padx=20, pady=10, font=("Arial", 12))
        boton_guardar.pack(pady=20)

        boton_psicrometrica = tk.Button(self.ventana_resultados, text="Graficar Psicrométrica", command=self.graficar_psicrometrica, padx=20, pady=10, font=("Arial", 12))
        boton_psicrometrica.pack(pady=20)

        boton_climograma = tk.Button(self.ventana_resultados, text="Graficar Climograma", command=self.graficar_climograma, padx=20, pady=10, font=("Arial", 12))
        boton_climograma.pack(pady=20)

    def crear_tabla(self):
        columnas = ("Tbs (°C)", "φ (%)", "Pvs (kPa)", "Pv (kPa)", "Ws (kg_vp/kg_AS)", "W (kg_vp/kg_AS)", "μ [G_sat]", 
                    "Veh (m³/kg_AS)", "h (kJ/kg_AS)", "Tpr (°C)", "Tbh (°C)")
        self.tabla = ttk.Treeview(self.ventana_resultados, columns=columnas, height=16)
        for col in columnas:
            self.tabla.heading(col, text=col)
            self.tabla.column(col, width=100, anchor='center')
        self.tabla.pack()

    def cargar_excel(self):
        try:
            self.ruta_excel = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
            if self.ruta_excel:
                self.datos_excel = self.manejo_datos.cargar_excel(self.ruta_excel)
                Tbs_list, W_list = [], []
                resultados = self.datos_excel.apply(lambda row: self.calcular_resultados_fila(row, Tbs_list, W_list), axis=1)

                for row in self.tabla.get_children():
                    self.tabla.delete(row)

                for resultado in resultados.values:
                    self.tabla.insert("", "end", values=resultado)
            else:
                messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al leer el archivo: {e}")

    def calcular_resultados_fila(self, row, Tbs_list, W_list):
        Tbs = row['Temperatura']
        Hr = row['Humedad']
        altura = row['Altura']

        if Hr <= 0:
            Hr = 0.00001
        if altura < 0:
            raise ValueError(f"Error: La altura {altura} no es válida. Solo se permiten valores positivos.")

        presionAt = self.calculadora.calcular_presion(altura)
        pvs2 = self.calculadora.calcular_pvs(Tbs)
        Pv = self.calculadora.calcular_pv(Hr, pvs2)
        W = self.calculadora.razon_humedad(Pv, presionAt)
        
        Tbs_list.append(Tbs)
        W_list.append(W)
        
        Ws = self.calculadora.razon_humedad_saturada(pvs2, presionAt)
        Gsaturacion = self.calculadora.grado_saturacion(W, Ws)
        Veh = self.calculadora.volumen_especifico(Tbs, presionAt, W)
        Tpr = self.calculadora.temperatura_punto_rocio(Tbs, Pv)
        h = self.calculadora.entalpia(Tbs, W)  
        Tbh = self.calculadora.bulbo_humedo(presionAt, Tbs, W)

        return [Tbs, Hr, pvs2, Pv, Ws, W, Gsaturacion, Veh, h, Tpr, Tbh]

    def guardar_excel(self):
        try:
            ruta_guardado = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if ruta_guardado:
                self.manejo_datos.guardar_excel(self.tabla, ruta_guardado)
                messagebox.showinfo("Éxito", f"Los datos se han guardado correctamente en '{ruta_guardado}'.")
            else:
                messagebox.showwarning("Advertencia", "No se ha seleccionado una ubicación para guardar el archivo.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al guardar el archivo: {e}")

    def graficar_psicrometrica(self):
        try:
            Tbs = list(range(0, 60, 1))
            Hr = np.arange(0.1, 1.1, 0.1)
            altura = 0

            valores_W = []
            for temperatura in Tbs:
                fila_W = []
                for humedad in Hr:
                    presionAt = self.calculadora.calcular_presion(altura)
                    pvs2 = self.calculadora.calcular_pvs(temperatura)
                    Pv = self.calculadora.calcular_pv(humedad, pvs2)
                    W = self.calculadora.razon_humedad(Pv, presionAt)
                    fila_W.append(W)
                valores_W.append(fila_W)

            plt.figure(figsize=(8, 6))

            for i, hr in enumerate(Hr):
                plt.plot(Tbs, [fila[i] for fila in valores_W], label=f'Hr {hr * 100:.0f}%')

            if self.datos_excel is not None:
                Tbs_list = self.datos_excel['Temperatura'].values
                W_list = [self.calculadora.razon_humedad(
                    self.calculadora.calcular_pv(row['Humedad'], self.calculadora.calcular_pvs(row['Temperatura'])),
                    self.calculadora.calcular_presion(row['Altura'])
                ) for _, row in self.datos_excel.iterrows()]
                plt.scatter(Tbs_list, W_list, marker='o', color='blue', label='Datos')

                grid_x, grid_y = np.mgrid[min(Tbs_list):max(Tbs_list):100j, min(W_list):max(W_list):100j]
                points = np.column_stack((Tbs_list, W_list))
                grid_z = griddata(points, W_list, (grid_x, grid_y), method='linear')

                plt.contourf(grid_x, grid_y, grid_z, levels=15, cmap='Greens', alpha=0.5)

            plt.xlabel('Tbs (°C)')
            plt.ylabel('W (kg_vp/kg_AS)')
            plt.title('Gráfico Psicrométrico: Tbs vs W')
            plt.legend()
            plt.grid(True)

            plt.show()
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar el gráfico psicrométrico: {e}")

    def graficar_climograma(self):
        try:
            if self.datos_excel is None:
                raise ValueError("No se han cargado datos desde un archivo Excel.")

            dias = np.arange(0, 90)
            temperatura = self.datos_excel.iloc[:90, 13]
            radiacion = self.datos_excel.iloc[:90, 14]

            fig, ax1 = plt.subplots()

            ax1.set_xlabel('Día')
            ax1.set_ylabel('Radiación Global Promedio (W/m²)', color='tab:blue')
            ax1.bar(dias, radiacion, color='blue', label='Radiación (W/m²)', alpha=0.7)
            ax1.tick_params(axis='y', labelcolor='tab:blue')

            ax2 = ax1.twinx()
            ax2.set_ylabel('Temperatura Promedio (°C)', color='tab:red')
            ax2.plot(dias, temperatura, color='red', marker='o', linestyle='-', label='Temperatura (°C)', linewidth=2, markersize=6)
            ax2.tick_params(axis='y', labelcolor='tab:red')

            ax1.set_xticks(np.arange(0, 91, 10))

            fig.tight_layout()
            plt.title("Climograma: Temperatura vs Radiación (Día 0 al 90)")
            plt.show()

        except Exception as e:
            messagebox.showerror("Error", f"Error al generar el climograma: {e}")

if __name__ == "__main__":
    calculadora = CalculadoraPropiedades()
    manejo_datos = ManejoDatos()
    interfaz = InterfazGrafica(calculadora, manejo_datos)
    interfaz.iniciar_interfaz()
