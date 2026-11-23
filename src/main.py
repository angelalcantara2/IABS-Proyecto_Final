"""
main.py: interfaz gráfica (tkinter) del SSDC.
Permite elegir un paciente, correr BFS, UCS o A*, ver la ruta en rojo sobre
el grafo y abrir o exportar el reporte de los 12 pacientes.
"""
import os
import sys
import time
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import networkx as nx
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from lector import lectura_datos, obtener_paciente, ruta_archivo
from grafo import construir_grafo, calcular_costo_arista
from busqueda import bfs, ucs, a_estrella, construir_heuristica


FACTORES = ["Insuficiencia_Renal", "Cardiopatia", "Edad_Avanzada_Fragil", "Inmunosupresion"]

# Mismas posiciones que dibujar_grafo en grafo.py
POSICIONES = {"Diagnostico": (0, 0), "Post_Op_Mayor": (1, -0.6), "Post_Op_Menor": (1, -1.8),
              "Consolidacion": (2, -1.2), "Alta": (3, 0)}

DIR_SRC = os.path.dirname(os.path.abspath(__file__))
SCRIPT_INFORME = os.path.join(DIR_SRC, "armar_informe.py")
PDF_INFORME = os.path.join(os.path.dirname(DIR_SRC), "docs", "SSDC_Entregables_2_y_3.pdf")

COLOR_FONDO = "#f5f7fa"
COLOR_TITULO = "#1f3b57"

def grafo_del_paciente(p, matriz):
    """Construye el grafo del paciente y su heurística para A*."""
    costos = [calcular_costo_arista(fila, p) for _, fila in matriz.iterrows()]
    G = construir_grafo(matriz, costos)
    heuristica = construir_heuristica(G, p["Severidad_Diagnostico"])
    return G, heuristica


def ejecutar(nombre, G, heuristica):
    """Ejecuta un algoritmo y mide el tiempo. Devuelve (ruta, costo, nodos, tiempo_ms)."""
    inicio_t = time.perf_counter()
    if nombre == "BFS":
        ruta, costo, nodos = bfs(G)
    elif nombre == "UCS":
        ruta, costo, nodos = ucs(G)
    else:
        ruta, costo, nodos = a_estrella(G, heuristica)
    tiempo_ms = (time.perf_counter() - inicio_t) * 1000
    return ruta, costo, nodos, tiempo_ms


def resumir_ruta(ruta, matriz):
    """Devuelve (tratamientos, costo en USD, efectividad estimada) de una ruta."""
    tratamientos = ruta[1::2]
    costo_usd = 0
    efectividad = 1.0
    for t in tratamientos:
        fila_t = matriz[matriz["Tratamiento"] == t].iloc[0]
        costo_usd += fila_t["Costo_Financiero_Base"]
        efectividad *= fila_t["Tasa_Efectividad_Base"]
    return tratamientos, costo_usd, efectividad


def tabla_comparativa(pacientes, matriz):
    """Los 12 pacientes x 3 algoritmos, como DataFrame."""
    filas = []
    for _, p in pacientes.iterrows():
        G, heuristica = grafo_del_paciente(p, matriz)
        for nombre in ["BFS", "UCS", "A*"]:
            ruta, costo, nodos, tiempo_ms = ejecutar(nombre, G, heuristica)
            tratamientos, costo_usd, efectividad = resumir_ruta(ruta, matriz)
            filas.append({"ID_Paciente": p["ID_Paciente"],
                          "Algoritmo": nombre,
                          "Ruta Propuesta": " → ".join(tratamientos),
                          "Costo Económico": costo_usd,
                          "Nodos Expandidos": nodos,
                          "Costo Unificado": round(costo, 2),
                          "Efectividad Estimada": round(efectividad, 3),
                          "Tiempo (ms)": round(tiempo_ms, 4)})
    return pd.DataFrame(filas)

class AppSSDC(tk.Tk):

    def __init__(self, pacientes, matriz):
        super().__init__()
        self.pacientes = pacientes
        self.matriz = matriz
        self.paciente = None
        self.G = None
        self.heuristica = None
        self.rutas = {}            

        self.title("SSDC - Sistema de Soporte a Decisiones Clínicas")
        self.geometry("1280x760")
        self.minsize(1050, 650)
        self.configure(bg=COLOR_FONDO)
        self._estilos()
        self._armar_ventana()
        self._llenar_lista()

    def _estilos(self):
        estilo = ttk.Style(self)
        if "clam" in estilo.theme_names():
            estilo.theme_use("clam")
        estilo.configure("TFrame", background=COLOR_FONDO)
        estilo.configure("Tarjeta.TFrame", background="white", relief="solid", borderwidth=1)
        estilo.configure("TLabel", background=COLOR_FONDO, font=("Segoe UI", 10))
        estilo.configure("Tarjeta.TLabel", background="white", font=("Segoe UI", 10))
        estilo.configure("Titulo.TLabel", font=("Segoe UI", 16, "bold"), foreground=COLOR_TITULO)
        estilo.configure("Seccion.TLabel", font=("Segoe UI", 11, "bold"), foreground=COLOR_TITULO)
        estilo.configure("SeccionT.TLabel", background="white", font=("Segoe UI", 11, "bold"),
                         foreground=COLOR_TITULO)
        estilo.configure("TButton", font=("Segoe UI", 10), padding=6)
        estilo.configure("Accion.TButton", font=("Segoe UI", 10, "bold"), padding=8)
        estilo.configure("Treeview", rowheight=26, font=("Segoe UI", 9))
        estilo.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _armar_ventana(self):
        barra = ttk.Frame(self, padding=(14, 10))
        barra.pack(fill="x")
        ttk.Label(barra, text="Sistema de Soporte a Decisiones Clínicas",
                  style="Titulo.TLabel").pack(side="left")
        ttk.Button(barra, text="📊  Reporte comparativo (12 pacientes)",
                   style="Accion.TButton", command=self.abrir_reporte).pack(side="right")
        self.btn_informe = ttk.Button(barra, text="📄  Exportar informe (PDF)",
                                      style="Accion.TButton", command=self.exportar_informe)
        self.btn_informe.pack(side="right", padx=(0, 8))

        cuerpo = ttk.Frame(self, padding=(14, 0, 14, 14))
        cuerpo.pack(fill="both", expand=True)

        izq = ttk.Frame(cuerpo, width=300)
        izq.pack(side="left", fill="y", padx=(0, 12))
        izq.pack_propagate(False)

        ttk.Label(izq, text="Buscar paciente", style="Seccion.TLabel").pack(anchor="w")
        self.var_buscar = tk.StringVar()
        self.var_buscar.trace_add("write", lambda *_: self._llenar_lista())
        entrada = ttk.Entry(izq, textvariable=self.var_buscar, font=("Segoe UI", 10))
        entrada.pack(fill="x", pady=(4, 6))
        entrada.focus_set()

        marco_lista = ttk.Frame(izq)
        marco_lista.pack(fill="x")
        self.lista = tk.Listbox(marco_lista, height=10, font=("Segoe UI", 10),
                                activestyle="none", selectbackground="#2e5578",
                                exportselection=False)
        self.lista.pack(side="left", fill="x", expand=True)
        barra_lista = ttk.Scrollbar(marco_lista, orient="vertical", command=self.lista.yview)
        barra_lista.pack(side="right", fill="y")
        self.lista.configure(yscrollcommand=barra_lista.set)
        self.lista.bind("<<ListboxSelect>>", self._al_elegir_paciente)

        tarjeta = ttk.Frame(izq, style="Tarjeta.TFrame", padding=12)
        tarjeta.pack(fill="both", expand=True, pady=(12, 0))
        ttk.Label(tarjeta, text="Perfil del paciente", style="SeccionT.TLabel").pack(anchor="w")
        self.lbl_perfil = ttk.Label(tarjeta, text="Elige un paciente de la lista.",
                                    style="Tarjeta.TLabel", justify="left", wraplength=260)
        self.lbl_perfil.pack(anchor="w", pady=(8, 0))

        # ---- Columna derecha: botones, resultados y grafo ----
        der = ttk.Frame(cuerpo)
        der.pack(side="left", fill="both", expand=True)

        ttk.Label(der, text="Algoritmo de búsqueda", style="Seccion.TLabel").pack(anchor="w")
        botones = ttk.Frame(der)
        botones.pack(fill="x", pady=(4, 8))
        textos = [("BFS  (Nivel 1)", ["BFS"]),
                  ("UCS  (Nivel 2)", ["UCS"]),
                  ("A*  (Nivel 3)", ["A*"]),
                  ("Comparar los 3", ["BFS", "UCS", "A*"])]
        self.botones_alg = []
        for texto, algoritmos in textos:
            b = ttk.Button(botones, text=texto, state="disabled",
                           command=lambda a=algoritmos: self.correr(a))
            b.pack(side="left", padx=(0, 8))
            self.botones_alg.append(b)

        columnas = ("Algoritmo", "Ruta", "USD", "Efectividad", "Costo unificado", "Nodos", "Tiempo (ms)")
        anchos = (80, 360, 80, 85, 115, 60, 85)
        self.tabla = ttk.Treeview(der, columns=columnas, show="headings", height=3,
                                  selectmode="browse")
        for col, ancho in zip(columnas, anchos):
            self.tabla.heading(col, text=col)
            self.tabla.column(col, width=ancho, anchor="w" if col == "Ruta" else "center")
        self.tabla.pack(fill="x")
        self.tabla.bind("<<TreeviewSelect>>", self._al_elegir_resultado)
        self.lbl_ayuda = ttk.Label(der, text="", foreground="#5a6b7b")
        self.lbl_ayuda.pack(anchor="w", pady=(4, 6))

        marco_grafo = ttk.Frame(der, style="Tarjeta.TFrame", padding=4)
        marco_grafo.pack(fill="both", expand=True)
        self.figura = Figure(figsize=(8, 4.5), dpi=100)
        self.ax = self.figura.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figura, master=marco_grafo)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self._dibujar(None, "Elige un paciente para ver su grafo")

    def _llenar_lista(self):
        texto = self.var_buscar.get().strip().lower()
        self.lista.delete(0, "end")
        self.ids_lista = []
        for _, p in self.pacientes.iterrows():
            etiqueta = f'{p["ID_Paciente"]}  ·  {p["Nombre_Anonimizado"]}'
            if texto in etiqueta.lower():
                self.lista.insert("end", etiqueta)
                self.ids_lista.append(p["ID_Paciente"])

    def _al_elegir_paciente(self, _evento=None):
        seleccion = self.lista.curselection()
        if not seleccion:
            return
        id_paciente = self.ids_lista[seleccion[0]]
        self.paciente = obtener_paciente(self.pacientes, id_paciente)
        self.G, self.heuristica = grafo_del_paciente(self.paciente, self.matriz)
        self.rutas = {}

        p = self.paciente
        factores = [f.replace("_", " ") for f in FACTORES if p[f] == 1]
        if p["Minimizar_Costo_Financiero"] == 1:
            preferencia = "💰 Minimizar costo financiero"
        else:
            preferencia = "❤️ Maximizar supervivencia"
        self.lbl_perfil.configure(text=(
            f'ID: {p["ID_Paciente"]}\n'
            f'Nombre: {p["Nombre_Anonimizado"]}\n'
            f'Severidad: {p["Severidad_Diagnostico"]} de 5\n\n'
            f'Factores de riesgo:\n  ' + ("\n  ".join(factores) if factores else "Ninguno") + "\n\n"
            f'Preferencia:\n  {preferencia}'))

        for b in self.botones_alg:
            b.state(["!disabled"])
        self.tabla.delete(*self.tabla.get_children())
        self.lbl_ayuda.configure(text="Pulsa un algoritmo para calcular la ruta.")
        self._dibujar(None, f'Grafo de {p["ID_Paciente"]} (costo unificado en cada flecha)')

    def correr(self, algoritmos):
        if self.paciente is None:
            return
        self.tabla.delete(*self.tabla.get_children())
        self.rutas = {}
        for nombre in algoritmos:
            ruta, costo, nodos, tiempo_ms = ejecutar(nombre, self.G, self.heuristica)
            if ruta is None:
                self.tabla.insert("", "end", iid=nombre,
                                  values=(nombre, "No se encontró ruta", "", "", "", nodos, ""))
                continue
            tratamientos, costo_usd, efectividad = resumir_ruta(ruta, self.matriz)
            self.rutas[nombre] = ruta
            self.tabla.insert("", "end", iid=nombre, values=(
                nombre, " → ".join(tratamientos), f"${costo_usd:,.0f}",
                f"{efectividad:.0%}", f"{costo:,.0f}", nodos, f"{tiempo_ms:.3f}"))

        ultimo = "UCS" if len(algoritmos) > 1 else algoritmos[0]
        if ultimo in self.rutas:
            self.tabla.selection_set(ultimo)
        if len(algoritmos) > 1:
            self.lbl_ayuda.configure(text="Haz clic en una fila para ver esa ruta en rojo en el grafo.")
        else:
            self.lbl_ayuda.configure(text="")

    def _al_elegir_resultado(self, _evento=None):
        seleccion = self.tabla.selection()
        if not seleccion or seleccion[0] not in self.rutas:
            return
        nombre = seleccion[0]
        self._dibujar(self.rutas[nombre], f'{self.paciente["ID_Paciente"]} - ruta de {nombre}')

    def _dibujar(self, ruta, titulo):
        ax = self.ax
        ax.clear()
        ax.set_axis_off()
        if self.G is None:
            ax.text(0.5, 0.5, titulo, ha="center", va="center", fontsize=12, color="#5a6b7b",
                    transform=ax.transAxes)
            self.canvas.draw()
            return

        G = self.G
        pos = {n: POSICIONES.get(n, (i, -2.5)) for i, n in enumerate(G.nodes())}

        nodos_ruta = ruta[0::2] if ruta else []
        aristas_ruta = list(zip(nodos_ruta, nodos_ruta[1:]))

        colores_nodos = ["salmon" if n in nodos_ruta else "lightblue" for n in G.nodes()]
        nx.draw_networkx_nodes(G, pos, ax=ax, node_color=colores_nodos, node_size=2300)
        nx.draw_networkx_labels(G, pos, ax=ax, font_size=8)
        nx.draw_networkx_edges(G, pos, ax=ax, node_size=2300, arrowsize=15, edge_color="#444")
        if aristas_ruta:
            nx.draw_networkx_edges(G, pos, ax=ax, edgelist=aristas_ruta, node_size=2300,
                                   arrowsize=20, edge_color="red", width=3)

        etiquetas = {}
        for u, v, k, d in G.edges(keys=True, data=True):
            etiquetas[(u, v)] = etiquetas.get((u, v), "") + f"{k}: {d['costo']:,.0f}\n"
        nx.draw_networkx_edge_labels(G, pos, ax=ax, edge_labels=etiquetas, font_size=7)

        ax.set_title(titulo, fontsize=11, color=COLOR_TITULO)
        ax.margins(0.12)
        self.figura.tight_layout()
        self.canvas.draw()

    def abrir_reporte(self):
        tabla = tabla_comparativa(self.pacientes, self.matriz)
        print("\n", tabla.to_string(index=False))

        ventana = tk.Toplevel(self)
        ventana.title("Reporte comparativo - 12 pacientes x 3 algoritmos")
        ventana.geometry("1180x640")
        ventana.configure(bg=COLOR_FONDO)

        arriba = ttk.Frame(ventana, padding=10)
        arriba.pack(fill="x")
        ttk.Label(arriba, text="Mostrar:", style="Seccion.TLabel").pack(side="left")
        var_filtro = tk.StringVar(value="Todos")
        combo = ttk.Combobox(arriba, textvariable=var_filtro, state="readonly", width=10,
                             values=["Todos", "BFS", "UCS", "A*"])
        combo.pack(side="left", padx=8)

        def exportar():
            ruta = filedialog.asksaveasfilename(parent=ventana, defaultextension=".csv",
                                                initialfile="reporte_comparativo.csv",
                                                filetypes=[("CSV", "*.csv")])
            if ruta:
                tabla.to_csv(ruta, index=False, encoding="utf-8-sig")
                messagebox.showinfo("Reporte", f"Reporte guardado en:\n{ruta}", parent=ventana)

        ttk.Button(arriba, text="💾  Exportar a CSV", command=exportar).pack(side="right")

        columnas = list(tabla.columns)
        anchos = {"ID_Paciente": 85, "Algoritmo": 75, "Ruta Propuesta": 340, "Costo Económico": 115,
                  "Nodos Expandidos": 130, "Costo Unificado": 120, "Efectividad Estimada": 145,
                  "Tiempo (ms)": 90}
        marco = ttk.Frame(ventana, padding=(10, 0, 10, 0))
        marco.pack(fill="both", expand=True)
        arbol = ttk.Treeview(marco, columns=columnas, show="headings")
        for col in columnas:
            arbol.heading(col, text=col)
            arbol.column(col, width=anchos.get(col, 100),
                         anchor="w" if col == "Ruta Propuesta" else "center")
        arbol.pack(side="left", fill="both", expand=True)
        barra = ttk.Scrollbar(marco, orient="vertical", command=arbol.yview)
        barra.pack(side="right", fill="y")
        arbol.configure(yscrollcommand=barra.set)
        arbol.tag_configure("par", background="#eef3f8")

        promedios = tabla.groupby("Algoritmo")["Nodos Expandidos"].mean()
        texto_resumen = "Promedio de nodos expandidos:   " + "     ".join(
            f"{a}: {promedios[a]:.2f}" for a in ["BFS", "UCS", "A*"])
        ttk.Label(ventana, text=texto_resumen, padding=10).pack(anchor="w")

        def llenar(*_):
            arbol.delete(*arbol.get_children())
            filtro = var_filtro.get()
            datos = tabla if filtro == "Todos" else tabla[tabla["Algoritmo"] == filtro]
            for i, (_, fila) in enumerate(datos.iterrows()):
                valores = list(fila)
                valores[3] = f"${fila['Costo Económico']:,.0f}"
                valores[5] = f"{fila['Costo Unificado']:,.0f}"
                valores[6] = f"{fila['Efectividad Estimada']:.0%}"
                arbol.insert("", "end", values=valores,
                             tags=("par",) if (i // 3) % 2 == 0 else ())

        combo.bind("<<ComboboxSelected>>", llenar)
        llenar()


    # ---- Exportar informe: corre armar_informe.py sin congelar la ventana ----
    def exportar_informe(self):
        self.btn_informe.configure(text="⏳  Generando informe...", state="disabled")
        self.config(cursor="watch")
        threading.Thread(target=self._generar_informe, daemon=True).start()

    def _generar_informe(self):
        try:
            res = subprocess.run([sys.executable, SCRIPT_INFORME], cwd=DIR_SRC,
                                 capture_output=True, text=True, encoding="utf-8", errors="replace")
            ok, salida = res.returncode == 0, (res.stdout + res.stderr).strip()
        except Exception as e:
            ok, salida = False, str(e)
        self.after(0, lambda: self._fin_informe(ok, salida))

    def _fin_informe(self, ok, salida):
        self.btn_informe.configure(text="📄  Exportar informe (PDF)", state="normal")
        self.config(cursor="")
        if not ok or not os.path.exists(PDF_INFORME):
            ultimas = "\n".join(salida.splitlines()[-8:])
            messagebox.showerror("Informe", f"No se pudo generar el informe.\n\n{ultimas}", parent=self)
            return
        if messagebox.askyesno("Informe", f"Informe generado en:\n{PDF_INFORME}\n\n¿Quieres abrir el PDF?",
                               parent=self):
            if sys.platform.startswith("win"):
                os.startfile(PDF_INFORME)
            else:
                subprocess.Popen(["xdg-open" if sys.platform.startswith("linux") else "open", PDF_INFORME])


def iniciar():
    try:
        pacientes, matriz = lectura_datos(ruta_archivo)
    except FileNotFoundError:
        raiz = tk.Tk()
        raiz.withdraw()
        messagebox.showerror("SSDC", f"No se encontró el Excel en:\n{ruta_archivo}")
        return
    app = AppSSDC(pacientes, matriz)
    app.mainloop()


iniciar()