import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
# Importamos las funciones lógicas de tu script original
from procesar_pdfs import procesar_pdf, crear_excel 

class AppProcesador(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Procesador Masivo de PDFs a Excel")
        self.geometry("600x450")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Variables
        self.ruta_carpeta = ctk.StringVar()
        self.idioma_ocr = ctk.StringVar(value="spa")

        # --- Layout ---
        self.grid_columnconfigure(0, weight=1)
        
        # Título
        self.label_titulo = ctk.CTkLabel(self, text="Conversor PDF → Excel", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_titulo.pack(pady=20)

        # Selección de Carpeta
        self.frame_carpeta = ctk.CTkFrame(self)
        self.frame_carpeta.pack(fill="x", padx=20, pady=10)
        
        self.entry_ruta = ctk.CTkEntry(self.frame_carpeta, textvariable=self.ruta_carpeta, width=350)
        self.entry_ruta.pack(side="left", padx=10, pady=10)
        
        self.btn_explorar = ctk.CTkButton(self.frame_carpeta, text="Explorar", command=self.seleccionar_carpeta)
        self.btn_explorar.pack(side="right", padx=10)

        # Configuración OCR
        self.frame_config = ctk.CTkFrame(self)
        self.frame_config.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(self.frame_config, text="Idioma OCR:").pack(side="left", padx=10)
        self.combo_idioma = ctk.CTkComboBox(self.frame_config, values=["spa", "eng"], variable=self.idioma_ocr)
        self.combo_idioma.pack(side="left", padx=10)

        # Barra de Progreso y Estado
        self.progreso = ctk.CTkProgressBar(self)
        self.progreso.pack(fill="x", padx=20, pady=20)
        self.progreso.set(0)

        self.label_estado = ctk.CTkLabel(self, text="Esperando selección...")
        self.label_estado.pack()

        # Botón Principal
        self.btn_iniciar = ctk.CTkButton(self, text="INICIAR PROCESAMIENTO", 
                                         command=self.iniciar_hilo, 
                                         fg_color="#2ecc71", hover_color="#27ae60",
                                         text_color="black", font=ctk.CTkFont(weight="bold"))
        self.btn_iniciar.pack(pady=20)

    def seleccionar_carpeta(self):
        ruta = filedialog.askdirectory()
        if ruta:
            self.ruta_carpeta.set(ruta)
            self.label_estado.configure(text=f"Carpeta seleccionada con éxito")

    def iniciar_hilo(self):
        if not self.ruta_carpeta.get():
            messagebox.showwarning("Atención", "Por favor selecciona una carpeta primero.")
            return
        
        # Ejecutar en un hilo separado para no congelar la GUI
        thread = threading.Thread(target=self.ejecutar_proceso)
        thread.start()

    def ejecutar_proceso(self):
        self.btn_iniciar.configure(state="disabled")
        carpeta = self.ruta_carpeta.get()
        
        # Buscar PDFs (Lógica basada en procesar_pdfs.py)
        pdfs = [os.path.join(carpeta, f) for f in os.listdir(carpeta) if f.lower().endswith(".pdf")]
        
        if not pdfs:
            messagebox.showerror("Error", "No se encontraron PDFs en la carpeta seleccionada.")
            self.btn_iniciar.configure(state="normal")
            return

        resultados = []
        total = len(pdfs)

        for i, ruta in enumerate(pdfs):
            self.label_estado.configure(text=f"Procesando: {os.path.basename(ruta)}")
            res = procesar_pdf(ruta)
            resultados.append(res)
            self.progreso.set((i + 1) / total)

        self.label_estado.configure(text="Generando archivo Excel...")
        ok, ocr, errores = crear_excel(resultados, "resultado_consolidado_gui.xlsx")
        
        self.btn_iniciar.configure(state="normal")
        self.label_estado.configure(text="¡Proceso Finalizado!")
        
        messagebox.showinfo("Completado", 
                            f"Proceso terminado:\n- Correctos: {ok}\n- Con OCR: {ocr}\n- Errores: {errores}\n\nArchivo guardado como: resultado_consolidado_gui.xlsx")

if __name__ == "__main__":
    app = AppProcesador()
    app.mainloop()