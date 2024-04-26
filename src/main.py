import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import oferta
import os
import sys


def get_resource_path():
    """ Retorna la ruta absoluta al recurso, para uso en desarrollo o en el ejecutable empaquetado. """
    if getattr(sys, 'frozen', False):
        # Si el programa ha sido empaquetado, el directorio base es el que PyInstaller proporciona
        base_path = sys._MEIPASS
    else:
        # Si estamos en desarrollo, utiliza la ubicación del script
        base_path = os.path.dirname(os.path.realpath(__file__))

    return base_path


script_directory=get_resource_path()

class App(tk.Tk):
    def __init__(self, titulo, geometria):
        #inicializacion
        super().__init__()
        self.title(titulo)
        self.geometry(f'{geometria[0]}x{geometria[1]}')
        self.resizable(True,True)
        self.minsize(869,500)
        # self.icono_e_imagen(self)
        
        # Configurar la expansión de las filas y columnas
        self.grid_columnconfigure(0, weight=1)  # Hace que la columna se expanda
        self.grid_rowconfigure(0, weight=1)     # Hace que la fila del FirstFrame se expanda
        self.grid_rowconfigure(1, weight=0)     # Mantiene la fila de los botones sin expandir


        #Botones de Navegación
        self.navigation_buttons()

        self.firstStep=SecondFrame(self)
        self.firstStep.grid(row=0,column=0,sticky='nsew',padx=16,pady=16)

        #Correr la app
        self.mainloop()



    def icono_e_imagen(self):
        self.iconbitmap(os.path.join(script_directory,"imagen.ico"))
        imagen_ico = Image.open(os.path.join(script_directory,"imagen.ico"))
        mi_imagen=imagen_ico.resize((48,48))
        mi_imagen = ImageTk.PhotoImage(mi_imagen)

    def style_configure(self):
        colordefondo='#2b394a'
        foregroundvariable='grey85'
        self.configure(background=colordefondo)

        style = ttk.Style()

        style.configure("Large.TLabel", font=("Arial", 14), foreground=foregroundvariable,
                        background=colordefondo)

        style.configure('Custom.TFrame', foreground=foregroundvariable,
                        background=colordefondo)

        style.configure('Custom.TRadiobutton', font=('Arial', 14), foreground=foregroundvariable,
                        background=colordefondo)

        style.configure("Bold.TLabel", font=("Arial", 17,'bold'), foreground=foregroundvariable,
                        background=colordefondo)

        style.configure("dis.TLabel", font=("Arial", 18), foreground='grey50',
                        background=colordefondo)

        style.configure('Custom.TButton', font=('Arial', 13,'bold'), 
                        foreground='grey15',background=colordefondo)

        style.configure('Custom.TEntry', 
                        font=('Arial', 25), 
                        fieldbackground='black', 
                        foreground='black', 
                        background=colordefondo)
        style.configure("Switch.TCheckbutton", background=colordefondo, 
                        foreground=foregroundvariable)
        
    def navigation_buttons(self):
        # Setup navigation buttons
        self.button_frame = ttk.Frame(self)
        self.button_frame.grid(row=1,column=0,sticky='ew',pady=20,padx=20)
        self.button_frame.grid_columnconfigure(0, weight=1)  # Asegura que el frame de botones se expanda horizontalmente


        self.back_button = ttk.Button(self.button_frame, text="Atrás")
        self.back_button.grid(row=1,column=0,sticky='e',pady=20,padx=(500,10))

        self.next_button = ttk.Button(self.button_frame, text="Siguiente")
        self.next_button.grid(row=1,column=1,sticky='e',pady=20,padx=10)

        self.cancel_button = ttk.Button(self.button_frame, text="Cancelar", command=self.quit)
        self.cancel_button.grid(row=1,column=2,sticky='e',pady=20,padx=30)

    def configurar_grid(self):
        # Configurar la expansión de las filas y columnas
        self.grid_columnconfigure(0, weight=1)  # Hace que la columna se expanda
        self.grid_rowconfigure(0, weight=1)     # Hace que la fila del FirstFrame se expanda
        self.grid_rowconfigure(1, weight=0)     # Mantiene la fila de los botones sin expandir



class FirstFrame(ttk.Frame):
    """
    Primer paso de la GUI. Dos entrys para seleccionar carpeta del proyecto y el PVP asociado
    PVP puede estar en cualquier carpeta, el paso posterior lo copia en la carpeta del proyecto.

    """
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets()
        self.place_widgets()
        
    def create_widgets(self):
        # Creation of GUI components to be used in the frame.
        self.labelCarpeta = ttk.Label(self, text='Seleccionar Carpeta de Proyecto')
        self.labelPVP = ttk.Label(self, text='Seleccione el PVP del Proyecto')
        
        # Initialization of variables for storing user input from file dialogs.
        self.entryCarpetaVar = tk.StringVar()
        self.entryPVPVar = tk.StringVar()
        
        # Entry widgets for displaying the paths, set to readonly to prevent manual user edits.
        self.entryCarpeta = ttk.Entry(self, textvariable=self.entryCarpetaVar, width=80, state='readonly')
        self.entryPVP = ttk.Entry(self, textvariable=self.entryPVPVar, width=80, state='readonly')
        
        # Buttons for triggering the browse dialogs.
        self.buttonCarpeta = ttk.Button(self, text='Examinar', width=25, command=self.browse_file_sp)
        self.buttonPVP = ttk.Button(self, text='Examinar', width=25, command=self.browse_file_pvp)

    def place_widgets(self):
        # Configure the grid layout, making sure some rows and columns are weighted to center content.
        self.columnconfigure(0, weight=1)
        self.columnconfigure(3, weight=1)

        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=0)
        self.rowconfigure(2, weight=0)
        self.rowconfigure(3, weight=1)
        self.rowconfigure(4, weight=0)
        self.rowconfigure(5, weight=0)
        self.rowconfigure(6, weight=1)

        # Placement of the widgets using grid layout to achieve desired UI structure.
        self.labelCarpeta.grid(row=1, column=1,sticky='w')
        self.labelPVP.grid(row=4, column=1,sticky='w')
        self.entryCarpeta.grid(row=2, column=1, sticky='ew')
        self.entryPVP.grid(row=5, column=1, sticky='ew')
        self.buttonCarpeta.grid(row=2, column=2, sticky='ew')
        self.buttonPVP.grid(row=5, column=2, sticky='ew')

    def browse_file_sp(self):
        # Browse for a directory and attempt to find a 'SP' prefixed file within it.
        carpeta = filedialog.askdirectory() #Ruta de la Carpeta del proyecto
        archivo_sp = next((archivo for archivo in os.listdir(carpeta) if archivo.startswith('SP')), None)
        if archivo_sp:
            self.ruta_sp=os.path.join(carpeta, archivo_sp) #ruta del sp
            self.entryCarpetaVar.set(carpeta)
        else:
            self.entryCarpetaVar.set('')
            messagebox.showerror("¡Error!","No hay un proyecto en la carpeta seleccionada!")

    def browse_file_pvp(self):
        # Browse for a file and validate it starts with 'PVP' to be considered a valid PVP file.
        ruta_pvp = filedialog.askopenfilename()
        archivo_pvp = ruta_pvp.split('/')[-1]
        if archivo_pvp.startswith('PVP'):
            self.entryPVPVar.set(ruta_pvp)
        else:
            self.entryCarpetaVar.set('')
            messagebox.showerror("¡Error!","Eso no es un PVP!")
            
class SecondFrame(ttk.Frame):

    def __init__(self,parent):
        super().__init__(parent)
        self.create_widgets(parent)
        self.place_widgets()

    def create_widgets(self,parent):
        #Labels encabezado
        self.carpeta=FirstFrame(parent).entryCarpetaVar.get().split('/')[-1]
        self.label=ttk.Label(self,text='Proyecto: ')
        self.labelProyecto=ttk.Label(self,text=self.carpeta)

        #TreeView
        self.comparacion=ttk.Treeview(self)

    def place_widgets(self):
        self.label.grid(row=0,column=0)
        self.labelProyecto.grid(row=0,column=1)
        self.comparacion.grid(row=1,column=0,columnspan=2)

    def comparar_sp_vs_pvp(self):
        pass



App("Crear Oferta",(1000,500))






