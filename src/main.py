import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import openpyxl
import oferta
import os
import sys
import shutil


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
        self.minsize(850,530)
        self.icono_e_imagen()
        
        # Configurar la expansión de las filas y columnas
        self.grid_columnconfigure(0, weight=1)  # Hace que la columna se expanda
        self.grid_rowconfigure(0, weight=1)     # Hace que la fila del Frame se expanda
        self.grid_rowconfigure(1, weight=0)     # Mantiene la fila de los botones sin expandir

        self.frames = {}

        self.frame_uno=FirstFrame(self)
        self.frame_uno.grid(row=0, column=0, sticky="nsew")
        self.frames[FirstFrame] = self.frame_uno

        self.frame_dos=SecondFrame(self)
        self.frame_dos.grid(row=0, column=0, sticky="nsew")
        self.frames[SecondFrame] = self.frame_dos

        self.navigation_buttons()

        self.current_frame = FirstFrame
        self.show_frame(FirstFrame)
        #Correr la app
        self.mainloop()



    def icono_e_imagen(self):
        self.iconbitmap(os.path.join(script_directory,'..','docs',"imagen.ico"))
        

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


        self.back_button = ttk.Button(self.button_frame, text="Atrás",command=self.go_back)
        self.back_button.grid(row=1,column=0,sticky='e',pady=20,padx=(500,10))

        self.next_button = ttk.Button(self.button_frame, text="Siguiente",command=self.go_next)
        self.next_button.grid(row=1,column=1,sticky='e',pady=20,padx=10)

        self.cancel_button = ttk.Button(self.button_frame, text="Cancelar", command=self.quit)
        self.cancel_button.grid(row=1,column=2,sticky='e',pady=20,padx=30)

    def update_buttons(self):
        
        self.back_button['state'] = 'normal' if self.current_frame != FirstFrame else 'disabled'
        self.next_button['state'] = 'normal' if self.current_frame != SecondFrame else 'disabled'
        self.next_button['text'] = 'Finalizar' if self.current_frame == SecondFrame else 'Siguiente'
        if self.current_frame == FirstFrame and not self.frames[FirstFrame].can_go_to_next_page():
            self.next_button['state'] = 'disabled'
        if self.current_frame == SecondFrame and self.frames[SecondFrame].canFinish():
            self.next_button['state'] = 'normal'

    def configurar_grid(self):
        # Configurar la expansión de las filas y columnas
        self.grid_columnconfigure(0, weight=1)  # Hace que la columna se expanda
        self.grid_rowconfigure(0, weight=1)     # Hace que la fila del FirstFrame se expanda
        self.grid_rowconfigure(1, weight=0)     # Mantiene la fila de los botones sin expandir

    def rutas(self):
       
        self.carpeta_proyecto=self.frame_uno.entryCarpetaVar.get()
        pvp=self.frame_uno.entryPVPVar.get()
        sp=self.frame_uno.ruta_sp
        try:
            destino=shutil.copy(pvp,self.carpeta_proyecto)
        except shutil.SameFileError:
            destino=pvp
        return (sp,destino)

    def show_frame(self, frame_class):
        frame = self.frames[frame_class]
        frame.tkraise()
        self.current_frame = frame_class
        self.update_buttons()

    def go_back(self):
        if self.current_frame == SecondFrame:
            self.show_frame(FirstFrame)


    def go_next(self):
        if self.current_frame == FirstFrame:
            sp,pvp=self.rutas()
            self.tablacomparativa=self.frame_dos.comparar_sp_vs_pvp(sp,pvp)
            self.frame_dos.create_widgets(self)
            self.show_frame(SecondFrame)
        elif self.current_frame == SecondFrame:
            try:
                oferta.llenar_oferta(self.carpeta_proyecto,self.frame_dos.df_pvp)
                messagebox.showinfo("Finalizado","Oferta finalizada\nRecuerde poner pre-requisitos.")
                exit()
            except Exception as e:
                messagebox.showerror("Error",str(e))


class FirstFrame(ttk.Frame):
    """
    Primer paso de la GUI. Dos entrys para seleccionar carpeta del proyecto y el PVP asociado
    PVP puede estar en cualquier carpeta, el paso posterior lo copia en la carpeta del proyecto.

    """
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets(parent)
        self.place_widgets()
        
    def create_widgets(self,parent):
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
        self.buttonCarpeta = ttk.Button(self, text='Examinar', width=25, command=lambda parent=parent:self.browse_project_directory(parent))
        self.buttonPVP = ttk.Button(self, text='Examinar', width=25, command=lambda parent=parent:self.browse_file_pvp(parent))

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

    def browse_project_directory(self,parent):
        # Browse for a directory and attempt to find a 'SP' prefixed file within it.
        carpeta = filedialog.askdirectory() #Ruta de la Carpeta del proyecto
        if carpeta:

            archivo_sp = next((archivo for archivo in os.listdir(carpeta) if archivo.startswith('SP')), None)

            if archivo_sp:
                self.ruta_sp=os.path.join(carpeta, archivo_sp) #ruta del sp
                self.entryCarpetaVar.set(carpeta)
            else:
                self.entryCarpetaVar.set('')
                messagebox.showerror("¡Error!","No hay un proyecto en la carpeta seleccionada!")
        else:
            self.entryCarpetaVar.set('')
            messagebox.showerror("¡Error!","No seleccionó nada!")

        parent.update_buttons()

    def browse_file_pvp(self,parent):
        # Browse for a file and validate it starts with 'PVP' to be considered a valid PVP file.
        ruta_pvp = filedialog.askopenfilename()
        archivo_pvp = ruta_pvp.split('/')[-1]
        if archivo_pvp.startswith('PVP'):
            self.entryPVPVar.set(ruta_pvp)
        else:
            self.entryPVPVar.set('')
            messagebox.showerror("¡Error!","Eso no es un PVP!")
        parent.update_buttons()

    def can_go_to_next_page(self):
        #Establishes if user can go to next page.
        return True if self.entryPVPVar.get()!='' and self.entryCarpetaVar.get() != '' else False
            
class SecondFrame(ttk.Frame):

    def __init__(self,parent):
        super().__init__(parent)

    def create_widgets(self,parent):

        #Frames
        
        titulo=parent.carpeta_proyecto.split('/')[-1]
        self.frameIzquierdo=ttk.LabelFrame(self,text=titulo,padding=(20, 20))
        self.frameDerecho= ttk.LabelFrame(self,text='',padding=(20, 20))
 

        #TreeView
        self.crear_tabla_comparativa(parent)

        #Frame derecho/Pregunta para continuar
        self.confirmacion_final(parent)
        #Frame Derecho/Boton fichas tecnicas
        self.boton_fichas_tecnicas(parent)

        
        self.place_widgets()

    def place_widgets(self):
        #SElf frame principal
        self.frameIzquierdo.grid(row=0,column=0,sticky='nsew',pady=(30,30),padx=(30,30))
        self.frameDerecho.grid(row=0,column=1)
        #frame izquierdo
        self.comparacion.grid(row=0,column=0,columnspan=2,pady=(10,15))        
        
    def comparar_sp_vs_pvp(self,sp,pvp):
        #Extrae la moneda del SP
        wb=openpyxl.load_workbook(sp, read_only=True)
        ws=wb.worksheets[0]
        moneda=ws['E17'].value
        wb.close()

        df_sp=oferta.dataframe_sp(sp)
        self.df_pvp,totales=oferta.dataframe_pvp(pvp)
        tabla_comparativa=oferta.generar_tabla_comparativa(df_sp,self.df_pvp,moneda)

        return (tabla_comparativa,totales,moneda)
    
    def crear_tabla_comparativa(self,parent):

        self.comparacion=ttk.Treeview(self.frameIzquierdo,show="headings")
        

        for i in self.comparacion.get_children():
            self.comparacion.delete(i) 

        tablacomparativa=parent.tablacomparativa
        df=tablacomparativa[0]
        totales=tablacomparativa[1]
        moneda=tablacomparativa[2]

        for i, (key, value) in enumerate(totales.items(),1):
            ttk.Label(self.frameIzquierdo, text=key).grid(row=i,column=0,pady=(0,10),padx=(10,10),sticky='e')
            ttk.Label(self.frameIzquierdo, text=moneda+f" {value:.2f}").grid(row=i,column=1,pady=(0,10),padx=(0,10),sticky='w')

        columns = list(df.columns)
        self.comparacion['columns'] = columns

        for col in columns:
            self.comparacion.heading(col, text=col)
            self.comparacion.column(col, width=100,anchor='center')

        self.comparacion.column('NOMBRE',width=60)
        self.comparacion.column('REFERENCIA',width=78)
        self.comparacion.column('CANTIDAD',width=70)
        self.comparacion.column(3,width=110)
        self.comparacion.column(4,width=110)
        self.comparacion.column(5,width=75)


        # Insertar datos del DataFrame al TreeView
        for index, row in df.iterrows():
            self.comparacion.insert("", tk.END, values=list(row))

    def confirmacion_final(self,parent):
        ttk.Label(self.frameDerecho,text='Está de acuerdo con el costeo?').grid(row=0,column=0,sticky='nsew',pady=20)
        
        self.selected_option = tk.IntVar()

        radio1=ttk.Radiobutton(self.frameDerecho,text='SÍ',variable=self.selected_option,value=1,command=parent.update_buttons)
        radio2=ttk.Radiobutton(self.frameDerecho,text='NO',variable=self.selected_option,value=0,command=parent.update_buttons)
        radio1.grid(row=1,column=1,sticky='nsew',pady=6,padx=5)
        radio2.grid(row=1,column=2,sticky='nsew',pady=6,padx=5)

    def canFinish(self):
        return True if self.selected_option.get()==1 else False
    
    def boton_fichas_tecnicas(self,parent):
        boton=ttk.Button(self.frameDerecho,text='Buscar Fichas Técnicas',command=lambda x=parent:self.fichas_tecnicas(x))
        boton.grid(row=2,column=0,columnspan=3,sticky='ns',padx=30,pady=50)

    def fichas_tecnicas(self,parent):
        import fichas_tecnicas

        carpeta_fichas=os.path.join(parent.carpeta_proyecto,'FICHAS_TECNICAS')
        encontradas,totales=fichas_tecnicas.main(carpeta_fichas,parent.carpeta_proyecto)
        messagebox.showinfo('Atención',f'Se encontraron {encontradas} fichas técnicas de {totales}.')



App("Crear Oferta",(1000,500))






