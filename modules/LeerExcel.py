from tkinter import Tk, Label, Button, Frame, messagebox, filedialog, ttk, Scrollbar, VERTICAL, HORIZONTAL
import pandas as pd

def leerExcel():
	ventana2 = Tk()
	ventana2.title('Leer datos de Excel')
	
	width= ventana2.winfo_screenwidth()
	height= ventana2.winfo_screenheight()
	ventana2.geometry("%dx%d" % (width, height))

	ventana2.config(bg='black')
	ventana2.minsize(width=600, height=400)

	ventana2.columnconfigure(0, weight = 25)
	ventana2.rowconfigure(0, weight= 25)
	ventana2.columnconfigure(0, weight = 1)
	ventana2.rowconfigure(1, weight= 1)

	frame1 = Frame(ventana2, bg='gray26')
	frame1.grid(column=0,row=0,sticky='nsew')
	frame2 = Frame(ventana2, bg='gray26')
	frame2.grid(column=0,row=1,sticky='nsew')

	frame1.columnconfigure(0, weight = 1)
	frame1.rowconfigure(0, weight= 1)

	frame2.columnconfigure(0, weight = 1)
	frame2.rowconfigure(0, weight= 1)
	frame2.columnconfigure(1, weight = 1)
	frame2.rowconfigure(0, weight= 1)

	frame2.columnconfigure(2, weight = 1)
	frame2.rowconfigure(0, weight= 1)

	frame2.columnconfigure(3, weight = 2)
	frame2.rowconfigure(0, weight= 1)


	def abrir_archivo():
		archivo = filedialog.askopenfilename(initialdir ='/', 
												title='Selecione archivo', 
												filetype=(('xlsx files', '*.xlsx*'),('All files', '*.*')))
		indica['text'] = archivo

	def datos_excel():

		datos_obtenidos = indica['text']
		try:
			archivoexcel = r'{}'.format(datos_obtenidos)
			df = pd.read_excel(archivoexcel)

		except ValueError:
			messagebox.showerror('Informacion', 'Formato incorrecto')
			return None

		except FileNotFoundError:
			messagebox.showerror('Informacion', 'El archivo esta \n malogrado')
			return None

		Limpiar()

		tabla['column'] = list(df.columns)
		tabla['show'] = "headings"  #encabezado
		

		for columna in tabla['column']:
			tabla.heading(columna, text= columna)
		

		df_fila = df.to_numpy().tolist()
		for fila in df_fila:
			tabla.insert('', 'end', values =fila)


	def Limpiar():
		tabla.delete(*tabla.get_children())
	
	def cerrar():
		ventana2.destroy()
	
	tabla = ttk.Treeview(frame1 , height=10)
	tabla.grid(column=0, row=0, sticky='nsew')

	ladox = Scrollbar(frame1, orient = HORIZONTAL, command= tabla.xview)
	ladox.grid(column=0, row = 1, sticky='ew') 

	ladoy = Scrollbar(frame1, orient =VERTICAL, command = tabla.yview)
	ladoy.grid(column = 1, row = 0, sticky='ns')

	tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

	estilo = ttk.Style(frame1)
	estilo.theme_use('clam') #  ('clam', 'alt', 'default', 'classic')
	estilo.configure(".", font= ('Arial', 14), foreground='red2')
	estilo.configure("Treeview", font= ('Helvetica', 12), foreground='black',  background='white')
	estilo.map('Treeview',background=[('selected', 'green2')], foreground=[('selected','black')] )


	boton1 = Button(frame2, text= 'Abrir', bg='green2', command= abrir_archivo)
	boton1.grid(column = 0, row = 0, sticky='nsew', padx=5, pady=5)

	boton2 = Button(frame2, text= 'Mostrar', bg='magenta', command= datos_excel)
	boton2.grid(column = 1, row = 0, sticky='nsew', padx=5, pady=5)

	boton3 = Button(frame2, text= 'Limpiar', bg='red', command= Limpiar)
	boton3.grid(column = 2, row = 0, sticky='nsew', padx=5, pady=5)
	
	boton4 = Button(frame2, text= 'Cerrar ventana', bg='red', command= cerrar)
	boton4.grid(column = 3, row = 0, sticky='nsew', padx=5, pady=5)


	indica = Label(frame2, fg= 'white', bg='gray26', text= 'Ubicación del archivo', font= ('Arial',10,'bold') )
	indica.grid(column=4, row = 0)

	ventana2.mainloop()
