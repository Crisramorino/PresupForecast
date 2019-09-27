from Conect2sqlsrv import conn

from tkinter import *
from tkinter import ttk
import math
import decimal

#Parametros del programa
SQL_table='SigmaGestion.Gestion.PresupForecast_copy'
sql_col_names = "select column_name from information_schema.columns where table_name = 'PresupForecast'"
n_pag=0
n_rows=10          #Número de registros a mostrar por página.

"""Funciones y comandos de GUI"""
"""**************************************************************************"""
def centering_win(rt,w,h):
    #w width for the Tk root
    #h height for the Tk root

    # get screen width and height
    ws = rt.winfo_screenwidth() # width of the screen
    hs = rt.winfo_screenheight() # height of the screen

    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    rt.geometry('%dx%d+%d+%d' % (w, h, x, y))
    return rt
def isnumeric(s):
    try:
        float(s)
        return True
    except ValueError:
        return False
#add_margin se encarga de añadir margenes a todos los widgets de un frame.
def add_margin(frame):
        for child in frame.winfo_children(): child.grid_configure(padx=5, pady=5)

def Get_Query():
#Aquí se define el procedimiento que se ejecuta para filtrar la datos a partir de un
#centro de costo y un numero que represente un año mes. El usuario ingresa manalmente
#estos campos que son considerados en el WHERE de la consulta SQL.

#Por otra parte, Get_Query realiza un a validación de datos para que la consulta
#sea ejecutada de manera correcta.
    global dataset,n_pag
    n_pag=0

    sql="SELECT * FROM " + SQL_table + ' WHERE activo=1'

    cc_get=CC.get()
    am_get=Ano_mes.get()

    lcc=len(cc_get)
    lam=len(am_get)

    if lcc != 0 or lam != 0:
        sql += " AND "

        if lcc != 0:
            sql += "CodiCC = '"+cc_get +"'"

        if lcc != 0 and lam != 0:
            sql += " AND "

        if lam !=0:

            if not str(Ano_mes.get()).isnumeric():

                alert_ano_mes_w = Tk()
                alert_ano_mes_f = ttk.Frame(alert_ano_mes_w,padding="3 3 12 12")
                alert_ano_mes_f.grid(sticky=(N,S,W,E))
                ttk.Label(alert_ano_mes_f,text="Año mes indicado no es válido.").grid(column=0,row=0)
                ttk.Button(alert_ano_mes_f,text="OK",command=alert_ano_mes_w.destroy).grid(column=0,row=1)
                return []
            sql += "AnoMes = '"+ am_get +"'"

    Values = get_values()
    distinct = compare_values(Values)
    modif_alert(distinct)
    cursor.execute(sql)

    dataset=cursor.fetchall()
    if len(dataset)==0:
        alert_w = Tk()
        alert_f = ttk.Frame(alert_w,padding="12 3 12 12")
        alert_f.grid(sticky=(N,S,W,E))
        ttk.Label(alert_f,text="No se encontraron registros para los criterios definidos.").grid(column=0,row=0)
        ttk.Button(alert_f,text="OK",command=alert_w.destroy).grid(column=0,row=1)
        # w=alert_w.winfo_width()
        # print(w)
        # h=alert_w.winfo_height()
        # alert_w=centering_win(alert_w,w,h)
        return []

    return dataset

def Show_Dataset(first_row=0,show_n=n_rows):


    global dataset, col_names, Original_data, Entries


    for child in DataFrame.winfo_children(): child.destroy()

    for j in range(len(col_names)-2):
        if not j in [1, 4,6, 7]:
            if j==2:
                ttk.Label(DataFrame,text='C.C.').grid(column=j, row=0, sticky=W)
            else:
                ttk.Label(DataFrame,text=col_names[j]).grid(column=j, row=0, sticky=W)

    Original_data=[]

    Entries=['']*min(show_n,len(dataset)-first_row)

    for i in range(min(show_n,len(dataset)-first_row)):

        Entries[i] = StringVar()
        sqlrow=dataset[i+first_row]



        for j in range(len(sqlrow)-3):
            if not j in [1, 4,6, 7]:

                ttk.Label(DataFrame,text=sqlrow[j]).grid(column=j, row=i+1, sticky=W)

        str_valor= '{0:,.0f}'.format(sqlrow[len(sqlrow)-3]).replace(',','.')

        Original_data.append(str_valor.replace('.',''))
        Valor=ttk.Entry(DataFrame,width=10, textvariable=Entries[i])
        # Valor[i].bind('<Leave>',lambda args: Valor[i].Delete(0,END))
        # Valor[i].bind('<Leave>',lambda *args: Valor[i].insert(0,'{0:,.0f}'.format(int(Entries[i].get().replace('.','')))))

        # Valor.bind('<KeyRelease>',lambda arg: print("hola"))
        Valor.insert(0,str_valor)
        Valor.grid(column=9, row=i+1,sticky=W,columnspan=2)
        add_margin(DataFrame)

def QueryExec(*args):

    global dataset
    dataset=Get_Query()
    Show_Dataset()

def save_regs():
    Values = get_values()
    if  Values:
        distinct = compare_values(Values)
        modif_alert(distinct)
    Show_Dataset(first_row=n_rows*n_pag)

def next(*arg):
    global n_pag
    n_dataset=len(dataset)
    n_pag+=1
    n_pag=min(n_pag,math.floor(n_dataset/n_rows))
    num_pag.config(text='pag: '+str(n_pag+1))
    Values = get_values()
    if  Values:

        distinct = compare_values(Values)
        modif_alert(distinct)
    Show_Dataset(first_row=n_rows*n_pag)

def previous(*arg):
    global n_pag
    n_pag-=1
    n_pag=max(n_pag,0)
    num_pag.config(text='pag: '+str(n_pag+1))
    Values = get_values()
    if not Values:
        distinct = compare_values(Values)
        modif_alert(distinct)
    Show_Dataset(first_row=n_rows*n_pag)

def add_records(win):

    insert_sql="INSERT INTO "+SQL_table+"("+col_names[1]+","
    insert_sql+=col_names[2]+","+col_names[3]+","+col_names[4]+","
    insert_sql+=col_names[5]+","+col_names[6]+","+col_names[7]+","
    insert_sql+=col_names[8]+","+col_names[9]+","+col_names[10]+","+col_names[11]+")"

    update_sql ="UPDATE "+SQL_table+" SET "+col_names[8]+"='NO VIGENTE', "+col_names[10]+"=0 WHERE "

    data2insert=[]

    sql_array="("

    for i in distinct:
        data2insert=[str(x) for x in dataset[i]]

        insert_sql_full = insert_sql + " VALUES ("+data2insert[1]+",'"
        insert_sql_full += data2insert[2]+"','"+data2insert[3]+"',"+data2insert[4]+","
        insert_sql_full += data2insert[5]+",'"+data2insert[6]+"','"+data2insert[7]+"','"
        insert_sql_full += data2insert[8]+"',"+Values[i]
        insert_sql_full += ",1, CURRENT_TIMESTAMP " +")"
        print(insert_sql_full)
        cursor.execute(insert_sql_full)
        conn.commit()

        sql_array+=data2insert[0]+","

    sql_array=sql_array[:len(sql_array)-1]

    update_sql+=col_names[0]+" IN "+sql_array+")"

    cursor.execute(update_sql)
    conn.commit()

    win.destroy()
    Show_Dataset(first_row=n_rows*n_pag)

def dont_add_modif(win):
    win.destroy()
    Show_Dataset(first_row=n_rows*n_pag)

def modif_alert(distinct):

    if len(distinct):

        alert=Tk()
        alert.title('Modificación de datos.')
        msg_frame=ttk.Frame(alert, padding="3 3 12 12")
        msg_frame.grid(column=0,row=0,sticky=(N, W, E, S))
        msg=ttk.Label(msg_frame,text='se han modificado algunos valores de para algunos centros costo \n ¿Desea aplicar estos cambios?')
        msg.grid(column=0,row=0,columnspan=2,sticky=(N, W, E, S))

        ok_butt = ttk.Button(msg_frame,text='Si',command=lambda : add_records(alert))
        ok_butt.grid(column=0,row=1,sticky=E)
        no_butt = ttk.Button(msg_frame,text='No',command=lambda : dont_add_modif(alert))
        no_butt.grid(column=1,row=1,sticky=W)
        # ok_butt.bind('<1>',lambda e:alert.destroy())
        alert.mainloop()
        add_margin(msg_frame)

    else:
        Show_Dataset(first_row=n_rows*n_pag)

def no_numeric_value():
    alert_val_w = Tk()
    alert_val_f = ttk.Frame(alert_val_w,padding="3 3 12 12")
    alert_val_f.grid(sticky=(N,S,W,E))
    ttk.Label(alert_val_f,text="Se han ingresado valores no numéricos, corregir para continuar.").grid(column=0,row=0)
    ttk.Button(alert_val_f,text="OK",command=alert_val_w.destroy).grid(column=0,row=1)
    return 0

def get_values():
    global Values
    Values=[]
    for E in Entries:
        val=E.get().replace('.','')

        if not isnumeric(val):
            no_numeric_value()
            return False
            break
        else:
            Values.append(val)
    return Values

def compare_values(Values):
    global distinct
    distinct=[]
    for i in range(len(Values)):
        if Values[i]!=Original_data[i]:
            distinct.append(i)

    return distinct

sql='select * from ' + SQL_table +' where activo=1'

cursor = conn.cursor()  #Se instancia una clase que realizará las consultas a la base de datos
cursor.execute(sql_col_names)
pyodbc_col = cursor.fetchall()
col_names = [x[0] for x in pyodbc_col]
cursor.execute(sql)

dataset=cursor.fetchall()
root = Tk()
root.title("Presupuesto Forecast")

"""Configuración del "Query Frame", en este contenedor se definen los campos de
 consulta por centro de costo y periodo"""
"""**************************************************************************"""
QueryFrame=ttk.Frame(root, padding="3 3 12 12")
QueryFrame.grid(column=0, row=0, sticky=(N, W, E, S))

image = PhotoImage(file='Sigma.png')

Logo_Sigma = ttk.Label(QueryFrame,image=image)
Logo_Sigma.grid(column=1,row=0, columnspan=2)

CC = StringVar()

Ano_mes = StringVar()

cc_label = ttk.Label(QueryFrame,text='Centro de Costo:')
cc_label.grid(column=1, row=1)
cc_entry = ttk.Entry(QueryFrame,width=7, textvariable = CC)
cc_entry.grid(column=2, row=1)

anomes_label = ttk.Label(QueryFrame,text='Año y Mes:')
anomes_label.grid(column=4, row=1)
ano_mes_entry = ttk.Entry(QueryFrame,width=7, textvariable=Ano_mes)
ano_mes_entry.grid(column=5, row=1)

QueryExec_button = ttk.Button(QueryFrame, text='Consultar',command=QueryExec)
QueryExec_button.grid(column=6, row=1)
add_margin(QueryFrame)
"""Configuración del "Data Frame", en este contenedor se despliegan los datos de
 la base de datos"""
"""**************************************************************************"""
DataFrame = ttk.Frame(root)
DataFrame.grid(column=0, row=1, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

Show_Dataset()



"""Configuración del paginador"""
"""**************************************************************************"""

PaginatorFrame = ttk.Frame(root, padding="3 3 12 12")
PaginatorFrame.grid(column=0, row=2, sticky=(N, W, E, S))

save_reg = ttk.Button(PaginatorFrame,text="Guardar\nRegistros",command=save_regs)
save_reg.pack(side=TOP, expand=YES)

num_pag=ttk.Label(PaginatorFrame,text='pag: '+str(n_pag+1))
num_pag.pack(side=TOP, expand=YES)
prev_butt=ttk.Button(PaginatorFrame,text='<',command=previous)
prev_butt.pack(side=LEFT, expand=YES)

next_butt=ttk.Button(PaginatorFrame,text='>',command=next)
next_butt.pack(side=RIGHT, expand=YES)


#feet_entry.focus()
root.bind('<Return>', QueryExec)

root.mainloop()
