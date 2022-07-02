import tkinter as tk 
from tkinter import Button, Label, messagebox
from tkinter.font import NORMAL         # messagebox модуль который содержит функци и классы по работе со всплывающими окнами и сообщениями
import tkinter.filedialog as fd
from idlelib.tooltip import Hovertip           # всплывающие окна
import openpyxl                   # need for work with Excel( especially .xlsx)
from datetime import date         # need to create today-date.


def createNewWindow():
    newWin = tk.Toplevel(win)
    newWin.geometry( f'350x400+200+210' )
    #photo = tk.PhotoImage(file='Icon.png')
    #newWin.iconphoto(False, photo)
    newWin['bg'] = '#9ba29b'
    newWin.title('Create_Table_SQL')
    newWin.resizable(True, True)
    tk.Label(newWin, font='PFDINDISPLAYPRO 10', text='Для корректного формирования кода \n ознакомтесь с правилами, \n которые необходимо соблюсти:', 
                    bg='#9ba29b', wraplength=200).pack()
    tk.Label(newWin, font='PFDINDISPLAYPRO 10', text='- в таблице Excel не должно быть объедененных полей; \n \
- в таблице Excel не дожно бы пустых ячеек(если в ячейче должно быть пустое значение, то используйте Null); \n \
- в таблице Excel должны быть указаны только наименование колонок, без объединяющих шапок;\n \
- в данной программе тип данных DATA имеет вид YYYY-MM-DD; \n \
- обратите внимание, на то что помимо пути к файлу мы указываем и наименование листа в Excel, где лижит еобходимая таблица', 
                    bg='#9ba29b', wraplength=200).pack()
    
                    


def open_file():                          # select path in file
    file_path = fd.askopenfilename()
    Entry_num.delete(0, tk.END)
    Entry_num.insert(0, file_path)
    print( 'открываем файл:', file_path )


def step_2_create_table_base():
    filename = str(Entry_num.get())
    wb = openpyxl.load_workbook(filename)                        #example for path: wb = openpyxl.load_workbook(filename = 'C:/Users/user/Desktop/Учеба МАИ/Pyton/1/Project/SQL_create_table/table_b2.xlsx')
    sheet = wb[Entry_num_name_sheet.get()]
    name_table = filename.split('/')[-1]

    list_excel = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    final_table = []
    for i in range( len(list_excel)):    # find first block in sheet
        for j in range(1, 20):
            if  sheet[list_excel[i]+str(j)].value != None:
                x_first = i     # !!! type NUMERIC
                y_first = j     # !!! type NUMERIC
                break
        if  sheet[list_excel[i]+str(j)].value != None:
            break
            
    # through print() we duplicate information in the console
    print( 'CREATE TABLE', name_table,'(', end = '\n')
    # create a variable a where we will collect in mass 'final_table' our text and then paste it into the tk.text
    a = 'CREATE TABLE'+' ' + name_table +'('
    final_table.append(a)
    
    print('   ', name_table, 'INT PRINMARY KEY AUTO_INCREMENT,', end ='\n' )
    a = '   ' +' '+ name_table +' '+ 'INT PRINMARY KEY AUTO_INCREMENT,'
    final_table.append(a)
    # create dict, were name_coloumn = type_in_SQL
    tabel_type_dict = {}
    # understand type column, fill out 'final_table' and 'tabel_type_dict', then CREATE TABLE
    # replace the value path with variables to reduce the code
    step = 0
    while sheet[list_excel[x_first+step]+str(y_first+1)].value != None:  
        first_val_sheet = sheet[list_excel[x_first+step] + str(y_first+1)].value
        sec_val_sheet = sheet[list_excel[x_first] + str(y_first+2)].value
        if (str(first_val_sheet)).isdigit() and (str(sec_val_sheet)).isdigit():          # test type to INT
            print( '    ' + str(sheet[list_excel[x_first+step] + str(y_first)].value), 'INT,', end ='\n' ) 
            a = '    ' + str(sheet[list_excel[x_first+step] + str(y_first)].value) + ' ' + 'INT,'
            final_table.append(a)
            tabel_type_dict[str(sheet[list_excel[x_first+step] + str(y_first)].value)] = 'INT'
    
        elif len(((str(first_val_sheet))).split('-')) == 3 and len(((str(first_val_sheet))).split('-')[0]) == 4 and (((str(first_val_sheet))).split('-')[0]).isdigit() and (((str(first_val_sheet))).split('-')[1]).isdigit()  and (((str(first_val_sheet))).split('-')[2]).isdigit():
            print( '    ' + str(sheet[list_excel[x_first+step] + str(y_first)].value), 'DATE,', end ='\n' ) 
            a = '    ' + str(sheet[list_excel[x_first+step] + str(y_first)].value) + ' ' + 'DATE,'
            final_table.append(a)
            tabel_type_dict[str(sheet[list_excel[x_first+step] + str(y_first)].value)] = 'DATE'
    
        elif (''.join((str(first_val_sheet)).split(',')).isdigit() and ''.join((str(sec_val_sheet)).split(',')).isdigit()) :                         # test type to DEMICAL
            len_demical = 0
            temp = int((first_val_sheet).split(',')[0])
            while temp > 0:
                temp //= 10
                len_demical += 1
            print( '    ' + str(sheet[list_excel[x_first+step] + str(y_first)].value)+ ' ' + 'DECIMAL(' + str(len_demical + 3) + ', 2' + '),', end ='\n' ) 
            a = '    ' + str(sheet[list_excel[x_first+step] + str(y_first)].value)+ ' ' + 'DECIMAL(' + str(len_demical + 3) + ', 2' + '),'
            final_table.append(a)
            tabel_type_dict[str(sheet[list_excel[x_first+step] + str(y_first)].value)] = 'DECIMAL'
    
        elif (str(first_val_sheet)).isalnum() or (str(sec_val_sheet)). isalnum() or (str(first_val_sheet)).isalpha() or (str(sec_val_sheet)). isalpha():                         # test type to STR
            step_column = 0
            len_max = 0
            while (sheet[list_excel[x_first+step] + str(y_first+1+step_column)].value) != None:                       
                if  len(str((sheet[list_excel[x_first+step] + str(y_first+1+step_column)].value))) > len_max:
                    len_max = len(str((sheet[list_excel[x_first+step] + str(y_first+1+step_column)].value)))
                step_column += 1
            print( '    ' + str(sheet[list_excel[x_first+step] + str(y_first)].value)+ ' ' + 'VARCHAR('+ str(len_max + 10) + '),', end ='\n' ) 
            a = '    ' + str(sheet[list_excel[x_first+step] + str(y_first)].value)+ ' ' + 'VARCHAR('+ str(len_max + 10) + '),'
            final_table.append(a)
            tabel_type_dict[str(sheet[list_excel[x_first+step] + str(y_first)].value)] = 'VARCHAR'
        step += 1
        
    
    
    if Chek_step_date_create.get() == 1:
        tabel_type_dict['DATE_CREATE'] = 'DATE'
        print( '    DATE_CREATE DATE', end = '\n' )
        a = '    DATE_CREATE DATE'
        final_table.append(a)
    else:
        final_table[-1]= final_table[-1][:-1]

    print( ');', end = '\n' )
    a = ');'
    final_table.append(a)
    # join our final_table and insert in tk.text
    text.insert(0.1, '\n'.join(final_table))
    # zero out the final_table values to create the code for filling the table
    a = ''
    final_table = []
    # end create TABLE
   
    # start to add values in Table
    matrix_values = []
    column_values = []
    step_for_values_y = 0
    step_for_values_x = 0
    
    while sheet[list_excel[x_first+step_for_values_x]+str(y_first+step_for_values_y)].value != None:
        name_coloumn = sheet[list_excel[x_first+step_for_values_x]+str(y_first)].value
        if tabel_type_dict[name_coloumn] == 'INT':
            column_values.append(str(sheet[list_excel[x_first+step_for_values_x]+str(y_first+step_for_values_y)].value))
        elif tabel_type_dict[name_coloumn] == 'DECIMAL':
            if sheet[list_excel[x_first+step_for_values_x]+str(y_first+step_for_values_y)].value != name_coloumn:
                column_values.append('.'.join((str(sheet[list_excel[x_first+step_for_values_x]+str(y_first+step_for_values_y)].value)).split(',')))
            else: 
                column_values.append(str(sheet[list_excel[x_first+step_for_values_x]+str(y_first+step_for_values_y)].value))
        elif tabel_type_dict[name_coloumn] == 'DATE':
            column_values.append(str(sheet[list_excel[x_first+step_for_values_x]+str(y_first+step_for_values_y)].value))
            
        elif tabel_type_dict[name_coloumn] == 'VARCHAR':
            column_values.append('"' + str(sheet[list_excel[x_first+step_for_values_x]+str(y_first+step_for_values_y)].value) + '"')

        step_for_values_y += 1
        if sheet[list_excel[x_first+step_for_values_x]+str(y_first+step_for_values_y)].value == None:
            step_for_values_x += 1
            step_for_values_y = 0
            matrix_values.append(column_values[1:])
            column_values = []
    
    print( 'INSERT INTO' + ' ' + name_table + '(' + ', '.join([k for k in tabel_type_dict]) + ')', end = '\n')
    a = 'INSERT INTO' + ' ' + name_table + '(' + ', '.join([k for k in tabel_type_dict]) + ')'
    final_table.append(a)
    print( 'VALUES ' , end = '\n' )
    a = 'VALUES '
    final_table.append(a)
    
    column_values = []
    for i in range(len(matrix_values[0])):
        for j in range(len(matrix_values)):
            column_values.append(matrix_values[j][i])      
        if i == len(matrix_values[0])-1:
            if Chek_step_date_create.get() == 1:
                print( '(' + ', '.join(column_values) + ', ' + '.'.join(str(date.today()).split('-')) + ')'+';', end = '\n' )
                a = '(' + ', '.join(column_values) + ', ' + str(date.today()) + ')'+';'
                final_table.append(a)
                column_values = []
            else:
                print( '(' + ', '.join(column_values) + ')'+';', end = '\n' )
                a = '(' + ', '.join(column_values) + ')'+';'
                final_table.append(a)
                column_values = []
        else: 
            if Chek_step_date_create.get() == 1:
                print( '(' + ', '.join(column_values) + ',  ' + '.'.join(str(date.today()).split('-')) + ')' + ',', end = '\n' )
                a = '(' + ', '.join(column_values) + ', ' + str(date.today()) + ')'+','
                final_table.append(a)
                column_values = []  
            else:    
                print( '(' + ', '.join(column_values) + ')'+',', end = '\n' )
                a = '(' + ', '.join(column_values) + ')'+','
                final_table.append(a)
                column_values = []
 
    text2.insert(0.1, '\n'.join(final_table))
    a = None
    final_table = []

# THIS CODE IS FOR CREATING AN ADDITIONAL Lable. MAYBE USEFUL IN MODIFICATION
#def add_field(n):               # func for add zoneEntry
#    from tkinter import *
#    l=[]
#    n=0
#    root=Tk()
#    root.geometry('600x600')
#    btn_add=Button(root,command=lambda:add_field(n))
#    def add_field(n):
#        s = StringVar()
#        Entry(root, textvariable=s).pack()
#        l.append(s)
#    btn_add.pack()
#    def print_text():
#        for t in l:
#            if t.get() != '':
#                print(t.get())
#    btn_pr=Button(root,command=print_text)
#    btn_pr.pack()
#    mainloop()

def copy_to_text():
    win.clipboard_clear()  # Очистить буфер обмена
    win.clipboard_append(text.get('1.0', tk.END).rstrip())


def clear_text():                          
    text.delete("1.0","end")
    print( 'Очищаем поле' )


def copy_to_text2():
    win.clipboard_clear()  # Очистить буфер обмена
    win.clipboard_append(text2.get('1.0', tk.END).rstrip())


def clear_text2():                          
    text2.delete("1.0","end")
    print( 'Очищаем поле' )







win = tk.Tk()
win.geometry( f'470x750+150+200' )
#photo = tk.PhotoImage(file='Icon.png')
#win.iconphoto(False, photo)
win['bg'] = '#9ba29b'
win.title('Create_Table_SQL')
win.resizable(True, True)

#-------------------------------------------------------------------------------------------------

#win.bind('<Key>', press_key)        #bind - обработчик событий. Тут привязываем кнопки к цифрам на клаве. <Key> - событие срабатывания при нажатии на клавишу, press_key - выполняемая функция
 
 
Title_head_1 = tk.Label(win, text='Создание и наполнение', font='PFDINDISPLAYPRO 14 bold', bg ='#9ba29b')
Title_head_1.grid(row=0, column=1,  columnspan=5, stick='wesn')
Title_head_1 = tk.Label(win, text='таблицы в SQL', font='PFDINDISPLAYPRO 14 bold', bg ='#9ba29b')
Title_head_1.grid(row=1, column=1,  columnspan=5, stick='wens')

#===================================== STEP 1 =================================================================================
step_1_text_1_pass = tk.Label(win, text='', font='PFDINDISPLAYPRO 14', bg ='#9ba29b')
step_1_text_1_pass.grid(row=2, column=0,  columnspan=1)

step_1_text_1 = tk.Label(win, text='Шаг 1.', font='PFDINDISPLAYPRO 12 bold', bg ='#9ba29b')
step_1_text_1.grid(row=3, column=0,  columnspan=1)
step_1_text_2 = tk.Label(win, text='Выберети файл Elcel, с необходимой таблицей данных', font='PFDINDISPLAYPRO 10', bg ='#9ba29b')
step_1_text_2.grid(row=3, column=1,  columnspan=6, stick='w')


tk.Button(win, text = 'Обзор', bd = 4, font=('PFDINDISPLAYPRO', 12), command=open_file, bg = '#ffffff', relief='raised', 
          activebackground ='#d7db00').grid(row = 4, column=0, stick='we', padx = 25, pady = 5, columnspan=1) 

Entry_num = tk.Entry(win, font = 'PFDINDISPLAYPRO 10', width=(2), justify=tk.CENTER, bd = 4)
Entry_num.grid( row=4, column=2, stick='we', padx = 5, pady = 5, columnspan=4)
Entry_num.insert(0, 'путь к файлу Excel')
#Entry_num['state'] = tk.DISABLED

step_1_text_3 = tk.Label(win, text='Укажите имя листа:', font='PFDINDISPLAYPRO 12', bg ='#9ba29b')
step_1_text_3.grid(row=5, column=0,  columnspan=2, stick='', padx = 25, pady = 0)

Entry_num_name_sheet = tk.Entry(win, font = 'PFDINDISPLAYPRO 10', width=(2), justify=tk.CENTER, bd = 4)
Entry_num_name_sheet.grid( row=5, column=2, stick='we', padx = 5, pady = 5, columnspan=2)
Entry_num_name_sheet.insert(0, 'Sheet1')


step_1_text_3 = tk.Label(win, text='', font='PFDINDISPLAYPRO 12', bg ='#9ba29b')
step_1_text_3.grid(row=6, column=0,  columnspan=2, stick='')


#===================================== STEP 2 ============================================================================

step_2_text_1 = tk.Label(win, text='Шаг 2.', font='PFDINDISPLAYPRO 12 bold', bg ='#9ba29b')
step_2_text_1.grid(row=7, column=0,  columnspan=1)
step_2_text_2 = tk.Label(win, text='Код для создания таблицы в SQL', font='PFDINDISPLAYPRO 10', bg ='#9ba29b')
step_2_text_2.grid(row=7, column=1,  columnspan=6, stick='w')



Create_button = tk.Button(win, text = 'Создать ', bd = 4, font=('PFDINDISPLAYPRO', 12), 
                          command = step_2_create_table_base, bg = '#ffffff', activebackground ='#d7db00' )
Create_button.grid(row = 8, column=0, stick='we', padx = 25, pady = 5, columnspan=1) 




Chek_step_date_create = tk.IntVar(win)
step_2_chose_create_base = tk.Checkbutton(win, text = "Добавить поле 'Дата загрузки'", font='PFDINDISPLAYPRO 12',  
                                          bg ='#9ba29b', variable=Chek_step_date_create, offvalue=0, onvalue=1, indicatoron=1)
step_2_chose_create_base.grid(row=9, column=0,  columnspan=4, stick='w', padx = 25, pady = 5)

Rulle_create_table = tk.Button(win, text = ' ?! ', bd = 4, font=('PFDINDISPLAYPRO', 12), bg = '#ffffff', activebackground ='#d7db00', command=createNewWindow )
Rulle_create_table.grid(row = 8, column=1, stick='w', padx = 1, pady = 5, columnspan=1) 
Hovertip(Rulle_create_table, "Во избежании ошибок, ознакомьтесь  \n с критериями исходной таблицы", hover_delay=250)

#===================================== STEP 3 Result ============================================================================

btn_del_text = Button(win, text = 'Clear', font=('PFDINDISPLAYPRO', 13), bd = 4, command=clear_text, bg = '#ffffff', 
                       activebackground ='#d7db00').grid(row = 10, column=6, stick='we', padx = 25, pady = 5, columnspan=1) 
btn_save_text = Button(win, text = 'Copy', font=('PFDINDISPLAYPRO', 13), bd = 4, command=copy_to_text, bg = '#ffffff', 
                       activebackground ='#d7db00').grid(row = 10, column=5, stick='w', padx = 1, pady = 5, columnspan=1) 

text = tk.Text(win, width=80, height=150, wrap='word')
text.grid(row=11, column=0,  columnspan=10, padx=25, pady=1, stick='nwes', rowspan=1)

scroll = tk.Scrollbar(text, command=text.yview)
scroll.pack(side= 'right', fill= 'y')

text.config(yscrollcommand=scroll.set)



btn_del_text2 = Button(win, text = 'Clear', font=('PFDINDISPLAYPRO', 13), bd = 4, command=clear_text2, bg = '#ffffff', 
                       activebackground ='#d7db00').grid(row = 12, column=6, stick='we', padx = 25, pady = 5, columnspan=1) 
btn_save_text2 = Button(win, text = 'Copy', font=('PFDINDISPLAYPRO', 13), bd = 4, command=copy_to_text2, bg = '#ffffff', 
                        activebackground ='#d7db00').grid(row = 12, column=5, stick='w', padx = 1, pady = 5, columnspan=1) 

text2 = tk.Text(win, width=80, height=150, wrap='word')
text2.grid(row=13, column=0,  columnspan=10, padx=25, pady=1, stick='nwes', rowspan=5)

scroll2 = tk.Scrollbar(text2, command=text2.yview)
scroll2.pack(side= 'right', fill= 'y')

text2.config(yscrollcommand=scroll2.set)





win.grid_rowconfigure(11, minsize='150')

win.grid_rowconfigure(13, minsize='150')



win.mainloop()