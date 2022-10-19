import pandas as pd
import datetime as dt
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox,tkinter.filedialog

root = tk.Tk()
root.title('excel small tool')
root.geometry('600x200')
fileaddr = ''
sheet_names = []
bl_executed = 0

values = ["列A", "列B", "列C", "列D", "列E", "列F", "列G", "列H", "列I", "列J", "列K", "列L", "列M", "列N", "列O", "列P"]
def getindex(ind_cont):
    for j in range(len(values)):
        if ind_cont == values[j]:
            return j

def btn_execute_pressup():
    global bl_executed
    bir_col = getindex(combobox1.get())
    cri_col = getindex(combobox2.get())

    if bir_col != '' and cri_col != '' and fileaddr != '' and bl_executed == 1:
        dfall = pd.read_excel(fileaddr, sheet_name=None)
        writer = pd.ExcelWriter(fileaddr,engine='openpyxl')
        try:
            for sheetname,df in dfall.items():
                adult = []
                for i in range(len(df)):
                    id_no = df.values[i,bir_col]
                    year = int(id_no[6:10])
                    month = int(id_no[10:12])
                    day = int(id_no[12:14])
                    if ((year % 4 == 0 and year % 100 != 0) or year % 400 == 0) and month ==2 and day == 29:
                        bir_day = dt.date(year+18,3,1)
                    else:
                        bir_day = dt.date(year + 18, month, day)
                    crime_day = dt.datetime.strptime(df.values[i, cri_col],'%Y-%m-%d %H:%M:%S').date()
                    duration = crime_day - bir_day
                    if duration.days >= 0:
                        adult.append('是')
                    else:
                        adult.append('否')
                df.insert(df.shape[1],'是否成年',adult)

                df.to_excel(writer,sheet_name=sheetname,engine='openpyxl')
                # file_dir = os.path.dirname(fileaddr)
                # file_name = os.path.basename(fileaddr)
                # temp_name = os.path.join(file_dir,sheetname+file_name)
                # df.to_excel(temp_name)
                # writer = pd.ExcelWriter(fileaddr, engine_kwargs='openpyxl')
                # book = load_workbook(writer.path)
                # writer.book = book
                # df.to_excel(excel_writer=writer, sheet_name=f'{sheetname}copy')
            writer.save()
            bl_executed = 2
            tk.messagebox.showinfo('congratulations', 'calculate ages succeed!')
        except Exception as e:
            tk.messagebox.showerror('error occured', 'have you selected the wrong column?')
    elif bl_executed == 2:
        tk.messagebox.showwarning('warning',"you have executed once,do not try it again")

def btn_select_pressup():
    global fileaddr
    global bl_executed
    bl_executed = 0
    fileaddr = tkinter.filedialog.askopenfilename()
    if fileaddr != '':
        try:
            d_name.set(fileaddr)
            sheets = pd.read_excel(fileaddr,sheet_name=None)
            combobox3['values'] = list(sheets)
            for sheet in sheets.keys():
                sheet_names.append(sheet)
            bl_executed = 1
        except:
            tk.messagebox.showerror('error','you have chosen the wrong file!')
    # id_no = str(df.values[0, 2])
    # bir_day = dt.date(int(id_no[6:10]) + 18, int(id_no[10:12]), int(id_no[12:14]))
    # crime_day = dt.datetime.date(df.values[0, 1])
    # duration = crime_day - bir_day

btn_execute = tk.Button(root,text='execute',command=btn_execute_pressup,height=1)
btn_execute.place(x=420,y=120,width=80)

btn_quit =tk.Button(root,text='quit',command=root.quit)
btn_quit.place(x=510,y=120,width=80)

btn_select = tk.Button(root,text='select excel file',command=btn_select_pressup,height=1)
btn_select.place(x=450,y=20)

d_name = tk.StringVar()
d_name.set("")
file_addr = tk.Entry(root, textvariable=d_name,state='readonly')
file_addr.place(x = 80, y = 20,height=30,width=350)

lbl_birth_name = tk.StringVar()
lbl_birth_name.set("身份证号码所属列：")
lbl_birth = tk.Label(root,textvariable=lbl_birth_name,
             fg='black',bg='white',font=('TimesNewRoman',8),
             justify='center',width=50,height=1)
lbl_birth.place(x=50,y=120,height=30,width=100)

lbl_crime_name = tk.StringVar()
lbl_crime_name.set("犯罪日期列：")
lbl_crime = tk.Label(root,textvariable=lbl_crime_name,
             fg='black',bg='white',font=('TimesNewRoman',8),
             justify='center',width=50,height=1)
lbl_crime.place(x=280,y=120,height=30,width=70)

lbl_sheet_name = tk.StringVar()
lbl_sheet_name.set("sheet名：")
lbl_sheet = tk.Label(root,textvariable=lbl_sheet_name,
             fg='black',bg='white',font=('TimesNewRoman',8),
             justify='center',width=50,height=1)
lbl_sheet.place(x=80,y=80,height=30,width=70)

combo1value = tk.StringVar()
combo1value.set("")

combobox1 = ttk.Combobox(master=root, height=10, width=20,state="readonly", cursor="arrow", font=("", 10),textvariable=combo1value,
values=values)
combobox1.place(x=150,y=120,height=30,width=50)
combobox1.current(6)

combo2value = tk.StringVar()
combo2value.set("")
# values = ["列A", "列B", "列C", "列D", "列E", "列F", "列G", "列H", "列I", "列J", "列K", "列L", "列M", "列N", "列O", "列P"]
combobox2 = ttk.Combobox(master=root, height=10, width=20,state="readonly", cursor="arrow", font=("", 10),textvariable=combo2value,
values=values, # 设置下拉框的选项
)
combobox2.place(x=350,y=120,height=30,width=50)
combobox2.current(9)

combo3value = tk.StringVar()
combo3value.set("")
values3 = []
# values = ["列A", "列B", "列C", "列D", "列E", "列F", "列G", "列H", "列I", "列J", "列K", "列L", "列M", "列N", "列O", "列P"]
combobox3 = ttk.Combobox(master=root, height=10, width=20,state="readonly", cursor="arrow", font=("", 10),textvariable=combo3value,
values=values3)
combobox3.place(x=160,y=80,height=30,width=150)

root.mainloop()



