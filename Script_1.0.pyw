import os, os.path
from tkinter import *  
import shutil


dir_teml = "\\\\artsrv\\2.PRODUCTION\\Library\\TEMPLATE\\"
dir_res = os.getcwd()

def copytree(src, dst, symlinks=False, ignore=None):
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)

def clicked_Start():  
    cur_dir = os.getcwd()
    search = txtBefore.get()
    new_pattern = txtAfter.get()
    list_dir = ["\\Docs",
                "\\Drawing"]
    for dir_name in list_dir :
        files = find_files(dir_name, search)
        for file_name in files :
            new_file_name = file_name.replace(search, new_pattern)
            os.rename(  cur_dir + dir_name + "\\" + file_name,
                        cur_dir + dir_name + "\\" + new_file_name)    

def clicked_move():  
    copytree(dir_teml,dir_res)




def find_files(dir_name, search):
    files = os.listdir( os.getcwd() +"\\"+ dir_name )
    res_list = []
    for file_name in files :
        if ( file_name.find(search) > -1 ):
            res_list.append(file_name)
    return res_list


    


if __name__ == "__main__":
    def_patern = "TEMPLATE-NAME"


    window = Tk()  
    window.title("File rename")  
    window.geometry('400x250')  

    lblBefore = Label(window, text="До")  
    lblBefore.grid(column=0, row=0)  

    txtBefore = Entry(window,width=50)  
    txtBefore.insert(0,def_patern)
    txtBefore.grid(column=1, row=0)  
    
    lblAfter = Label(window, text="После")  
    lblAfter.grid(column=0, row=2)  

    txtAfter = Entry(window,width=50)  
    txtAfter.grid(column=1, row=2)  
    
    
    btnRename = Button(window, text="Start", command=clicked_Start)  
    btnMoveTempl = Button(window, text="Templ", command=clicked_move)  
    btnRename.grid(column=1, row=3) 
    btnMoveTempl.grid(column=1, row=4) 

    window.mainloop()