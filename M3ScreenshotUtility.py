from Tkinter import *
import tkMessageBox
import os
from os.path import expanduser
import time
import datetime
import PIL.ImageGrab
import subprocess, sys

screenshots_path_list = []
comments_window = []

docx_prefix = "Doc_File_"
filename_prefix = "Screenshot_"

home_directory = expanduser("~")
folder_directory = home_directory + "\captor_utility"
screenshot_directory = folder_directory + "/"

if not os.path.exists(folder_directory):
    os.makedirs(folder_directory)

def find(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)

shell_path=find("powershell.exe","C:\\Windows\\System32\\")
shell_path.replace("\\","\\\\")

def capture():
    main_root.iconify()
    time.sleep(1)
    
    global screenshots_path_list
    global screenshot_directory
    
    current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H.%M.%S")
        
    snapshot = PIL.ImageGrab.grab()
    save_path = screenshot_directory + filename_prefix + current_time + '.jpg'
    snapshot.save(save_path)
    screenshots_path_list.append(save_path)

    for comment_window in comments_window:
        comment_window.destroy()
    del comments_window[:]
    
    main_root.deiconify()

def comment():
    
    def show_comment():
        content = entry_box.get()
        global show_comment_root
        show_comment_root = Tk()
        show_comment_root.title("Your Comment")
        show_comment_root.geometry('220x50+810+1')

        comments_window.append(show_comment_root)

        show_comment_frame = Frame(show_comment_root)
        show_comment_frame.pack()

        show_comment_label = Label(show_comment_frame, text = content)   
        show_comment_label.pack(side="top")

        comment_root.destroy()

    global comment_root
    comment_root = Tk()
    comment_root.title("Comment Box")
    # root.geometry('width(px)*height(px) + x axis + y axis')
    comment_root.geometry('250x80+550+1')
    comment_root.resizable(width=False, height=False)
    comment_root.wm_attributes("-toolwindow", "true")
    
    entry_value = StringVar()

    comment_label = Label(comment_root, text="Type your comment here...!!!")
    comment_label.pack()

    comment_root_top_frame = Frame(comment_root)
    comment_root_top_frame.pack()
    comment_root_bottom_frame = Frame(comment_root)
    comment_root_bottom_frame.pack()

    entry_box = Entry(comment_root_top_frame, width=30, bd=5)
    entry_box.pack()

    comment_button1 = Button(comment_root_bottom_frame, text="Comment", fg="white", bg="gray", command = show_comment, width=10)
    comment_button1.pack(side="left")
    comment_button2 = Button(comment_root_bottom_frame, text="Cancel", fg="white", bg="gray", command = comment_root.destroy, width=10)
    comment_button2.pack(side="left")

    comment_root.mainloop()

def pin_window():
    for comment_window in comments_window:
        comment_window.overrideredirect(1)

def unpin_window():
    for comment_window in comments_window:
        comment_window.overrideredirect(0)

def create_file():
    if(len(screenshots_path_list) == 0):
        global status_root
        status_root = Tk()
        status_root.title("Status Message")
        status_root.geometry('280x55+530+310')
        status_root.resizable(width=False, height=False)
        status_root.wm_attributes("-toolwindow", "true")

        label = Label(status_root, text="Please take screenshot first.")
        label.pack(anchor=CENTER)
        
        status_root.mainloop()

    else: 
        def submit():    
            recieved_file_name = file_name_entry_box.get()+ ".docx"
            temp_file=open("C:\\Users\\Public\\time.txt","w")
            temp_file.write(str(recieved_file_name))
            temp_file.close()
            subprocess.call([shell_path, "Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force"])
            subprocess.call([shell_path, "& \"C:\\Program Files\\Screenshot Utility\\powershellscript.ps1\";"])
            file_root.destroy()

        current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H.%M.%S")
        file_name =  docx_prefix + current_time
        
        global file_root
        file_root = Tk()
        file_root.title("File Name Box")
        file_root.geometry('280x55+350+1')

        top_frame = Frame(file_root, bg="gray")
        top_frame.pack(side="top")
        bottom_frame = Frame(file_root, bg="gray")
        bottom_frame.pack(side="bottom")

        label = Label(top_frame, text = "File Name :")
        label.pack(side="left")

        file_name_entry_box = Entry(top_frame, width=30, bd=5)
        file_name_entry_box.insert(0, file_name)
        file_name_entry_box.pack(side="right")    

        submit_button = Button(bottom_frame, text="Submit", command = submit, bg="gray", width=10)
        submit_button.pack(side="left")

        file_root.mainloop()

        
def exit_window():
    value = tkMessageBox.askyesno("Exit Window", "Are you really want to exit?")
    if(value > 0):
        main_root.destroy()

main_root = Tk()
main_root.title("Utility Toolbar")

main_root.geometry('520x26+1+1')
main_root.resizable(width=False, height=False)
#main_root.wm_attributes("-toolwindow", "true")

toolbar = Frame(main_root, bg="gray")
toolbar.pack(side="top", fill=X)

capture_button = Button(toolbar, text="Capture", command = capture, bg="gray", width=10)
capture_button.pack(side="left")
comment_button = Button(toolbar, text="Comment", command = comment, bg="gray", width=10)
comment_button.pack(side="left")
pin_button = Button(toolbar, text="Pin Comment", command = pin_window, bg="gray", width=14)
pin_button.pack(side="left")
unpin_button = Button(toolbar, text="Unpin Comment", command = unpin_window, bg="gray", width=14)
unpin_button.pack(side="left")
create_file_button = Button(toolbar, text="Create File", command = create_file, bg="gray", width=10)
create_file_button.pack(side="left")
exit_button = Button(toolbar, text="Exit", command = exit_window, bg="gray", width=10)
exit_button.pack(side="left")

main_root.mainloop()
