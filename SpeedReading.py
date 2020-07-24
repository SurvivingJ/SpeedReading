#Import necessary libraries
from tkinter import *
import tkinter as tk, tkinter.font as tkfont 
from tkinter import ttk as ttk
from tkinter import filedialog
import time
import docx2txt as d2

def browse_files():
    #Select the necessary file
    global fname
    fname = str(filedialog.askopenfilename(filetypes = (("All files", "*"), ("Template files", "*.type"))))
    print(fname)

class GUI():  
    #Boolean
    boolean = False 

    def stop_function(self):
        #Reset label and change boolean value to break the for loop
        self.passage['text'] = ""
        self.boolean = True

    def __init__(self, parent):
        #Container
        self.myParent = parent

        self.main_container = tk.Frame(background="#6699ff")
        self.main_container.grid(row=0, column=0, sticky="nsew")

        self.myParent.grid_rowconfigure(0, weight=1)
        self.myParent.grid_columnconfigure(0, weight=1)

        #Top & Bottom Frames
        self.top_frame = tk.Frame(self.main_container, background="#6699ff")
        self.top_frame.grid(row=0, column=0, sticky="ew")

        self.bottom_frame = tk.Frame(self.main_container, background="#6699ff")
        self.bottom_frame.grid(row=1, column=0,sticky="nsew")

        self.main_container.grid_rowconfigure(1, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        #Split Bottom Frame in two
        self.bottom_top_frame = tk.Frame(self.bottom_frame, background="#6699ff")
        self.bottom_top_frame.grid(row=0, column=0, sticky="n")

        self.bottom_bottom_frame = tk.Frame(self.bottom_frame, background="#6699ff")
        self.bottom_bottom_frame.grid(row=1, column=0, sticky="s")

        #Split Bottom Top Frame in three
        self.bottom_top_left_frame = tk.Frame(self.bottom_top_frame, background="#6699ff")
        self.bottom_top_left_frame.grid(row=1, column=0, sticky="w")

        self.bottom_top_right_frame = tk.Frame(self.bottom_top_frame, background="#6699ff")
        self.bottom_top_right_frame.grid(row=1, column=2, sticky="e")

        self.bottom_top_middle_frame = tk.Frame(self.bottom_top_frame, background="#6699ff")
        self.bottom_top_middle_frame.grid(row=1, column=1, sticky="s")

        #Buttons
        self.browse_button = tk.Button(self.bottom_top_left_frame, text="Browse", command=browse_files)
        self.browse_button.grid(row=1, column=0, sticky="w")
        self.browse_button.config(font=("Times New Roman", 20))

        self.run_button = tk.Button(self.bottom_top_middle_frame, text="Run", command=self.read_text)
        self.run_button.grid(row=1, column=1, sticky="e")
        self.run_button.config(font=("Times New Roman", 20))

        self.end_button = tk.Button(self.bottom_top_right_frame, text="End", command=self.stop_function)
        self.end_button.grid(row=1, column=2, sticky="s")
        self.end_button.config(font=("Times New Roman", 20))

        #Passage objects
        self.passage = tk.Label(self.top_frame, text="")
        self.passage.grid(row=0, column=1, sticky="e")
        self.passage.config(font=("Times New Roman", 30))

        #Speed set objects
        self.speed_control_label = tk.Label(self.bottom_bottom_frame, text="Time Displayed (decimal)")
        self.speed_control_label.grid(row=1, column=0, sticky="w")
        self.speed_control_label.config(font=("Times New Roman", 20))

        self.speed_control = tk.Entry(self.bottom_bottom_frame)
        self.speed_control.grid(row=1, column=1, sticky="e")
        self.speed_control.config(font=("Times New Roman", 20))

        #Number of words objects
        self.words_displayed_label = tk.Label(self.bottom_bottom_frame, text="No. Words Displayed")
        self.words_displayed_label.grid(row=2, column=0, sticky="w")
        self.words_displayed_label.config(font=("Times New Roman", 20))

        self.words_control = tk.Entry(self.bottom_bottom_frame)
        self.words_control.grid(row=2, column=1, sticky="e")
        self.words_control.config(font=("Times New Roman", 20))

    def read_text(self):
        #Define variables
        short_passage = ""
        i = 0

        #Clear fields
        self.passage['text'] = ""

        #Get values of user input
        time_displayed = float(self.speed_control.get())
        words_displayed = int(self.words_control.get())
        
        #Open the text document and read it
        if fname.endswith(".docx"):
            lines = d2.process(fname)
            word_list = lines.split()

            #For each line, read through each word
            for word in word_list:
                #Get values of user input
                time_displayed = float(self.speed_control.get())
                words_displayed = int(self.words_control.get())
                short_passage = short_passage + " " + word
                i += 1

                #Check to make sure only the requested number of words is displayed at one time
                if i == words_displayed:
                    self.passage['text'] = short_passage
                    root.update()
                    i = 0
                    time.sleep(time_displayed)
                    short_passage = ""
                elif self.boolean == True:
                    #End the loop
                    self.boolean = False
                    break
                else:
                    continue
            else:
                self.passage['text'] = ""
        else:
            text = open(fname, encoding='utf-8')
            lines = text.readlines()

            #For each line, read through each word
            for line in lines:
                word_list = line.split()

                for word in word_list:
                    short_passage = short_passage + " " + word
                    i += 1

                    #Check to make sure only the requested number of words is displayed at one time
                    if i == words_displayed:
                        self.passage['text'] = short_passage
                        root.update()
                        i = 0
                        time.sleep(time_displayed)
                        short_passage = ""
                    elif self.boolean == True:
                        #End the loop
                        self.boolean = False
                        return
                    else:
                        continue

#Initialise the window
def main():
    global root
    root = Tk()
    root.title("SpeedReader")
    w, h = root.winfo_screenwidth(), root.winfo_screenheight()
    varMediaPlayer = GUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()