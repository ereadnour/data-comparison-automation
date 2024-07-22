import compare_toyota_combined_gui
import compare_gm_combined_gui
import tkinter as tk

def main_menu():
    global root
    root = tk.Tk()
    to_button = tk.Button(root,command =toyota_button_clicked, text = 'Compare Toyota', activebackground = 'SpringGreen4', activeforeground = 'white',font = ('Sans','20','bold'))
    to_button.pack(padx = 40, pady = 30)
    
    gm_button = tk.Button(root,command = gm_button_clicked, text = 'Compare GM', activebackground = 'DodgerBlue3', activeforeground = 'snow', font = ('Sans','20','bold'))
    gm_button.pack(padx = 40, pady = 30)
   
    quit_button = tk.Button(root,command = root.destroy, text = 'QUIT', activebackground = 'red', activeforeground = 'snow', font = ('Sans','20','bold'))
    quit_button.pack(padx = 40, pady = 30)

    root.mainloop()

def toyota_button_clicked():
    root.withdraw()
    compare_toyota_combined_gui.main_to()
    complete = tk.Label(text = "TO Comparison Complete\nCheck folder for Excel Comparison File\n")
    complete.pack()
    root.deiconify()
    
def gm_button_clicked():
    root.withdraw()
    compare_gm_combined_gui.main_gm()
    complete = tk.Label(text = "GM Comparison Complete\nCheck folder for Excel Comparison File\n")
    complete.pack()
    root.deiconify()

if __name__ == "__main__":
    main_menu()
    