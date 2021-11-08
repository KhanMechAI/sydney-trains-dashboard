from tkinter import Tk, Label, Button, Entry, filedialog
import Dashboard
from Dashboard import MakeDashboard


class MyFirstGUI:
    def __init__(self, master):
        self.master = master
        master.title("Dashboard")

        self.label = Label(master, text="This is our first GUI!")
        self.label.grid(row=0, column=2)

        self.gen_dashboard = Button(master, text="Make Dashboard", command=self.make_dashboard)
        self.gen_dashboard.grid(row=2, column=2)

        self.close_button = Button(master, text="Close", command=master.quit)
        self.close_button.grid(row=3, column=2)

        Label(master, text="Previous Dashboard data_path").grid(row=0)
        Label(master, text="Path to PM Sheets").grid(row=1)
        Label(master, text="Output data_path").grid(row=2)
        Label(master, text="BST Data Path").grid(row=3)

        self.prev_dash_path_e = Entry(master)
        self.pm_sheets_path_e = Entry(master)
        self.out_path_e = Entry(master)
        self.bst_path_e = Entry(master)

        self.prev_dash_path_e.grid(row=0, column=1)
        self.pm_sheets_path_e.grid(row=1, column=1)
        self.out_path_e.grid(row=2, column=1)
        self.bst_path_e.grid(row=3, column=1)

        self.prev_dash_path_browser = filedialog.FileDialog(master)
        # self.prev_dash_path_browser.grid(row=0, column=3)
        

    def make_dashboard(self):
        dash = MakeDashboard(
            prev_dash_path = self.prev_dash_path_e.get(),
            pm_sheets_path = self.pm_sheets_path_e.get(),
            out_path = self.out_path_e.get(),
            bst_path = self.bst_path_e.get(),
        )
        dash.run('Sydney Trains')


root = Tk()
my_gui = MyFirstGUI(root)
root.mainloop()