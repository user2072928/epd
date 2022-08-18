import tkinter as tk
from view import InputFrame, QueryFrame, DeleteFrame, ChangeFrame, AboutFrame, UploadFrame

class MenuPage(object):
    def __init__(self, master = None):
        self.root = master
        self.root.title("Menu Page")
        self.root.geometry('%dx%d' % (1400, 770)) #create window
        self.create_page()
        self.inputPage = InputFrame(self.root)    #call functions in view.py
        self.uploadPage = UploadFrame(self.root)
        self.queryPage = QueryFrame(self.root)
        self.deletePage = DeleteFrame(self.root)
        self.changePage = ChangeFrame(self.root)
        self.aboutPage = AboutFrame(self.root)
        self.inputPage.pack()

    def create_page(self):

        menubar = tk.Menu(self.root)
        # add_command
        menubar.add_command(label="New", command=self.input_data)  # label add command
        menubar.add_command(label="Upload", command=self.upload_data)  
        menubar.add_command(label="Query", command=self.query_data)  
        menubar.add_command(label="Delete", command=self.delete_data) 
        menubar.add_command(label="Change", command=self.change_data) 
        menubar.add_command(label="About", command=self.about_data)  

        self.root.config(menu=menubar)


    def input_data(self):
        self.inputPage.pack()           # open inputpage
        self.uploadPage.pack_forget()   # close other pages
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.changePage.pack_forget()
        self.aboutPage.pack_forget()
        
    def upload_data(self):
        self.inputPage.pack_forget()
        self.uploadPage.pack()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.changePage.pack_forget()
        self.aboutPage.pack_forget()        

    def query_data(self):
        self.inputPage.pack_forget()
        self.uploadPage.pack_forget()
        self.queryPage.pack()
        self.deletePage.pack_forget()
        self.changePage.pack_forget()
        self.aboutPage.pack_forget()

    def delete_data(self):
        self.inputPage.pack_forget()
        self.uploadPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack()
        self.changePage.pack_forget()
        self.aboutPage.pack_forget()

    def change_data(self):
        self.inputPage.pack_forget()
        self.uploadPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.changePage.pack()
        self.aboutPage.pack_forget()

    def about_data(self):
        self.inputPage.pack_forget()
        self.uploadPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.changePage.pack_forget()
        self.aboutPage.pack()

if __name__ == "__main__":
    root = tk.Tk()
    MenuPage(root)
    root.mainloop()