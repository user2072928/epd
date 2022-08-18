import tkinter as tk
from tkinter.messagebox import showinfo
from MenuPage import MenuPage

class LoginPage(object):
    def __init__(self, master=None):
        self.root = master  
        self.root.geometry('%dx%d' % (300, 180))  #create window
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.createPage()

    def createPage(self):
        self.page = tk.Frame(self.root)  
        self.page.pack()
        tk.Label(self.page).grid(row=0, stick=tk.W)
        tk.Label(self.page, text="Username:").grid(row=1, stick=tk.W, pady=10)
        tk.Entry(self.page, textvariable=self.username).grid(row=1, column=1, stick=tk.E)
        tk.Label(self.page, text="Password:").grid(row=2, stick=tk.W, pady=10)
        tk.Entry(self.page, textvariable=self.password, show='*').grid(row=2, column=1, stick=tk.E)
        tk.Button(self.page, text="Log in", command=self.loginCheck).grid(row=3, stick=tk.W, pady=10)
        tk.Button(self.page, text="Quit", command=self.page.quit).grid(row=3, column=1, stick=tk.E)

    # check username password
    def loginCheck(self):
        name = self.username.get()
        password = self.password.get()
        if name == "egglighting" and password == "123456":
            self.page.destroy()
            MenuPage(self.root)
        else:
            showinfo(title="Warning", message="Incorrect username or password")
