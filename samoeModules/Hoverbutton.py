from tkinter import Label

class HoverButton(Label) :

    def on_hover(self, event):
        if self.selected == False :
            if self.images :
                if self.images.hovered[0] != None :
                    self.config(image = self.images.hovered)
            if self.colors :
                self.config(bg = self.colors[1])

    def on_unhover(self, event):
        if self.selected == False :
            if self.images :
                self.config(image = self.images.idle)
            if self.colors :            
                self.config(bg = self.colors[0])
            
    def on_clicked(self, event):
        if self.images :
            if self.images.clicked[0] != None :
                self.config(image = self.images.clicked)
        if self.colors :        
            self.config(self, bg = self.colors[2])
        
        self.root.update()

        if self.command != None :
            self.command()

        if self.selectable == False :
            self.selected = False
            try :
                self.root.after(100, self.on_hover(self))
            except :
                print()
        else :
            self.selected = True
            if self.images.selected != None :
                self.config(image = self.images.selected)
            if self.colors :
                self.root.after(50, self.config(bg = self.colors[1]))

    def __init__(self, master, root, text = None, command = None, colors = ['#E6E6E6', '#F2F2F3', '#CCCCCB'], images = None, menuButton = False, font = None, fg = 'black') :
        super(HoverButton, self).__init__(master)
        Label.config(self, text = text, highlightthickness = 0, bd = 0, relief = 'flat', bg = colors[0], font = font, fg = fg)

        self.root = root

        self.colors = colors
        self.images = images

        if self.images :
            self['image'] = self.images.idle

        self.selectable = False
        self.selected = False
        self.menuButton = menuButton

        self.command = command

        self.bind("<Enter>", self.on_hover)
        self.bind("<Leave>", self.on_unhover)
        self.bind("<Button-1>", self.on_clicked)