from tkinter import Tk
from PdfConvert import PdfConvert

def main():
    root = Tk()
    app = PdfConvert(root)  
    root.mainloop()

if __name__ == "__main__":
    main()
