import pytesseract
import cv2
import tkinter as tk
from tkinter import Label, ttk
from tkinter import filedialog
import pdf2image
from PIL import ImageTk, Image
import os

# Clearer Ui using ctypes
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)

pytesseract.pytesseract.tesseract_cmd = 'D:/Tesseract/tesseract.exe'
PDFTOPPMPATH = r"D:\poppler-0.68.0\bin\pdftoppm.exe"

def openFile():
    global imageOrg
    global label
    label = tk.Label(root,image='')
    path=filedialog.askopenfilename(filetypes=[("Image File",('.jpg',".png",".pdf"))])
    if path.endswith('.pdf'):
        imageOrgs = pdf2image.convert_from_path(path)
        imageOrgs[0].save('temp.jpg',"JPEG")
        imageOrg = cv2.imread('temp.jpg')
        imageOrg = cv2.resize(imageOrg,(0,0),fx=(1080/imageOrg.shape[0]),fy=(1080/imageOrg.shape[0]),interpolation=cv2.INTER_AREA)
        os.remove("temp.jpg")
    else:
        imageOrg = cv2.imread(path)
    
    image = cv2.cvtColor(imageOrg,cv2.COLOR_BGR2RGB)

    photo = ImageTk.PhotoImage(image = Image.fromarray(image))

    label.config(image=photo)
    label.image = photo
    label.grid(row=1,column=0,padx=8,pady=2)
    
def process():
    global imageOrg
    global label
    try:
        img = imageOrg
        img_final = imageOrg
    except:
        tk.messagebox.showerror("Error","Please import a file first")
        return
    label = tk.Label(root,image='')
        # get screen width and height
    ws = root.winfo_screenwidth() # width of the screen
    hs = root.winfo_screenheight() # height of the screen
    w = 240
    h = 100
    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)

    # set the dimensions of the screen 
    # and where it is placed
    prowin = tk.Tk()
    prowin.geometry('%dx%d+%d+%d' % (w, h, x, y))
    prowin.title("")
    prowin.resizable(width = False, height = False)
    prowin.overrideredirect(True)

    prolabel = tk.Label(prowin,text="Please Wait")
    prolabel.pack(padx=8,pady=2)
    probar = ttk.Progressbar(prowin,length=100, orient="horizontal",mode="determinate")
    probar.pack(padx=8,pady=2)

    prolabel.config(text = "Preprocessing Image")
    probar["value"] = 20
    prowin.update()
    root.update()   
    img2gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    ret, mask = cv2.threshold(img2gray, 180, 255, cv2.THRESH_BINARY)
    image_final = cv2.bitwise_and(img2gray, img2gray, mask=mask)
    ret, new_img = cv2.threshold(image_final, 180, 255, cv2.THRESH_BINARY_INV)
    # for black text , cv.THRESH_BINARY_INV # for white text , cv.THRESH_BINARY
    
    kernel = cv2.getStructuringElement(cv2.MORPH_CROSS, (3,3))
    prowin.update()
    dilated = cv2.dilate(new_img, kernel, iterations=5)
    prolabel.config(text = "Finding Text")
    probar["value"] = 40
    prowin.update()
    global Text 
    Text = ""
    contours, hierarchy = cv2.findContours(dilated,cv2.RETR_TREE,cv2.CHAIN_APPROX_NONE)
    for contour in contours:
        prowin.update()
        root.update()
        [x, y, w, h] = cv2.boundingRect(contour)

        if w < 25 and h < 25:
            continue
        
        cropped = img_final[y :y +  h , x : x + w]
        prowin.update()
        text = pytesseract.image_to_string(cropped,lang='eng',config='--psm 1')
        prowin.update()
        check = str(text)
        check = "".join(check.split())
        if len(check) > 1:
            cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 255), 2)
            Text += text
    probar["value"] = 80
    prolabel.config(text = "Reading Text")
    prowin.update()
    img = cv2.cvtColor(img,cv2.COLOR_BGR2RGB)
    photo = ImageTk.PhotoImage(image = Image.fromarray(img))
    label = tk.Label(root,image=photo)
    label.image = photo
    label.grid(row=1,column=0)
    Text = Text.replace('\f','')
    prolabel.config(text = "Finished")
    probar["value"] = 100
    prowin.update()
    prowin.destroy()

def saveText():
    filepath = filedialog.asksaveasfile(initialfile = 'Untitled',defaultextension=".txt")
    if filepath is None:
        return
    tk.messagebox.showinfo(" PDF Text Extractor ","    Saved text file    ")
    filedata = filepath.name
    file = open("".join(str(filedata)),"w")
    file.write(Text)
    file.close()
    
# UI

Text = ""

root = tk.Tk()
root.title("PDF Text Extractor")

ws = root.winfo_screenwidth()
hs = root.winfo_screenheight()
w = 900
h = 960
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)
root.geometry('%dx%d+%d+%d' % (w, h, x, y))
root.resizable(width = True, height = True)
root.grid_columnconfigure(0, weight = 1)

frameLeft = tk.Frame(root)
frameLeft.grid(row=0,column=0,padx=8,pady=2)

FrameButtonsLeft = tk.Frame(frameLeft)
FrameButtonsLeft.grid(row=1,column=0)

buttonOpen = tk.Button(FrameButtonsLeft,text="Open",command = openFile,width=12)
buttonOpen.grid(row=0,column=0,padx=8,pady=2)

buttonProcess = tk.Button(FrameButtonsLeft,text="Detect Text",command = process,width=12)
buttonProcess.grid(row=0,column=1,padx=8,pady=2)

buttonSave = tk.Button(FrameButtonsLeft,text="Save Text File",command = saveText,width=12)
buttonSave.grid(row=0,column=2,padx=8,pady=2)

root.mainloop()

# Ui