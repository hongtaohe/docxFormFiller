from docx import Document
from tkinter import *

def replace_text(E1, E2, E3, E4):

    doc = Document(E1.get().replace('\\', '/'))
    oldText = E3.get()
    newText = E4.get()
    for paragraph in doc.paragraphs: #each paragraph is separated
        if oldText in paragraph.text: #find string to be replaced in each paragraph
            inline = paragraph.runs #runs of each paragraph is grouped by the same style
            for i in range(len(inline)): # Loop added to work with runs(strings with same style)
                if oldText in inline[i].text: #if the old String exists in a run
                    text = inline[i].text.replace(oldText, newText) #every occurrence of old text is to be replaced with new text
                    inline[i].text = text #new text is set and replaced old String
            print(paragraph.text)
    doc.save(E2.get().replace('\\', '/'))
    return 1

window = Tk()
window.title("Docx Form Filler")
window.geometry("400x500")
L1 = Label(window, text="Open File")
E1 = Entry(window, width=40)
L2 = Label(window, text="Save File")
E2 = Entry(window, width=40)
L3 = Label(window, text="Old Text")
E3 = Entry(window, width=40)
L4 = Label(window, text="New Text")
E4 = Entry(window, width=40)
B1 = Button(window, text="Enter", command=lambda: replace_text(E1,E2,E3,E4))
B2 = Button(window, text="Quit", command=window.quit)
inst = Text(window,width=40)
inst.insert(END,"Choose a MS Word file to open. \nExample C:\Folder\Folder\Document1.docx (Do not forget the .docx at end)\n")
inst.insert(END,"\nChoose a MS Word file to save. \nExample C:\Folder\Folder\Saved1.docx (Do not forget the .docx at end)\n")
inst.insert(END,"\nType word to replace. Example HOTEL\n")
inst.insert(END,"\nThen type a new word. Example Sandman\n")
inst.insert(END,"\nThen press Enter\n")
inst.insert(END,"\nReuse by changing the Save File. \nExample C:\Folder\Folder\SavedTwo.docx")
L1.grid(column=0,row=0)
E1.grid(column=1,row=0)
L2.grid(column=0,row=1)
E2.grid(column=1,row=1)
L3.grid(column=0,row=2)
E3.grid(column=1,row=2)
L4.grid(column=0,row=3)
E4.grid(column=1,row=3)
B1.grid(column=0,row=4)
B2.grid(column=1,row=4)
inst.grid(row=6,columnspan=2)
window.mainloop()