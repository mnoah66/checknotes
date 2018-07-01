from tkinter import *
from tkinter.ttk import *

def isPrime(num):
    return all(num % i for i in range(2, num))

def startSearching():
    primes = []
    for i in range(100000):
        if isPrime(i):
            primes.append(i)
            displayedText.set(len(primes))
            label.update_idletasks()
    displayedText.set('Scan is done.')

root = Tk()

displayedText = StringVar()

label = Label(root, textvariable=displayedText)
label.grid()

root.after(0, startSearching)
root.mainloop()