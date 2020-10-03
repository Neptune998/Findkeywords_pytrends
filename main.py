from pytrends.request import TrendReq
import pandas as pd
from tkinter import *
import os

# create class
class MyKeywordApp():

    def __init__(self):

        self.newWindow()

    def newWindow(self):

        # define your window
        root = Tk()
        root.geometry("600x200")
        root.resizable(False, False)
        root.title("App using pytrends, tkinter")

        # logo image for application
        logo = PhotoImage(file='logo_image.png')
        root.iconphoto(False, logo)

        # add labels
        label1 = Label(text='Input a Keyword')
        label1.pack()
        canvas1 = Canvas(root)
        canvas1.pack()
        entry1 = Entry(root)
        canvas1.create_window(200, 20, window=entry1)

        # Function to write in excel file.
        def excelWriter():

            # get the user-input variable
            user_input = entry1.get()
            canvas1.create_window(200, 210)

            # get our Google Trends data
            pytrend = TrendReq()
            suggested_keywords = pytrend.suggestions(keyword=user_input)

            # Dataframe
            df = pd.DataFrame(suggested_keywords)
            df = df.drop(columns='mid')

            # create excel writer object
            writer = pd.ExcelWriter('keywords.xlsx')
            df.to_excel(writer)
            writer.save()

            # open your excel file
            os.system("keywords.xlsx")
            print(df)

        # add button and run loop
        button1 = Button(canvas1, text='Run Report', command=excelWriter)
        canvas1.create_window(200, 50, window=button1)
        root.mainloop()

# call class
MyKeywordApp()

