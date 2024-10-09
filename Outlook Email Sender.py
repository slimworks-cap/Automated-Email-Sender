# Python - create an email sender for Holiday Notifications
# import win32com.client as client

import pandas as pd
import win32com.client as client
import time
from random import randint
import tkinter as tk
from tkinter import messagebox


# ---------------------- Hardcoded stuff --------------------- #

# edit the below line to reflect the location of your file in your computer and sheet name that contains the email contacts 
file_path = r"C:\Users\loremipsum\Desktop\Sample Contact File.xlsx"  

# setting the sheet that we are gonna access in the excel file
file_path_sheet = 'Sheet1'

# start an instance of outlook
outlook = client.Dispatch("Outlook.Application")

# below line reads the file that contains the email contacts we need to send the email to
# 
sheet = pd.read_excel(file_path, sheet_name = file_path_sheet)

# list of values that can be put in the excel sheet status column so the script will exclude them from being sent emails 
nonsend = ['exclude','dead', 'sent']

# set the dataframe 
df = sheet.iloc[0: , 0:4]
# -------------------------- functions ------------------------- #

# I've found that when updating and SAVING the xlsx file (for example: editing a row where the email address was incorrect or deleting a row)
# that script will only take into account the state of the file when it was first read by the bit of code on line 19
# Hence, the script will not take into account updates made on the file until the script was closed 
# to work around this I added a refresh button on the UI of the script to "refresh" the state of the sheet
# PLEASE TAKE NOTE: that you will need to save the xlsx file after making the updates before refreshing for the updates to be taken account by the script 

def refresh():
    # reminder prompt
    tk.messagebox.showinfo("Please Note.",  "Excel File has been Refreshed")
    # To refresh the file being accessed 
    global df
    # local_sheet has to have read_excel method reading the same file as global sheet variable on line 16
    local_sheet = pd.read_excel(file_path, sheet_name = file_path_sheet)

    function_df = local_sheet.iloc[0: , 0:4]
    df = function_df

    print('Refreshed File')

# -------------------------------------------------------------- #

# please take note that the email_draft(): and email_send(): functions are functionally the same. the only differences are that
# email draft will only create drafts and doesnt have the time delay aspect that the email_send(): function has on line 115-116
# and the differences in reminder prompts 

def email_draft():
    # reminder prompt
    tk.messagebox.showinfo("Please Note.",  "Please double check your drafted emails before sending")
    # runtime counter
    start_time = time.time()

    for index, row in df.iterrows():

        if (type(row['STATUS']) == str and row['STATUS'].lower() not in nonsend) or (type(row['STATUS']) != str):

            message = outlook.CreateItem(0)
            message.To = str(row['EMAIL ADDRESS']) 
            message.CC = ""                    # this line is for the purposes of adding contacts to be copied in the email 
            message.Subject = ""               # this line is for the purposes of adding a subject to the email

            # in case you want to include an attachment just uncomment below line and add fullpath of the attachment in below line
            # message.Attachments.Add(r'C:\Users\loremipsum\Desktop\attachment.jpg') 

            message.HTMLBody = txt_edit.get("1.0", tk.END)
            message.Save()
            
            print(f"Row index: {index + 2}")    # prints the current row that is being accessed in the xlsx file 
            print(f"Row data: \n{row}")         # prints the data in each row on a new line on the command prompt

    # end of runtime counter
    end_time = time.time()
    # calculate run time
    run_time = end_time - start_time
    # print runtime
    print(f'script run time: {run_time} seconds to run')

# -------------------------------------------------------------- #

def email_send():
    # Warning prompt (yes/no)
    warning = tk.messagebox.askyesno("Warning", "You are about to send an email to all applicable contacts.\n\nPlease draft and double check email to be sent before proceeding.\n\nDo you wish to proceed?")
    # if statement for the above warning prompt
    if warning == True:
        # runtime counter
        start_time = time.time()

        for index, row in df.iterrows():

            if (type(row['STATUS']) == str and row['STATUS'].lower() not in nonsend) or (type(row['STATUS']) != str):

                message = outlook.CreateItem(0)
                # message.Display()
                message.To = str(row['EMAIL ADDRESS'])
                message.CC = ""                     # this line is for the purposes of adding contacts to be copied in the email 
                message.Subject = ""                # this line is for the purposes of adding a subject to the email
                
                # in case you want to include an attachment just uncomment below line and add fullpath of the attachment in below line
                # message.Attachments.Add(r'C:\Users\loremipsum\Desktop\attachment.jpg') 

                message.HTMLBody = txt_edit.get("1.0", tk.END)
                message.Send()

                print(f"Row index: {index + 2}")    # prints the current row that is being accessed in the xlsx file 
                print(f"Row data: \n{row}")         # prints the data in each row on a new line on the command prompt

                # delay next send by 30 seconds + a random number in between 1 and 30 seconds
                randseconds = randint(1, 30)
                time.sleep(30 + randseconds)

        # end of runtime counter
        end_time = time.time()
        # calculate run time
        run_time = end_time - start_time
        # print runtime
        return print(f'script run time: {run_time} seconds to run')
    elif warning == False:
        pass

# ------------- UI Configuration/Start of Code ----------------- #

window = tk.Tk()
window.title('Outlook Email Sender')
window.rowconfigure(0, minsize = 800, weight = 1)
window.columnconfigure(1, minsize = 800, weight = 1)

txt_edit = tk.Text(window)
frm_buttons = tk.Frame(window, relief = tk.RAISED, bd = 2)
# button to create drafts
btn_draft = tk.Button(frm_buttons, text = 'Draft', command=email_draft)
# button to send emails
btn_send = tk.Button(frm_buttons, text = 'Send', command=email_send)
# button to refresh
btn_refresh = tk.Button(frm_buttons, text = 'Refresh', command=refresh)


btn_draft.grid(row = 0, column = 0, sticky = 'ew', padx = 5, pady = 5)
btn_send.grid(row = 1, column = 0, sticky = 'ew', padx = 5, pady = 5)
btn_refresh.grid(row = 2, column = 0, sticky = 'ew', padx = 5)

frm_buttons.grid(row=0, column=0, sticky="ns")
txt_edit.grid(row=0, column=1, sticky="nsew")

window.mainloop()
