# Automated Outlook Email Sender

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#about-the-project">About The Project</a></li>
    <li><a href="#built-with">Built With</a></li>
    <li><a href="#initial-setup">Intial Setup</a></li>
    <li><a href="#usage">Usage</a></li>
    <li>
      <a href="#customization">Customization</a>
       <ul>
         <li><a href="#i-want-to-include-attachments-in-the-emails-that-i-send-how-do-i-include-this">I have more than just 4 columns in my sheet?</a></li>
         <li><a href="#i-want-to-further-personalize-the-subject-to-have-the-company-i-am-sending-to-be-included-in-the-subject-but-i-dont-want-to-send-each-email-manually-how-do-i-do-that-without-sending-each-manually">I want to further personalize the subject</a></li>
         <li><a href="#i-want-to-add-more-exclusion-words-in-the-excel-sheets-status-column-to-be-more-specific-to-my-task-how-do-i-do-that">I want to add more exclusion words</a></li>
         <li><a href="#i-want-to-include-attachments-in-the-emails-that-i-send-how-do-i-include-this">I want to include attachments in the emails</a></li>
       </ul>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>

<!-- ABOUT THE PROJECT -->
# About The Project

This Python project uses email html as a template to send uniform emails to several different contacts. It could be used for email marketing as well as mass email notifications. 

![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/UI.JPG)

This project was originally written to automate a task I performed at work where we would need to inform our business partners and clients of our offices closure during holidays a week in advance.

We would need to send each client/business partner a uniform email notification template manually via outlook. Given it would take hours to finish sending to the entire list of email contacts we had; I wrote this script to automate this process minimal human intervention.

It was originally written to only be run through the command prompt but I rewrote it to be more user friendly for my coworkers who did not know how to program.

Eventually I realized it could be very useful tool for those who work in jobs that regularly send uniform emails to different email contacts and that is why I published this to github w/ instructions on how to use it. 


# Built With
* [Python][Python-url]
* [Pandas][Pandas-url]
* [Pywin32][Pywin32-url]
* [Tkinter][Tkinter-url]

# Initial Setup

In order to use this application you would need to have the below software and modules installed: 
* Python 3
* Any IDE (You may download Atom, Visual Studio Code, or even Notepad!)
* Microsoft Outlook (Ideally, Outlook 2016)
* Microsoft Excel
* the above mentioned python modules in the "Built With" Section

<!-- section of installation of Python and Python modules -->

Once all the software and modules installed you will need use the sample contact list.xlsx file with your email contacts or create a an excel file with your contacts with a similar format as the sample pic below: 

![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/Sample%20contact%20file.jpg)

(Please note that you should have an EMAIL ADDRESS column and STATUS column for the application to work properly. Please also make sure that these column headers are in UPPER CASE)

After this is done we need to use your installed IDE to edit some lines in the script (Outlook Email Sender.py) I will walk you through how to do this and why this needs to be done. 

1. `Line 19 and Line 41` need to be edited to reflect the file that contains the email contacts that you will be sending the uniform emails to: 

      ```py
   local_sheet = pd.read_excel(r"C:\Users\loremipsum\Desktop\Sample Contact File.xlsx", sheet_name = 'Sheet1')
      ```

   In both lines please place the absolute path of the excel file in between the double quotes and the sheet where your email contacts are in the excel file. This is so that the application will be able to parse through your email contacts and send them. 


2. Next We need to edit `lines 66-67` and `lines 102-103` to reflect the email contacts you need have copied in your email (as needed) as well as the subject of the uniform email. 
      ```py
   message.CC = ""                     # this line is for the purposes of adding contacts to be copied in the email 
   message.Subject = ""                # this line is for the purposes of adding a subject to the email
      ```

   If you need to copy multiple email contacts you have to input the email contacts similar to how you would do it in an Outlook email instance. for example:

      ```py
   first.contact@company.com; second.contact@company.com;`
      ```


3. Once done with the above steps, place the `Outlook Email Sender.py` file and the `excel file` in the same directory/folder. 

From here the initial setup is now finished and you may now use the Automated-Outlook-Email-Sender. Please proceed to the Usage section for guidance on how to use to the application

# Usage

In order to understand how the script works and how to use it you need to have some understanding of how Emails are made. Emails are sent and recieved in two formats: **plain text ** or ** HTML **. Plain text emails have no formatting at all. While emails that have any sort of styling (i.e. bold or italicized letters or even animated GIF's) are formatted with HTML. 

in Outlook you can access the HTML format of an email by by clicking right-clicking an email you have sent or have been sent and clicking "View Source" 

<!-- insert pic of view source -->
![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/View%20Source.png)

This should open an email.txt file with the HTML format of the email. 

<!-- insert pic of html format -->
![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/emailtxt.png)

Hence you will first need to write the body of the uniform email you intend to send and then send this to yourself. The HTML format within the same Email is what the application uses as the basis for the body of the uniform emails you will send across all your email contacts each. 

1. Run the Outlook Email Sender.py file by right clicking on it and Open with > Python. This should open the User Interface (as seen below) of the application and the command prompt

   ![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/UI.JPG)

2. Compose the email you plan to send and then send this to yourself.

<!-- have an example maybe? -->

3. Once you have recieved the Email that you have sent to yourself, right-click on it and click view source and copy paste all of the html in the email.txt file that pops out into the text space on the User Interface of the app.

<!-- picture of the UI with HTML -->
   ![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/UI%20w%20html.png)

4. From here you can either click the `Draft` or `Send` or `Refresh` Button.
   * `Draft`  - clicking the draft button will create a draft email for each of the contacts you are going to send to. You may use this to double check if the emails format is correct or to further personalize the email to each of your contacts so you can send manually.
     
   * `Send`   - clicking the send button will have the application send emails to each of your contacts in an automated fashion. It will send an email to each of your contacts piecemeal starting from the top row of your list of email contacts in the excel sheet with a 30 second to 1 minute time delay in between each email sending. (Note that the Send function will exclude sending all rows with the STATUS column that are marked with the following: `exclude`,`dead`, `sent`).
  
       **Note** _The time delay is a feature not a bug. I have added the time delay in order to work around the restrictions some email providers have on the amount of emails you can send in a given time. i.e.: sending multiple emails in a single minute may have your emails marked as spam by your email provider or mark you as a suspect for phising scams and block your access_
  
   * `Refresh` - clicking the refresh button will refresh the instance of the excel sheet you are using to send to your email contacts. The application will only take into account the first of instance of that excel sheet upon opening the application. it will **NOT** take into consideration any changes made to the sheet after the application is opened unless you save the sheet and then click the refresh button
   
    _i.e. If you had opened the application but make changes your excel sheet in real time to exclude some email contacts from sending click the Send button the application will only take into account the first instance of excel before you made the changes. You will need to save the excel sheet and click refresh on the User Interface to refresh the instance of your excel sheet to reflect the changes to the app script. Otherwise the emails you intended to exclude will still be sent emails since the applications is still parsing through the first instance of the sheet._

You may also opt to monitor the progress of your sending via the command prompt that will show details of each contact you send. For example. below the script had only sent Ted Elow an email because Curt Qutie was marked as "excluded" in the status column. 

<!-- Show picture of command prompt using sample -->
![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/Command%20Prompt.png)

# Customization

Even if you aren't completely familiar with python or programming in general you may still opt to further cuztomize the script of the application to meet your needs as a user. This part of the guide is to help you do so. You would need to open the `Outlook Email Sender.py` script file in an IDE and edit it manually and then save it to make these customizations

I will list some examples of further customization you can do to make the application more attuned to your needs.

## I have more than just 4 columns in my sheet. How do I take into account the other columns in my spread sheet?

  You may opt to edit `line 25` to reflect the amount of columns you want reflected when running the application.
```py
df = sheet.iloc[0: , 0:4] 
```
  Change the number 4 in the above line to reflect the amount of columns you have in your spread sheet.
  for example: if you have 8 columns in your excel sheet you can edit like below.
      
```py
df = sheet.iloc[0: , 0:8] 
```

## I want to further personalize the subject to have the company I am sending to be included in the subject but I don't want to send each email manually. How do I do that without sending each manually? 

   For the subject you may also opt to use `string interpolation` through f-strings in `line 66` and `line 102` like the example below that should make the subject as follows:
     
```py
message.Subject = f"Email Marketing // Promos and Services // {row['Name']}"
```
      
![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/Sample%20contact%20file.jpg)

   Taking for example the above image; this should make it so that the first email sent will have the subject: "Email Marketing // Promos and Services // Curt Qutie" and the next email sent will have "Email Marketing // Promos and Services // Ted Elow" as the subject.

You may further customize lines `line 66` and `line 102` to your needs. 

## I want to add more exclusion words in the excel sheets STATUS column to be more specific to my task. How do I do that.

  You can edit the "nonsend" list in `line 22` to have more exclusion words you can mark in the excel sheet. 

```py
nonsend = ['exclude','dead', 'sent']
```

  For example, you would like to include "bounced" as an exclusion word so that this email contact will not recieve any emails when you send through the application you can edit the script like below.
        
```py
nonsend = ['exclude','dead', 'sent','bounced']
```

## I want to include attachments in the emails that I send. How do I include this? 

You may also add attachments using `line 70` and `line 106`. Simply remove the hashtag `#` from both lines and then include the absolute path of the file you want to attach in to the single quotes of both lines.

```py
# message.Attachments.Add(r'C:\Users\loremipsum\Desktop\attachment.jpg') 
```

For example, you would like to include a file called "pic1.png" in your pictures folder as an attachment in your emails. You may opt to edit  `line 70` and `line 106` like below.

```py
message.Attachments.Add(r'C:\Users\loremipsum\Pictures\pic1.png') 
```
This should include "pic1.png" in all emails that you send to your contacts.

# Acknowledgments

In general, you can edit the script to your liking as much as you want feel free to download the script or if you have any requests or questions feel free to leave a comment!

I would not have been able to complete this project without the help of these wonderful guides: 
* [Python Outlook Automation by Izzy Analytics][Pywin32-guide-url] - Guide for Pywin32
* [Tkinter Guide by Real Python][Tkinter-guide-url] - Guide for Tkinter
* [Best-README-Template][markdown-url] - Guide for writing a README.MD file in Github
  
Please do give them a read as they have helped me and they may help you~

<!-- MARKDOWN LINKS & IMAGES -->

[Python-url]: https://www.python.org/
[Pywin32-url]: https://pypi.org/project/pywin32/
[Pandas-url]: https://pandas.pydata.org/
[Tkinter-url]: https://docs.python.org/3/library/tkinter.html
[Pywin32-guide-url]: https://www.youtube.com/watch?v=J3SiyMingRk&list=PLHnSLOMOPT11njaNmENJN6p2ro9MTc7t_
[Tkinter-guide-url]: https://realpython.com/python-gui-tkinter/
[markdown-url]: https://github.com/othneildrew/Best-README-Template/blob/master/README.md?plain=1
