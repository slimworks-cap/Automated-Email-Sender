<!-- ABOUT THE PROJECT -->
## About The Project

This project uses email html as a template to send uniform emails to several different contacts. It could be used for email marketing as well as mass email notifications. 

![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/UI.JPG)

This project was originally written to automate a task I performed at work where we would need to inform our business partners and clients of our offices closure during holidays a week in advance.

We would need to send each client/business partner a uniform email notification template manually via outlook. Given it would take hours to finish sending to the entire list of email contacts we had; I wrote this script to automate this process minimal human intervention.

It was originally written to only be run through the command prompt but I rewrote it to be more user friendly for my coworkers who did not know how to program.

Eventually I realized it could be very useful tool for those who work in jobs that regularly send uniform emails to different email contacts and that is why I published this to github w/ instructions on how to use it. 

## Built With
* [Python][Python-url]
* [Pandas][Pandas-url]
* [Pywin32][Pywin32-url]
* [Tkinter][Tkinter-url]

## Initial Setup

In order to use this application you would need to have the below software and modules installed: 
* Python 3
* Any IDE (You may download Atom, Visual Studio Code, or even Notepad!)
* Microsoft Outlook (Ideally, Outlook 2016)
* Microsoft Excel
* the above mentioned python modules in the "Built With" Section

<!-- section of installation of Python and Python modules -->

Once all the software and modules installed you will need to create an excel sheet or use the sample with your email contacts in the below sample format: 

![alt text](https://github.com/slimworks-cap/Automated-Outlook-Email-Sender/blob/main/Images/Sample%20contact%20file.jpg)

(Please note that you should have an EMAIL ADDRESS column and STATUS column for the application to work properly. Please also make sure that these column headers are in UPPER CASE)

After this is done we need to use your installed IDE to edit some lines in the script (Outlook Email Sender.py) I will walk you through how to do this and why this needs to be done. 

1. `Line 19 and Line 41` need to be edited to reflect the file that contains the email contacts that you will be sending the uniform emails to: 

      ```sh
   local_sheet = pd.read_excel(r"C:\Users\loremipsum\Desktop\Sample Contact File.xlsx", sheet_name = 'Sheet1')
      ```

   In both lines please place the absolute path of the excel file in between the double quotes and the sheet where your email contacts are in the excel file. This is so that the application will be able to parse through your email contacts and send them. 


2. Next We need to edit `lines 66-67` and `lines 102-103` to reflect the email contacts you need have copied in your email (as needed) as well as the subject of the uniform email. 
      ```sh
   message.CC = ""                     # this line is for the purposes of adding contacts to be copied in the email 
   message.Subject = ""                # this line is for the purposes of adding a subject to the email
      ```

   If you need to copy multiple email contacts you have to input the email contacts similar to how you would do it in an Outlook email instance. for example:

      ```sh
   first.contact@company.com; second.contact@company.com;`
      ```

   For the subject you may also opt to use `string interpolation` through f-strings like the example below that should make the subject as follows:
      ```sh
   message.Subject = f"Email Marketing // Promos and Services // {row['Company Name']}"
      ```
   You should be able to have an email subject sent to each contact per their company such as "Email Marketing // Promos and Services // Sample Company A"

From here the initial setup is now finished and you may now use the Automated-Outlook-Email-Sender. Please proceed to the Usage section for guidance on how to use to the application

## Usage



<!-- MARKDOWN LINKS & IMAGES -->

[Python-url]: https://www.python.org/
[Pywin32-url]: https://pypi.org/project/pywin32/
[Pandas-url]: https://pandas.pydata.org/
[Tkinter-url]: https://docs.python.org/3/library/tkinter.html
