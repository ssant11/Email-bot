# Email-bot
The bot is designed to send customers an e-mail asking them to complete surveys, but it can easily be transformed in order to provide different functionalities.
In this version of the code each client has an individual and unique link to the survey. Client's sex, name, surname, email adress and the link to the survey should be prepared in an .xlsx file. The data should have at least the fields shown in the file **mailing_data.xlsx**. Fields must have the same names but they don't need to be in the same order, and there can be some additional fields.
## Configuration
The files required for the program to run are:
- **mailing_bot.py** call file
- gender-specific message content in **message_f.txt**, **message_m.txt** files (content can be changed but the name of the files must remain the same)
- signature in **signature.txt** (content can be changed but the name of the files must remain the same)
- a file of any name in the .xlsx format containing the addressees' data and links in the above-mentioned layout
