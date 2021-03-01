# Email-bot
The bot is designed to send customers an e-mail asking them to complete surveys, but it can easily be transformed in order to provide different functionalities.
Each client has an individual and unique link to the survey. Client's sex, name, surname, email adress and the link to the survey should be prepared in an .xlsx file.
The data should have at least the fields shown in the file **mailing_data.xlsx**.
## Configuration
The files required for the program to run are:
- **mailing_bot.py** call file
- message content adapted to the addressee's gender contained in **message_f.txt**, **message_m.txt** files
- signature in **signature.txt**
- a file of any name in the .xlsx format containing the addressees' data and links in the above-mentioned layout
