# EXchangeOnline2CSV
Export mails from Exchange online mailbox folders and subfolders to single CSV file

Powershell program which exports mails from list of folders in mailbox located in Exchange Online to single CSV file.
    Program utilize functionality provided by Microsoft Graph and does not need MS outlook to be installed on a computer where you run the program
# Parameters
- UPN - User Principal Name. Usually in a form of email of user which mailbox we going to process.
- pathToSaveFiles - path to create CSV file, if not specified will take current script execution path.
- maxRecordToProcessForEachFolder - max records to process for each folder. If you want "no limit" put some huge value there.. like 50000000
