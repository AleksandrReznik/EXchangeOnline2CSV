![EXO2csv](https://raw.githubusercontent.com/AleksandrReznik/EXchangeOnline2CSV/main/EXO2CSV.jpg)
# EXchangeOnline2CSV
Export mails from Exchange online mailbox folders and subfolders to single CSV file.

Powershell program which exports mails from list of folders in mailbox located in Exchange Online to single CSV file.
    Program utilize functionality provided by Microsoft Graph and does not need MS outlook to be installed on a computer where you run the program.
# Parameters
- **UPN** - User Principal Name. Usually in a form of email of user which mailbox we going to process.
- **pathToSaveFiles** - path to create CSV file, if not specified will take current script execution path.
- **maxRecordToProcessForEachFolder** - max records to process for each folder. If you want "no limit" put some huge value there.. like 50000000.
- **foldersToExport** - second level folders to start export from. Subfolder will also be processed.

# Example usage
Example usage from powershell:  
`.\exchangeOnline2CSV.ps1 -upn "user@domain.com" -maxRecordToProcessForEachFolder 10 -foldersToExport "inbox","sent items"`
