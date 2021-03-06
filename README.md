# Optimize-MailboxesPerDatabase

# Description

PowerShell script that evenly distributes the number of Exchange mailboxes per database

## Getting Started

Download the Optimize-MailboxesPerDatabase.ps1 file

## Basic Examples
Redistribute mailboxes  
```powershell
  .\Optimize-MailboxesPerDatabase.ps1 -Databases 'USDAG-003', 'USDAG-062', 'USDAG-262'  
```
Redistribute mailboxes, move the smallest mailboxes  
```powershell
  .\Optimize-MailboxesPerDatabase.ps1 -Databases 'USDAG-003', 'USDAG-062', 'USDAG-262' -BySize Smallest  
```
Redistribute mailboxes, move the largest mailboxes
```powershell
  .\Optimize-MailboxesPerDatabase.ps1 -Databases 'USDAG-003', 'USDAG-062', 'USDAG-262' -BySize Largest  
``` 
Redistribute mailboxes for all Databases 
```powershell
$Databases = Get-MailboxDatabase  
.\Optimize-MailboxesPerDatabase.ps1 -Databases $Databases.Name  
```
## Script Processing Time

Environment Tested on: Exchange 2016 CU5    
Total Databases: 265  
Total Mailboxes: 7175  

Creating New Move Requests during actuall run time will vary depending on your environment. Using the WhatIf parameter I measure the command run time with the below results.   

Average RunTime: 60 Seconds  

Filtering for Largest mailboxes: 9 minutes  

Filtering for smallest mailboxes: 4 minutes  

If the script processing time is to long you may want to look into incorporating Invoke-Parallel in to the script.   https://github.com/RamblingCookieMonster/Invoke-Parallel

## Contributing

1. Fork it!
2. Create your feature branch: `git checkout -b my-new-feature`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin my-new-feature`
5. Submit a pull request :D

## License

This project is licensed under the MIT License - see the LICENSE.md file for details


