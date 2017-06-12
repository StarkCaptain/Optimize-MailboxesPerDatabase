<#
  .SYNOPSIS
  Redistributes mailboxes evenly among databases
  .DESCRIPTION
  Gets a list of mailboxes per database(s), Gets the average of mailboxe count for al databases. Creates move requests to distribute mailboxes evenly among all databases.
  .PARAMETER Databases
  Specify a comma delimited list of databases
  .PARAMETER BySize
  You can use the BySize Parameter to redistribute the largest or smallest mailboxes per database
  .EXAMPLE
  Redistribute mailboxes
  .\Optimize-MailboxesPerDatabase.ps1 -Databases 'USDAG-003', 'USDAG-062', 'USDAG-262'
  .EXAMPLE
  Redistribute mailboxes, move the smallest mailboxes 
  .\Optimize-MailboxesPerDatabase.ps1 -Databases 'USDAG-003', 'USDAG-062', 'USDAG-262' -BySize Smallest
  .EXAMPLE
  Redistribute mailboxes, move the largest mailboxes
  .\Optimize-MailboxesPerDatabase.ps1 -Databases 'USDAG-003', 'USDAG-062', 'USDAG-262' -BySize Largest
  .EXAMPLE
  Redistribute mailboxes for all Databases
  .\Optimize-MailboxesPerDatabase.ps1 -Databases $Databases.Name
#>


[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low',DefaultParametersetName='ParamDefault')]
param(
    [Parameter(Mandatory=$true,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNull()]
        [Alias('Database')]
        [string[]]$Databases,

    [Parameter(Mandatory=$false)]
        [ValidateSet('Largest','Smallest')]
        [string]$BySize
)

begin {
    
    #Add the Exchange Powershell SnapIn
    try{
        .'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1' Connect-ExchangeServer -auto
        ADD-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
    }
    catch{
        Write-Error -Message 'Unable to Load Exchange PowerShell Module'
    }
}

process{

    Write-Verbose -Message 'Getting MailboxCount Per Database'

    $DatabasesInfo = @()
    
    #Get the Count of Mailboxes Per Database
    ForEach ($Database in $Databases){
    
        $MailboxCount = (Get-Mailbox -Database $Database -ResultSize Unlimited).Count
    
        $Object = New-Object PSObject -Property @{
            'Database' = $Database
            'MailboxCount' = $MailboxCount
        }
    
        $DatabasesInfo += $Object
    
    }
    
    Write-Verbose -Message 'Getting the Average and Totals for all Databases'
    #Get the Average and Totals
    $DatabaseMeasures = $DatabasesInfo.MailboxCount | Measure-Object -Sum -Average
    $MailboxTotal = $DatabaseMeasures.Sum
    $MailboxAvgPerDatabase = [math]::Round($DatabaseMeasures.Average)
    
    Write-Verbose -Message "Total Mailboxes: $MailboxTotal"
    Write-Verbose -Message "Average Mailboxes Per Database: $MailboxAvgPerDatabase"

    #Get All Databases Over the Average
    $DatabasesOverAvg = $DatabasesInfo | Where MailboxCount -gt $MailboxAvgPerDatabase
    
    #Get All Databases Under the Average
    $DatabasesUnderAvg = $DatabasesInfo | Where MailboxCount -lt $MailboxAvgPerDatabase
    
    Write-Verbose -Message "Getting the number of mailboxe moves required per database"
    #Calculate the number of mailboxes to move for databases over the average
    ForEach ($DatabaseOverAvg in $DatabasesOverAvg){ 

        $MailboxMoveCount = [math]::Round($DatabaseOverAvg.MailboxCount - $MailboxAvgPerDatabase)
        $DatabaseOverAvg | Add-Member -Name 'MailboxMoveCount' -Value $MailboxMoveCount -MemberType NoteProperty
        Write-Verbose -Message "$DatabaseOverAvg number of moves required $MailboxMoveCount"
    }
    
    #Calculate the number of mailboxes to move for databases over under the average
    ForEach ($DatabaseUnderAvg in $DatabasesUnderAvg){ 
    
        $MailboxMoveCount = [math]::Round($MailboxAvgPerDatabase - $DatabaseUnderAvg.MailboxCount)
        $DatabaseUnderAvg | Add-Member -Name 'MailboxMoveCount' -Value $MailboxMoveCount -MemberType NoteProperty
        Write-Verbose -Message "$DatabaseUnderAvg number of moves required $MailboxMoveCount"
    }

    $MailboxMoves = @()
    
    #If the BySize Parameter is specified Get the Mailbox Statistics, this will add proccessing time
    If ($BySize){
        #Get all Mailboxes to Move for Each Database over the Average
        ForEach ($DatabaseOverAvg in $DatabasesOverAvg){ 
            $Mailboxes = Get-Mailbox -ResultSize Unlimited -Database $DatabaseOverAvg.Database | Get-MailboxStatistics

            switch ($BySize.ToLower()){
                'smallest' {$Mailboxes = $Mailboxes | Sort-Object TotalItemSize | Select -First $DatabaseOverAvg.MailboxMoveCount}
                'largest' {$Mailboxes = $Mailboxes | Sort-Object TotalItemSize -Descending | Select -First $DatabaseOverAvg.MailboxMoveCount}
            }
            $MailboxMoves += $Mailboxes
        }

    }Else{

        #Get all Mailboxes to Move for Each Database over the Average
        ForEach ($DatabaseOverAvg in $DatabasesOverAvg){ 
        
            $Mailboxes = Get-Mailbox -ResultSize Unlimited -Database $DatabaseOverAvg.Database | Select -First $DatabaseOverAvg.MailboxMoveCount    
            $MailboxMoves += $Mailboxes
        }
    }

    $IndexStart = 0
    $Count = 0

    #Create Mailbox Moves for Each Database Under the Average
    ForEach ($DatabaseUnderAvg in $DatabasesUnderAvg){
        
        $IndexEnd = ($IndexStart + $DatabaseUnderAvg.MailboxMoveCount) - 1
      
        ForEach ($MailboxMove in $MailboxMoves[$IndexStart..$IndexEnd]){
            
            Write-Host $Count $MailboxMove.SamAccountName $DatabaseUnderAvg.Database -ForegroundColor Red
            If ($pscmdlet.ShouldProcess("New-MoveRequest on $($MailboxMove.DisplayName) to $($DatabaseUnderAvg.Database)")){
                New-MoveRequest -Identity $MailboxMove.SamAccountName -TargetDatabase $DatabaseUnderAvg.Database -BadItemLimit 500 -AcceptLargeDataLoss
            }

            $MailboxMove | Add-Member -Name 'TargetDatabase' -Value $DatabaseUnderAvg.Database -MemberType NoteProperty -Force
        }
    
        #Increase the Index Start
        $Count += 1
        $IndexStart = $IndexEnd + 1

    }
    
   Write-Host "`n*************************************************************`n" -ForegroundColor Green
   Write-Host "Total Databases: $($Databases.Count)" 
   Write-Host "Total Mailboxes: $MailboxTotal"
   Write-Host "Average Target Mailboxes Per Database: $MailboxAvgPerDatabase"
   Write-Host "`nDatabase Information" -ForegroundColor Green
   $DatabasesInfo | Format-Table Database, MailboxCount, MailboxMoveCount -AutoSize
   Write-Host 'Mailbox Moves Required' -ForegroundColor Green
   $MailboxMoves | Format-Table DisplayName, Database, TargetDatabase, TotalItemSize -AutoSize
   
   Return $MailboxMoves
}
