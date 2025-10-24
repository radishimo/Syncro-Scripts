# This is an amalgamation of Steven Grabowski's & Sam James' scripts. Thanks Steven & Sam!

# This will deploy Screenconnect to your Syncro asset also enabling syncro integration.
#   It will set the asset friendly name and company name from syncro in your screenconnect portal. (currently, the official syncro integration dosn't do this)
#   It will integrate screenconnect with syncro to allow starting a remote session from the Syncro asset page. (Just like the official integration)

# When creating this in syncro, create two Syncro platform variables...
#   Variable Name - $FriendlyName / Variable Type - platform / Value - asset_custom_field_device_name
#   Variable Name - $CompanyName / Variable Type - platform / Value - customer_business_name_or_customer_full_name


Import-Module $env:SyncroModule

# Convert spaces to %20
$CompanyName = $CompanyName.replace(' ','%20')

$FriendlyName = $FriendlyName.replace(' ','%20')

# URL for ScreenConnect msi download
# Edit this URL to include variables: $FriendlyName & $CompanyName (Variables should go near the end of the URL)(this is to set the company name & asset friendly name in the screenconnect portal)
$url = "https://bullertech.screenconnect.com/Bin/ConnectWiseControl.ClientSetup.msi?h=instance-a3so5x-relay.screenconnect.com&p=443&k=BgIAAACkAABSU0ExAAgAAAEAAQB5QXGO2JuPFsEZNWRKUJ2FspbU%2Fwv4BGoj5TW16R3gLHtFgvR6vvHJLfVyvW6SwEbdbi5pRZQneyYHkd9V%2F71nG1X4YcO94LS5DxgYHtLeJH0OmP0mrXWfoRTPGGJ4VnnkdG4skpQ59RHm8TBxpo5AgScL8PQf3id9AJ2KUwUQFDfNqwIzRvTfwOeX5TxKFL8%2Fo2V1SO%2Flx8rLVYXm37388PSdl0OlIqbSE0pYQmOXpdGo7CsHK8125UMByqiOXV2awA9JvMBn5eu6wHEaRucjS000x9D5Q0xKIvsRiuT3nmwbWn56v3W2jTfFkhF5jLU%2FT1uwSrorqWMKF9Xypq2u&e=Access&y=Guest&t=$FriendlyName&c=$CompanyName&c=&c=&c=&c=&c=&c=&c="

# URL for your screenconnect instance, you can customize the port or leave it off
$scdomain = "bullertech.screenconnect.com"

# put your instance string here find in add/remove programs
$scinstance = "c4d53e2bd6ff64ec"

# put your instance string here find in add/remove programs
$serviceName = "ScreenConnect Client ($scinstance)"

# Your syncro subdomain
$subdomain = "bullertech"

# Your email (for the ticket submission)
$yourEmail = "support@bullertech.com"

If (Get-Service $serviceName -ErrorAction SilentlyContinue) {
   If ((Get-Service $serviceName).Status -eq 'Running') {
       Write-Host "$serviceName is running, skipping service start or install, and performing syncro GUID match."
   } Else {
       #$ticket = Create-Syncro-Ticket -Subdomain $subdomain -Subject "ScreenConnect needs started on $env:computername" -IssueType "Automated" -Status "New"
       Write-Host "$serviceName found, but it is not running for some reason."
       Write-Host "starting $servicename"
       start-service $serviceName
       #$startAt = (Get-Date).AddMinutes(-30).toString("o")
       #Create-Syncro-Ticket-TimerEntry -Subdomain $subdomain -TicketIdOrNumber $ticket.ticket.id -StartTime $startAt -DurationMinutes 5 -Notes "Automated system cleaned up the disk space." -UserIdOrEmail $yourEmail
        
        If ((Get-Service $serviceName).Status -eq 'Running') {
           Write-Host "$serviceName started after manual start commant, now performing syncro GUID match."
        } Else {
            Write-Host "$serviceName not responding to start commant, performing uninstall and fresh install, then performing syncro GUID match."
            cmd.exe /c 'wmic product where name="ScreenConnect Client (c4d53e2bd6ff64ec)" call uninstall /nointeractive'
            
            Write-Host "$serviceName not found - need to install"
            (new-object System.Net.WebClient).DownloadFile($url,'C:\windows\temp\sc.msi')
            msiexec.exe /i c:\windows\temp\sc.msi /quiet
            Write-Host "$serviceName install command sent."
            
        }
   }
} Else {
   #$ticket = Create-Syncro-Ticket -Subdomain $subdomain -Subject "ScreenConnect needs installed on $env:computername" -IssueType "Automated" -Status "New"
       
   Write-Host "$serviceName not found - need to install"
   (new-object System.Net.WebClient).DownloadFile($url,'C:\windows\temp\sc.msi')
   msiexec.exe /i c:\windows\temp\sc.msi /quiet
   #$startAt = (Get-Date).AddMinutes(-30).toString("o")
   #Create-Syncro-Ticket-TimerEntry -Subdomain $subdomain -TicketIdOrNumber $ticket.ticket.id -StartTime $startAt -DurationMinutes 5 -Notes "Automated system cleaned up the disk space." -UserIdOrEmail $yourEmail

}

#Get the Screenconnect URL and write to Syncro asset.
$Keys = Get-ChildItem HKLM:\System\ControlSet001\Services
$Guid = "Null";
$Items = $Keys | Foreach-Object {Get-ItemProperty $_.PsPath }

    ForEach ($Item in $Items)
    {
        if ($item.PSChildName -like "*ScreenConnect Client*")
    {
    $SubKeyName = $Item.PSChildName
    $Guid = (Get-ItemProperty "HKLM:\SYSTEM\ControlSet001\Services\$SubKeyName").ImagePath
    }
}

$GuidParser1 = $Guid -split "&s="
$GuidParser2 = $GuidParser1[1] -split "&k="
$Guid = $GuidParser2[0]
$ScreenConnectUrl = "https://$scdomain/Host#Access/All%20Machines//$Guid/Join"

Write-Host ScreenConnect URL Is: $ScreenConnectUrl

Set-Asset-Field -Subdomain $subdomain -Name "Screenconnect" -Value $ScreenConnectUrl

Start-Sleep -Seconds 10

#Get the Screenconnect GUID and write to Syncro asset.
$val = (Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName").ImagePath
$Regex = [Regex]::new("(?<=s=)(.*?)(?=&)")
$Match = $Regex.Match($val)
if($Match.Success)
{
Set-Asset-Field -Subdomain $subdomain -Name "ScreenConnect GUID" -Value $Match
}