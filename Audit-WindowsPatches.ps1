#Clearing these is not required, but its superstition
Clear-Variable Machine;
Clear-Variable KB;
Clear-Variable Result;
[int]$i = 1;
[array]$Table = @();

Clear-Host
Write-Host "How would you like to import the Host and KB list?`n";
Write-host "1. Import from machinelist.txt and kblist.txt in the same directory the script is in" -ForegroundColor Green;
Write-host "2. Use the values stored in the script" -ForegroundColor Yellow;
Write-Host "3. Default is option 1";
$X = Read-Host -Prompt "`nEnter your choice (1 or 2): ";

switch ($X) {
    1 {
        [array]$MachineList = @(Get-Content -Path .\Machinelist.txt);
        [array]$KBList = @(Get-Content -Path .\KBList.txt);
    }
    2 {
        [array]$MachineList = @();
        [array]$KBList = @();
    }
    Default {
        [array]$MachineList = @(Get-Content -Path .\Machinelist.txt);
        [array]$KBList = @(Get-Content -Path .\KBList.txt);
    }
}


#Clear-Host;
Write-Host "`n*****************************" -ForegroundColor Green;
Write-Host "Ad-Hoc Patch Audit tool.`nAuthor: Joshua Porrata`nVersion: 1.01`nFor instrutions, read the code.`nPress enter to continue" -ForegroundColor white;
Write-Host "*****************************" -ForegroundColor Green;
Read-Host;
Write-Host "First I will Test-Connection with a single attempt.`nIf that Passes then I will use RPC to connect to the remote machine.`nFinally I will invoke 'Get-Hotfix' on the remote machine." -ForegroundColor white;

Write-Host "`nEnsure that you have the CSV at"$env:userprofile\Desktop\KB-AuditResults.csv" closed or else this will all be for nothing`n" -ForegroundColor Yellow

Write-Host "There are "$MachineList.count" hosts and "$KBList.count" Patches to Check" -ForegroundColor White
Write-Host "Get a Coffee and please be patient....`n" -ForegroundColor white;

[datetime]$Start = [System.DateTime]::Now;
$start = $start.ToString("MM-dd-yyyy-HHmm");
#Write-Host "Loop Start Time $start";

foreach ($Machine in $MachineList) {
    Write-Host "Checking host number $i $Machine";
    if (Test-Connection -ComputerName $Machine -Count 1 -ttl 225 -Quiet) {
        foreach ($KB in $KBList) {
            try {
                $result = Get-HotFix -ComputerName $Machine -Id $KB -ErrorAction Stop
                $Row = New-Object PSObject;
                $Row | Add-Member -MemberType NoteProperty -Name "MachineName" -Value $Machine;
                $Row | Add-Member -MemberType NoteProperty -Name "KB" -Value $Result.HotFixID;
                $Row | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $Result.InstalledOn;
                $Table += $Row;
            }
            catch [System.Runtime.InteropServices.COMException] {
                #Write-Host "The Machine named $Machine Could not be connected to" -ForegroundColor Red
                $Row = New-Object PSObject;
                $Row | Add-Member -MemberType NoteProperty -Name "MachineName" -Value $Machine;
                $Row | Add-Member -MemberType NoteProperty -Name "KB" -Value "ConnectionError";
                $Row | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value "ConnectionError";
                $Table += $Row;
            }
            catch [System.Management.Automation.RuntimeException] {
                #Write-Host "The KB named $KB Was not found" -ForegroundColor Red
                $Row = New-Object PSObject;
                $Row | Add-Member -MemberType NoteProperty -Name "MachineName" -Value $Machine;
                $Row | Add-Member -MemberType NoteProperty -Name "KB" -Value $KB;
                $Row | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value "KBNotFound";
                $Table += $Row;
            }            
        }
    }
    else {
        $Row = New-Object PSObject;
        $Row | Add-Member -MemberType NoteProperty -Name "MachineName" -Value $Machine;
        $Row | Add-Member -MemberType NoteProperty -Name "KB" -Value "ConnectionError";
        $Row | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value "ConnectionError";
        $Table += $Row;
    }
    $i++;
}

[datetime]$End = [System.DateTime]::Now;
$End = $End.ToString("MM-dd-yyyy-HHmm");

try {
    $table | Export-Csv -Path "$env:userprofile\Desktop\KB-AuditResults.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop;
    Write-Host "`nAll Done and I have stored the results in $env:userprofile\Desktop\KB-AuditResults.csv`n";
}
catch [System.IO.IOException]{
    Write-Host "`nI was not able to write the results to $env:userprofile\Desktop\KB-AuditResults.csv, There was an IO Exeption." -ForegroundColor Red
    Write-Host "Press enter and I will Dump the results to the console where you can attempt to copy/pate them" -ForegroundColor Red
    Read-Host
    $Table | Format-Table
}

Write-Host "Loop Start Time $start" -ForegroundColor Green;
Write-Host "Loop End Time $end`n" -ForegroundColor Red;