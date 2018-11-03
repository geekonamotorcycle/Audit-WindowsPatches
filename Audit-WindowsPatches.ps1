#Clearing these is not required, but its superstition
Clear-Variable Machine;
Clear-Variable KB;
Clear-Variable Result;

#all import files should be in the same path as the script. If you dont want to import from a file, then add to the arrays
# To import hosts from a file, remove the '#' from #[array]$MachineList = @(Get-Content -Path .\Machinelist.txt); and add it to the front of the line below it.
[array]$MachineList = @(Get-Content -Path .\Machinelist.txt);
#[array]$MachineList = @("HOST1","HOST2", "HOST3");
# To import KB numbers from a file, remove the '#' from #[array]$KBList = @(Get-Content -Path .\KBList.txt); add it to the front of the line below it.
[array]$KBList = @(Get-Content -Path .\KBList.txt);
#[array]$KBList = @("KB4100480", "KB4465477");

[int]$i = 1;
[array]$Table = @();

Clear-Host;
Write-Host "*****************************" -ForegroundColor Green;
Write-Host "Ad-Hoc Patch Audit tool.`nAuthor: Joshua Porrata`nVersion: 1.00`nFor instrutions, read the code.`nPress enter to continue" -ForegroundColor white;
Write-Host "*****************************" -ForegroundColor Green;
Read-Host;
Write-Host "First I will Test-Connection with a single attempt.`nIf that Passes then I will use RPC to connect to the remote machine.`nFinally I will invoke 'Get-Hotfix' on the remote machine." -ForegroundColor white;

Write-Host "`nEnsure that you have the CSV at "$env:userprofile\Desktop\KB-AuditResults.csv" closed or else this will all be for nothing`n" -ForegroundColor Yellow

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
$table | Export-Csv -Path "$env:userprofile\Desktop\KB-AuditResults.csv" -NoTypeInformation -Encoding UTF8;

Write-Host "`nAll Done and I have stored the results in $env:userprofile\Desktop\KB-AuditResults.csv`n";
Write-Host "Loop Start Time $start" -ForegroundColor Green;
Write-Host "Loop End Time $end`n" -ForegroundColor Red;