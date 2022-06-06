$date = Get-Date -Format g
$path = "$env:windir\Prefetch"
$outPath = "\\192.168.1.1\Share$\$dateFileName.csv"

#Имя рабочей станции
$compName = gwmi Win32_ComputerSystem| %{$_.DNSHostName + '.' + $_.Domain}

#Последний пользователь системы
$lastLogon = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" -Name LastLoggedOnUser
$lastLogon = $lastLogon.LastLoggedOnUser

#Excel
$latestExcel = Get-ChildItem -Path $path | Where-Object {$_.Name -match "excel"}  | Sort CreationTime -Descending | Select -First 1
$latestExcel = $latestExcel.LastWriteTime.tostring(“dd-MM-yyyy”)

#Word
$latestWord = Get-ChildItem -Path $path | Where-Object {$_.Name -match "winword"}  | Sort CreationTime -Descending | Select -First 1
$latestWord = $latestWord.LastWriteTime.tostring(“dd-MM-yyyy”)

#Outlook
$latestOutlook = Get-ChildItem -Path $path | Where-Object {$_.Name -match "outlook"}  | Sort CreationTime -Descending | Select -First 1
$latestOutlook = $latestOutlook.LastWriteTime.tostring(“dd-MM-yyyy”)

#PowerPoint
$latestPowerpnt = Get-ChildItem -Path $path | Where-Object {$_.Name -match "powerpnt"}  | Sort CreationTime -Descending | Select -First 1
$latestPowerpnt = $latestPowerpnt.LastWriteTime.tostring(“dd-MM-yyyy”)

#Preftech статус
$prefetch = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters" -Name EnablePrefetcher
$prefetch = $prefetch.EnablePrefetcher
if($prefetch = 0)
{
    $prefetchStatus = "Disabled"
}
else
{
    $prefetchStatus = "Enabled"
}

#Экспорт в файл
$isfile = Test-Path $outPath 
if($isfile -eq "True") 
{
    #Проверка наличия записи в отчете
    $outContent = Get-Content $outPath | Select-String -pattern $compName | Select-String -pattern "Enabled"
    if($outContent -match $compName)
    {
        exit
    }
    else
    {
        $arr = "$date;$env:UserName;$lastLogon;$compName;$latestExcel;$latestWord;$latestOutlook;$latestPowerpnt;$prefetchStatus" | Out-File $outPath -Append
    }
}
else 
{
    $arr = "Date;User;LastLogon;Host;Excel;Word;Outlook;PowerPoint;Prefetch Status" | Out-File $outPath
    $arr = "$date;$env:UserName;$lastLogon;$compName;$latestExcel;$latestWord;$latestOutlook;$latestPowerpnt;$prefetchStatus" | Out-File $outPath -Append
}
