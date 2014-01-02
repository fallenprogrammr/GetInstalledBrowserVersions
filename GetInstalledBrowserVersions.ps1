Try
{
    $excel = New-Object -ComObject "Excel.Application"
}
Catch [System.Exception]
{
    Write-Error "Exception occured while instatiating excel."
    Write-Error $Error[0].Exception.StackTrace
}

if($excel)
{
    $excelFile="crossbrowser.xls"
    if(Test-Path($excelFile))
    {
        Try
        {
            $vmWorkbook = $excel.Workbooks.Open($excelFile)
        }
        Catch [System.Exception]
        {
            Write-Error "Exception occurred while opening the excel sheet."
            Write-Error $Error[0].Exception.StackTrace
        }
        if ($vmWorkbook)
            {
            $vmWorksheet = $vmWorkbook.ActiveSheet
            $Row = 2
            do
            {
                $vmName = $vmWorksheet.Range("A$Row").Text
                if ($vmName) 
                {
                    Write-Output "Pinging $vmName"
                    $canPing = Test-Connection -ComputerName $vmName -Quiet
                    if($canPing) 
                    {
                        $ver=GetIEVersion($vmName)
                        "IE [$env:COMPUTERNAME]: $ver" | Export-Csv -Append list.csv
                        $ver=GetChromeVersion($env:COMPUTERNAME)
                        "Chrome [$env:COMPUTERNAME]: $ver" | Export-Csv -Append list.csv
                        $ver=GetFirefoxVersion($env:COMPUTERNAME)
                        "Firefox [$env:COMPUTERNAME]: $ver" | Export-Csv -Append list.csv
                    }
                }
                $Row++
            } While($vmName)
            $excel.Quit()
        }
    }
    else
    {
        $ver=GetIEVersion($env:COMPUTERNAME)
        Write-Output "IE [$env:COMPUTERNAME]: $ver"
        $ver=GetChromeVersion($env:COMPUTERNAME)
        Write-Output "Chrome [$env:COMPUTERNAME]: $ver"
        $ver=GetFirefoxVersion($env:COMPUTERNAME)
        Write-Output "Firefox [$env:COMPUTERNAME]: $ver"
    }
    
}

Function GetIEVersion([string]$machineName)
{
        $hklm = 2147483650
        $key = "SOFTWARE\Microsoft\Internet Explorer"
        $value = "Version"
        $wmi = [wmiclass]"\\$machineName\root\default:stdRegProv"
        return ($wmi.GetStringValue($hklm,$key,$value)).svalue
}

Function GetChromeVersion([string]$machineName)
{
    $hkcr=2147483648
    $key="Wow6432Node\CLSID\{5C65F4B0-3651-4514-B207-D10CB699B14B}\LocalServer32"
    $value="ServerExecutable"
    $wmi = [wmiclass]"\\$machineName\root\default:stdRegProv"
    [string]$versionString = ($wmi.GetStringValue($hkcr,$key,$value)).svalue
    $versionString=$versionString.Substring(0,$versionString.LastIndexOf('\'))
    [int]$posn=$versionString.LastIndexOf('\')
    $posn=$posn + 1
    return $versionString.Substring($posn)
}

Function GetFirefoxVersion([string]$machineName)
{
    $hklm=2147483650
    $key="SOFTWARE\Wow6432Node\Mozilla\Mozilla Firefox"
    $value="CurrentVersion"
    $wmi = [wmiclass]"\\$machineName\root\default:stdRegProv"
    $versionString = ($wmi.GetStringValue($hklm,$key,$value)).svalue
    $versionString = $versionString.Substring(0,$versionString.IndexOf(' '))
    return $versionString
}