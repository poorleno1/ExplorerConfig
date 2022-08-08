function Modify-Explorer ($param1, $param2) {
    $RegistryKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    $RegistryKey2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState"
    $values = @{
        Hidden                = 1
        HideFileExt           = 0
        HideDrivesWithNoMedia = 0
        ShowSuperHidden       = 0
        #0 = Always combine, hide labels, 1 = Combine when taskbar is full, 2 = Never combine
        TaskbarGlomLevel      = 2
        FolderContentsInfoTip = 1
    }
    
    $values2 = @{
        FullPath = 1
    }

    $restartExplorer = $false
    $values.GetEnumerator()  | ForEach-Object {
        $value = (Get-ItemProperty -Path $RegistryKey -Name $_.Key -ErrorAction SilentlyContinue) | select -ExpandProperty $_.Key
        #Write-Host "Current value for $($_.Key) is $value, should be $($_.Value)"
        if ($value -ne $_.Value) {
            $restartExplorer = $true  
        }

        Set-ItemProperty -Path $RegistryKey -Name $_.Key -Value $_.Value
    }

    $values2.GetEnumerator()  | ForEach-Object {
        $value = (Get-ItemProperty -Path $RegistryKey2 -Name $_.Key -ErrorAction SilentlyContinue) | select -ExpandProperty $_.Key
        #Write-Host "Current value for $($_.Key) is $value, should be $($_.Value)"
        if ($value -ne $_.Value) {
            $restartExplorer = $true  
        }

        Set-ItemProperty -Path $RegistryKey2 -Name $_.Key -Value $_.Value
    }

    if ($restartExplorer) {
        Get-Process explorer | Stop-Process
		Start-Process explorer
    }
}


Function Set-PinTaskbar {
    Param (
        [parameter(Mandatory = $True, HelpMessage = "Target item to pin")]
        [ValidateNotNullOrEmpty()]
        [string] $Target
        ,
        [Parameter(Mandatory = $False, HelpMessage = "Target item to unpin")]
        [switch]$Unpin
    )
    If (!(Test-Path $Target)) {
        Write-Verbose "$Target does not exist"
        Break
    }

    
    $KeyPath1 = "HKLM:\SOFTWARE\Classes"
    $KeyPath2 = "*"
    $KeyPath3 = "shell"
    $KeyPath4 = "{:}"
    $ValueName = "ExplorerCommandHandler"
    $ValueData =
        (Get-ItemProperty `
    ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\" + `
            "CommandStore\shell\Windows.taskbarpin")
        ).ExplorerCommandHandler

    $Key2 = (Get-Item $KeyPath1).OpenSubKey($KeyPath2, $true)
    $Key3 = $Key2.CreateSubKey($KeyPath3, $true)
    $Key4 = $Key3.CreateSubKey($KeyPath4, $true)
    $Key4.SetValue($ValueName, $ValueData)
    #>

    $Shell = New-Object -ComObject "Shell.Application"
    $Folder = $Shell.Namespace((Get-Item $Target).DirectoryName)
    $Item = $Folder.ParseName((Get-Item $Target).Name)

    # Registry key where the pinned items are located
    $RegistryKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    # Binary registry value where the pinned items are located
    $RegistryValue = "FavoritesResolve"
    # Gets the contents into an ASCII format
    $CurrentPinsProperty = ([system.text.encoding]::ASCII.GetString((Get-ItemProperty -Path $RegistryKey -Name $RegistryValue | Select-Object -ExpandProperty $RegistryValue)))
    # Specifies the wildcard of the current executable to be pinned, so that it won't attempt to unpin / repin
    $Executable = "*" + (Split-Path $Target -Leaf) + "*"
    # Filters the results for only the characters that we are looking for, so that the search will function
    [string]$CurrentPinsResults = $CurrentPinsProperty -Replace '[^\x20-\x2f^\x30-\x39\x41-\x5A\x61-\x7F]+', ''

    # Unpin if the application is pinned
    If ($Unpin.IsPresent) {
        If ($CurrentPinsResults -like $Executable) {
            $Item.InvokeVerb("{:}")
        }
    }
    Else {
        # Only pin the application if it hasn't been pinned
        If (!($CurrentPinsResults -like $Executable)) {
            $Item.InvokeVerb("{:}")
        }
    }
    
    $Key3.DeleteSubKey($KeyPath4)
    If ($Key3.SubKeyCount -eq 0 -and $Key3.ValueCount -eq 0) {
        $Key2.DeleteSubKey($KeyPath3)
    }
    #>
}

                


$app = @{pin = @("C:\tools\tcpview64.exe","C:\tools\procexp64.exe","C:\tools\Procmon64.exe","C:\windows\system32\cmd.exe", "C:\windows\system32\services.msc", "C:\windows\system32\eventvwr.msc", "C:\Windows\system32\taskmgr.exe"); unpin = @("Internet Explorer") }

$app.pin | ForEach-Object {
    Set-PinTaskbar $_
}

$app.unpin | ForEach-Object {
    $appName = $_
    ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ? { $_.Name -eq $appName }).Verbs() | ? { $_.Name.replace('&', '') -match 'Unpin from taskbar' } | % { $_.DoIt(); $exec = $true }
}

                
Modify-Explorer
