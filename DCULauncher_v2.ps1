# Check Internet connectivity status
Function CheckInternetConnectivityStatus {

    # First we create the request.
    $HTTP_Request = [System.Net.WebRequest]::Create('http://download.windowsupdate.com')

    # We then get a response from the site.
    $HTTP_Response = $HTTP_Request.GetResponse()

    # We then get the HTTP code as an integer.
    $HTTP_Status = [int]$HTTP_Response.StatusCode

    If ($HTTP_Status -eq 200) {
        Return "Connected"
    }
    Else {
        Return "NotConnected"
    }
}


# Check connection location (Remote vs OnSite)
Function CheckConnectionLocation {
    $PingFRMA1149 = Test-Connection -ComputerName FRMA1149 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    $PingFRMA0838 = Test-Connection -ComputerName FRMA0838 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    $InternetStatus = CheckInternetConnectivityStatus

    If ($PingFRMA1149 -or $PingFRMA0838) { 
        $vpnconnection = Get-WmiObject -Class win32_networkadapter -computer localhost | Where-Object {$_.ServiceName -eq "PanGpd"}
        If ($vpnconnection.NetConnectionStatus -eq 2) { 
            Write-Output "Connection status : Remote"
            Return "Remote"
        }
        Else {
            Write-Output "Connection status : OnSite"
            Return "OnSite"
        }
    }
    ElseIf ($InternetStatus -eq "Connected")  { 
        Write-Output "Connection status : Remote"
        Return "Remote"
    }
    Else {
        Exit 0
    }

}


# Variables
$ConnectionStatus = CheckConnectionLocation
$File = "C:\Dell\InstallDay.txt"
[int]$CurrentMonth = Get-Date -Format MM
[int]$CurrentDay = Get-Date -Format dd
[Int]$CurrentDayOfWeek = (Get-Date).DayOfWeek
[int]$InstallDay = Get-Content -Path $File
[Int]$InstallDayOfWeek = (Get-Date -Day $InstallDay -Month $CurrentMonth).DayOfWeek


# Generate day number between 1 and 30
If ( !(Test-Path $File) ) {
    $Day = Get-Random -Minimum 1 -Maximum 30
    New-Item -Path C:\Dell\ -Name file.txt -ItemType File -Value $Day -Force
}


# Change InstallDay for February
If ( $CurrentMonth -eq "02" -and ($InstallDay -eq "29" -or $InstallDay -eq "30") ) {
$InstallDay = 28
[Int]$InstallDayOfWeek = (Get-Date -Day $InstallDay -Month $CurrentMonth).DayOfWeek
}


# Change $InstallDay if it falls during a week end -> Delay the $InstallDay to a random day next week.
If ($InstallDayOfWeek -eq 6) { # Saturday
    [Int]$AddDay = Get-Random -Minimum 2 -Maximum 6
    $DateTemp = (Get-Date -Day $InstallDay).AddDays(+"$AddDay")
    [int]$InstallDay = Get-Date -Date $DateTemp -Format dd
    Set-Content -Path $File -Value $InstallDay
}

If ($InstallDayOfWeek -eq 7) { # Sunday
    [Int]$AddDay = Get-Random -Minimum 1 -Maximum 5
    $DateTemp = (Get-Date -Day $InstallDay).AddDays(+"$AddDay")
    [int]$InstallDay = Get-Date -Date $DateTemp -Format dd
    Set-Content -Path $File -Value $InstallDay
}


# If CurrentDay = LaunchDay
# Launch Dell Command Update only if the device is located Remotly. If the device is OnSite -> delay the $InstallDay to a random day next week.
If ($CurrentDay -eq $InstallDay) {
	If ( $ConnectionStatus -eq "OnSite" ) {
		If ($CurrentDayOfWeek -ne "5") { # Delay +1 day if week day = Monday to Thursday
			$DateTemp = (Get-Date -Day $InstallDay).AddDays(+1)
            [int]$InstallDay = Get-Date -Date $DateTemp -Format dd
            Set-Content -Path $File -Value $InstallDay # Write/replace the new InstallDay into the file "C:\Dell\InstallDay.txt"
		}
        Else { # Delay +3 day if week day = Friday
            $DateTemp = (Get-Date -Day $InstallDay).AddDays(+3)
            [int]$InstallDay = Get-Date -Date $DateTemp -Format dd
            Set-Content -Path $File -Value $InstallDay # Write/replace the new InstallDay into the file "C:\Dell\InstallDay.txt"
        }

	}
	Else {  # Device is remote = Launch DCU 
        Start-Process -FilePath "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe" -ArgumentList "/applyUpdates -reboot=disable -autoSuspendBitLocker=enable -silent -updateSeverity=security,critical -outputLog=""C:\Dell\DellCommandUpdate.log"" -encryptedPassword=""AdwSCABGAkC7xWyh/BQkWpYLnhR3+5lLeuhAr353ILGxoeE2ypKjxXcLBecD/pnBngorgmxt5wCtVVVcZA=="" -encryptionKey=""bioMerieux69""" -Wait -WindowStyle Hidden
        
        # Remove old install flags
        Remove-Item -Path "C:\Dell\UpdateNotNeeded.txt","C:\Dell\UpdateNeeded.txt","C:\Dell\Updating.txt" -ErrorAction SilentlyContinue

        # Create new flag "UpdateNotNeeded"
        New-Item -Path "C:\Dell\UpdateNotNeeded.txt" -ItemType File -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
    }
}

# If CurrentDay <> LaunchDay ==> exit script
Else {
	Return 0
}





# SCCM Deployment : launch every day @ 12pm


