# Controle op de aanwezigheid van de ActiveDirectory module
if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Write-Host "De module is aanwezig."
	Write-Host "En doorrrrr"
} else {
    Write-Host "De module is niet geïnstalleerd."
	Write-Host "Installeren van module......"
    Install-Module -Name ActiveDirectory -Force -AllowClobber -Scope CurrentUser
}
Import-Module -Name ActiveDirectory
# Defineren en invullen
# Pad waar de output moet komen.
$csvPath = "C:\path\to\output.csv"

# OU's waar gekeken moet worden. 
# Alles scannen = Get-ADComputer -Filter 
# Wil je specifiek iets scannen?
# Haal de # weg voor $ou en pas $computers aan:
# specifieke OU's:  						Get-ADComputer -Filter * -SearchBase $ou
# Specifieke OU's en onderliggende OU's: 	Get-ADComputer -Filter * -SearchBase $ou -SearchScope Subtree
# 
#$ou = "OU=SpecifiekeOU,DC=bieb,DC=eastbridge,DC=eu"
$computers = Get-ADComputer -Filter *

# Importeer bestaande gegevens als het CSV-bestand al bestaat, anders maak een leeg array
if (Test-Path $csvPath) {
    $csvData = Import-Csv $csvPath
} else {
    $csvData = @()
}


# Loopje om alle computer info op te halen.
foreach ($computer in $computers) {
    # Ophalen computernaam
    $computerName = $computer.Name

    # Fabrikant
    $Manufacturer = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $computerName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Manufacturer

    # Serienummer < Kleiner dan 10 wordt opgeslagen.
    $serialNumber = (Get-CimInstance Win32_Bios -ComputerName $computerName -ErrorAction SilentlyContinue).SerialNumber
    if ($serialNumber.Length -lt 10) {
        $serialNumber
    } else {
        $serialNumber = $null
    }

    # Laatste login
    $lastLogonDate = (Get-ADComputer $computerName -Properties LastLogonDate).LastLogonDate
    if ($lastLogonDate -ne $null) {
        $lastLogonDate = $lastLogonDate.ToString('dd/MM/yyyy')
    }

    # Laatste boot-up
    $LastBoot = (Get-CimInstance -ClassName win32_operatingsystem -ComputerName $computerName -ErrorAction SilentlyContinue).lastbootuptime
    if ($LastBoot -ne $null) {
        $LastBoot = $LastBoot.ToString('dd/MM/yyyy')
    }

    # OS naam en versie
    $osInfo = Get-CimInstance Win32_OperatingSystem -ComputerName $computerName -ErrorAction SilentlyContinue
    $osName = $osInfo.Caption
    $osVersion = $osInfo.Version

    # Controle of de computernaam al bestaat in het csv document.
    $csvData = Import-Csv $csvPath -ErrorAction SilentlyContinue
    $existingComputer = $csvData | Where-Object { $_."Name" -eq $computerName }

    # Als de computernaam al bestaat, update dan het serienummer en andere gegevens als ze niet leeg zijn.
    if ($existingComputer) {
        $existingComputer.Name = $computerName
        if (![string]::IsNullOrEmpty($Manufacturer)) {
            $existingComputer.Fabrikant = $Manufacturer
        }
        if (![string]::IsNullOrEmpty($serialNumber)) {
            $existingComputer.SerialNumber = $serialNumber
        }
        if (![string]::IsNullOrEmpty($lastLogonDate)) {
            $existingComputer.LastLogonDate = $lastLogonDate
        }
        if (![string]::IsNullOrEmpty($LastBoot)) {
            $existingComputer.LastBoot = $LastBoot
        }
        if (![string]::IsNullOrEmpty($osName)) {
            $existingComputer.osName = $osName
        }
        if (![string]::IsNullOrEmpty($osVersion)) {
            $existingComputer.osVersion = $osVersion
        }
    }
    # Als de computernaam nog niet bestaat, voeg een nieuwe rij toe aan het csv document.
    else {
        $newRow = New-Object PSObject -Property @{
            Name = $computerName
            Fabrikant = $Manufacturer
            SerialNumber = $serialNumber
            LastLogonDate = $lastLogonDate
            LastBoot = $LastBoot
            osName = $osName
            osVersion = $osVersion
        }
        $csvData += $newRow
    }

    # Exporteer de gegevens naar het CSV-bestand en specificeer de gewenste volgorde van kolommen
    $csvData | Select-Object Name, Fabrikant, SerialNumber, LastLogonDate, LastBoot, osName, osVersion | Export-Csv -Path $csvPath -NoTypeInformation
}