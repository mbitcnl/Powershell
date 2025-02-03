## Deze script is gemaakt voor het bijwerken van DNS records in Windows Server op basis van de gegevens uit PHPIPAM.
## Het script haalt de IP-adressen en hostnames op uit PHPIPAM, voegt of werkt de DNS-records bij en controleert of de PTR-records correct zijn.
## Het script logt alle acties en fouten in een logbestand. 
## Gemaakt door: Marvin Bock


# Variabelen
$phpipamBaseUrl = "https://phpipam/api/myapp"
$apiToken = "<api-token>"
$subnetId = "<subnet-id>"  # ID van het subnet in PHPIPAM
$dnsZone = "<dns-zone>"  # DNS-zone in Windows Server
$logFilePath = Join-Path -Path (Split-Path -Path $MyInvocation.MyCommand.Path) -ChildPath "dns_update_log.txt"  # Standaard logbestand

# Pas logbestand aan als je een andere locatie wilt
# $logFilePath = "C:\path\to\custom_log.txt"

# Functie om te loggen
function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    Add-Content -Path $logFilePath -Value $logEntry
    Write-Host $logEntry
}

# Functie om API-verzoeken naar PHPIPAM te doen
function Get-PhpipamData {
    param (
        [string]$Endpoint
    )
    $headers = @{
        "token" = $apiToken
    }
    $response = Invoke-RestMethod -Uri "$phpipamBaseUrl/$Endpoint" -Headers $headers -Method Get
    return $response
}

# Haal alle IP-adressen en hostnames op uit het subnet
Write-Log "Ophalen van gegevens uit PHPIPAM..."
$subnetData = Get-PhpipamData -Endpoint "subnets/$subnetId/addresses"

# Controleer of data is opgehaald
if (-not $subnetData) {
    Write-Log "Geen gegevens gevonden voor subnet ID $subnetId." -Type "ERROR"
    exit
}

# Functie om PTR record te controleren en aan te maken indien nodig
function Ensure-PtrRecord {
    param (
        [string]$IpAddress,
        [string]$Hostname,
        [string]$ZoneName
    )
    
    try {
        # Haal de reverse lookup zone op basis van IP-adres
        $ipParts = $IpAddress.Split('.')
        $reverseZone = "$($ipParts[2]).$($ipParts[1]).$($ipParts[0]).in-addr.arpa"
        
        # Controleer of PTR record bestaat
        $ptrRecord = Get-DnsServerResourceRecord -ZoneName $reverseZone -Name $ipParts[3] -ErrorAction SilentlyContinue
        
        if (-not $ptrRecord) {
            # Voeg PTR record toe als deze niet bestaat
            Add-DnsServerResourceRecordPtr -ZoneName $reverseZone -Name $ipParts[3] -PtrDomainName "$Hostname.$ZoneName"
            Write-Log "Nieuw PTR-record toegevoegd voor $IpAddress -> $Hostname.$ZoneName"
        } else {
            # Controleer of het bestaande PTR record correct is
            $currentPtrName = $ptrRecord.RecordData.PtrDomainName
            if ($currentPtrName -ne "$Hostname.$ZoneName") {
                # Update PTR record als het niet correct is
                Remove-DnsServerResourceRecord -ZoneName $reverseZone -Name $ipParts[3] -RRType Ptr -Force
                Add-DnsServerResourceRecordPtr -ZoneName $reverseZone -Name $ipParts[3] -PtrDomainName "$Hostname.$ZoneName"
                Write-Log "PTR-record bijgewerkt voor $IpAddress van $currentPtrName naar $Hostname.$ZoneName"
            } else {
                Write-Log "PTR-record bestaat al en is correct voor $IpAddress -> $Hostname.$ZoneName"
            }
        }
    } catch {
        $errorMessage = "Fout bij verwerken PTR-record voor {0}: {1}" -f $IpAddress, $_.Exception.Message
        Write-Log $errorMessage -Type "ERROR"
    }
}

# Loop door elk IP-adres en voeg of werk DNS-records bij
Write-Log "Start verwerking van DNS-records..."
foreach ($entry in $subnetData.data) {
    $hostname = $entry.hostname
    $ipAddress = $entry.ip

    if ([string]::IsNullOrWhiteSpace($hostname) -or [string]::IsNullOrWhiteSpace($ipAddress)) {
        Write-Log "Overslaan van leeg of onvolledig record." -Type "WARNING"
        continue
    }

    # Controleer of het DNS-record al bestaat
    $existingRecord = Get-DnsServerResourceRecord -Name $hostname -ZoneName $dnsZone -ErrorAction SilentlyContinue

    if ($existingRecord) {
        # Controleer of het IP-adres verschilt
        $existingIp = $existingRecord.RecordData.IPv4Address.ToString()
        if ($existingIp -ne $ipAddress) {
            # Update het record als het IP-adres verschilt
            try {
                Remove-DnsServerResourceRecord -Name $hostname -ZoneName $dnsZone -Force
                Add-DnsServerResourceRecordA -Name $hostname -ZoneName $dnsZone -IPv4Address $ipAddress
                Write-Log "Bijgewerkt DNS-record: $hostname -> $ipAddress (was $existingIp)"
                # Controleer en update PTR record
                Ensure-PtrRecord -IpAddress $ipAddress -Hostname $hostname -ZoneName $dnsZone
            } catch {
                $errorMessage = "Fout bij bijwerken van {0} -> {1}: {2}" -f $hostname, $ipAddress, $_.Exception.Message
                Write-Log $errorMessage -Type "ERROR"
            }
        } else {
            Write-Log "DNS-record bestaat al en is up-to-date: $hostname -> $ipAddress"
            # Controleer nog steeds PTR record
            Ensure-PtrRecord -IpAddress $ipAddress -Hostname $hostname -ZoneName $dnsZone
        }
    } else {
        # Voeg nieuw record toe als het niet bestaat
        try {
            Add-DnsServerResourceRecordA -Name $hostname -ZoneName $dnsZone -IPv4Address $ipAddress
            Write-Log "Nieuw DNS-record toegevoegd: $hostname -> $ipAddress"
            # Voeg PTR record toe
            Ensure-PtrRecord -IpAddress $ipAddress -Hostname $hostname -ZoneName $dnsZone
        } catch {
            $errorMessage = "Fout bij toevoegen van {0} -> {1}: {2}" -f $hostname, $ipAddress, $_.Exception.Message
            Write-Log $errorMessage -Type "ERROR"
        }
    }
}

Write-Log "Script voltooid!"
Write-Host "Logbestand opgeslagen in: $logFilePath"