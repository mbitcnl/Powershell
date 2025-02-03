# This script adds or updates DNS records in Active Directory based on data from PHPIPAM.
# Dit script voegt DNS-records toe of werkt deze bij in Active Directory op basis van gegevens uit PHPIPAM.

# Variables
$phpipamBaseUrl = "https://phpipam/api/myapp"
$apiToken = "<api-token>"

# Voorbeeldconfiguratie waarin twee DNS-zones worden gebruikt voor PHPIPAM-subnets
$commonDnsZones = @("example.com", "example.net")

# Define an array of configurations.
# Elke configuratie bevat een PHPIPAM subnet-ID en een DNS-zone
$syncConfigs = @(
    @{
        "subnetId" = "<subnet-id-1>"
        "dnsZones" = $commonDnsZones
    },
    @{
        "subnetId" = "<subnet-id-2>"
        "dnsZones" = $commonDnsZones
    }
)

# Logbestand locatie (pas dit aan indien gewenst)
$logFilePath = "D:\Tools\logs\ipam_updater_subnet.txt"

# Function to log messages / Functie om te loggen
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

# Function to call PHPIPAM API / Functie om API-verzoeken naar PHPIPAM te doen
function Get-PhpipamData {
    param (
        [string]$Endpoint
    )
    $headers = @{
        "token" = $apiToken
    }
    try {
        # Gebruik de base URL-variabele om de volledige URL voor de API-aanroep op te bouwen.
        $url = "$phpipamBaseUrl/$Endpoint"
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
    }
    catch {
        Write-Log "Fout bij het ophalen van data van PHPIPAM endpoint ${Endpoint}: $($_.Exception.Message)" -Type "ERROR"
        return $null
    }
    return $response
}

# Function to ensure a PTR record exists and is correct / Functie om PTR-record te controleren
function Ensure-PtrRecord {
    param (
        [string]$IpAddress,
        [string]$Hostname,
        [string]$ZoneName
    )
    try {
        $ipParts = $IpAddress.Split('.')
        if ($ipParts.Count -ne 4) {
            Write-Log "Ongeldig IP-adres formaat: ${IpAddress}" -Type "ERROR"
            return
        }
        $reverseZone = "$($ipParts[2]).$($ipParts[1]).$($ipParts[0]).in-addr.arpa"
        $ptrRecord = Get-DnsServerResourceRecord -ZoneName $reverseZone -Name $ipParts[3] -ErrorAction SilentlyContinue
        if (-not $ptrRecord) {
            Add-DnsServerResourceRecordPtr -ZoneName $reverseZone -Name $ipParts[3] -PtrDomainName "$Hostname.$ZoneName"
            Write-Log "Nieuw PTR-record aangemaakt voor ${IpAddress} -> ${Hostname}.${ZoneName}"
        }
        else {
            $currentPtrName = $ptrRecord.RecordData.PtrDomainName
            if ($currentPtrName -ne "$Hostname.$ZoneName") {
                Remove-DnsServerResourceRecord -ZoneName $reverseZone -Name $ipParts[3] -RRType Ptr -Force
                Add-DnsServerResourceRecordPtr -ZoneName $reverseZone -Name $ipParts[3] -PtrDomainName "$Hostname.$ZoneName"
                Write-Log "PTR-record bijgewerkt voor ${IpAddress} van ${currentPtrName} naar ${Hostname}.${ZoneName}"
            }
            else {
                Write-Log "PTR-record bestaat al en is correct voor ${IpAddress} -> ${Hostname}.${ZoneName}"
            }
        }
    }
    catch {
        Write-Log "Fout bij het verwerken van PTR-record voor ${IpAddress}: $($_.Exception.Message)" -Type "ERROR"
    }
}

# Main processing loop for adding/updating DNS records
Write-Log "Start verwerking van DNS-records toevoegen/bijwerken..."

foreach ($config in $syncConfigs) {
    $subnetId = $config.subnetId

    # Gebruik alleen de eerste DNS-zone als de standaardzone voor de hostnaam
    $defaultZone = $config.dnsZones[0]
    Write-Log "Ophalen van PHPIPAM data voor Subnet ID ${subnetId} met standaard DNS-zone ${defaultZone}..."
    $subnetData = Get-PhpipamData -Endpoint "subnets/$subnetId/addresses"
    if (-not $subnetData) {
        Write-Log "Geen data gevonden voor Subnet ID ${subnetId}." -Type "ERROR"
        continue
    }

    foreach ($entry in $subnetData.data) {
        Write-Log "Entry details: $(ConvertTo-Json $entry -Depth 3)"
        # Verwerk de standaard hostnaam.
        # Als de hostname een FQDN bevat, wordt alleen het eerste deel gebruikt.
        $hostnameRaw = $entry.hostname
        if ([string]::IsNullOrWhiteSpace($hostnameRaw)) {
            Write-Log "Overslaan van record met lege hostname." -Type "WARNING"
            continue
        }
        if ($hostnameRaw -match "\.") {
            $defaultHostname = $hostnameRaw.Split('.')[0]
        }
        else {
            $defaultHostname = $hostnameRaw
        }

        $ipAddress = $entry.ip
        if ([string]::IsNullOrWhiteSpace($ipAddress)) {
            Write-Log "Overslaan van record met leeg IP-adres." -Type "WARNING"
            continue
        }

        # Verwerk standaard DNS-record enkel in de standaardzone
        $existingRecord = Get-DnsServerResourceRecord -Name $defaultHostname -ZoneName $defaultZone -ErrorAction SilentlyContinue
        if ($existingRecord) {
            try {
                $existingIp = $existingRecord.RecordData.IPv4Address.ToString()
            }
            catch {
                $existingIp = ""
            }
            if ($existingIp -ne $ipAddress) {
                try {
                    Remove-DnsServerResourceRecord -Name $defaultHostname -ZoneName $defaultZone -Force
                    Add-DnsServerResourceRecordA -Name $defaultHostname -ZoneName $defaultZone -IPv4Address $ipAddress
                    Write-Log "DNS-record bijgewerkt: ${defaultHostname} -> ${ipAddress} in zone ${defaultZone} (was ${existingIp})"
                    Ensure-PtrRecord -IpAddress $ipAddress -Hostname $defaultHostname -ZoneName $defaultZone
                }
                catch {
                    Write-Log "Fout bij het bijwerken van DNS-record ${defaultHostname} in zone ${defaultZone}: $($_.Exception.Message)" -Type "ERROR"
                }
            }
            else {
                Write-Log "DNS-record bestaat en is up-to-date: ${defaultHostname} -> ${ipAddress} in zone ${defaultZone}"
                Ensure-PtrRecord -IpAddress $ipAddress -Hostname $defaultHostname -ZoneName $defaultZone
            }
        }
        else {
            try {
                Add-DnsServerResourceRecordA -Name $defaultHostname -ZoneName $defaultZone -IPv4Address $ipAddress
                Write-Log "Nieuw DNS-record toegevoegd: ${defaultHostname} -> ${ipAddress} in zone ${defaultZone}"
                Ensure-PtrRecord -IpAddress $ipAddress -Hostname $defaultHostname -ZoneName $defaultZone
            }
            catch {
                Write-Log "Fout bij het toevoegen van DNS-record ${defaultHostname} in zone ${defaultZone}: $($_.Exception.Message)" -Type "ERROR"
            }
        }

        # Verwerk aanvullende custom DNS-records enkel indien er een waarde is opgegeven in custom_DNS-name, uit de custom_fields property.
        if ($entry.custom_fields -and $entry.custom_fields."custom_DNS-name" -and -not [string]::IsNullOrWhiteSpace($entry.custom_fields."custom_DNS-name")) {
            $customDnsFqdns = $entry.custom_fields."custom_DNS-name" -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
            foreach ($customFqdn in $customDnsFqdns) {
                Write-Log "Bezig met verwerken van custom DNS FQDN: ${customFqdn}"
                $parts = $customFqdn.Split('.')
                if ($parts.Count -ge 2) {
                    $customRecordName = $parts[0]
                    # Haal het zonedeel op (alles na de eerste punt)
                    $customZone = $customFqdn.Substring($customRecordName.Length + 1)
                    Write-Log "Extracted custom record name: ${customRecordName}, custom zone: ${customZone}"
                }
                else {
                    Write-Log "Ongeldig formaat voor custom DNS name: ${customFqdn}" -Type "ERROR"
                    continue
                }

                # Controleer of de custom DNS-zone bestaat voordat je verder gaat
                $customZoneExists = Get-DnsServerZone -Name $customZone -ErrorAction SilentlyContinue
                if (-not $customZoneExists) {
                    Write-Log "DNS-zone bestaat niet: ${customZone}" -Type "ERROR"
                    continue
                }

                # Probeer een bestaand record in de custom zone op te halen.
                $existingCustomRecord = Get-DnsServerResourceRecord -Name $customRecordName -ZoneName $customZone -ErrorAction SilentlyContinue

                if ($existingCustomRecord) {
                    Write-Log "Custom DNS-record bestaat al voor ${customRecordName} in zone ${customZone}."
                    try {
                        $existingCustomIp = $existingCustomRecord.RecordData.IPv4Address.ToString()
                    }
                    catch {
                        $existingCustomIp = ""
                    }
                    if ($existingCustomIp -ne $ipAddress) {
                        try {
                            Write-Log "Bestaand custom record IP (${existingCustomIp}) komt niet overeen met huidig IP (${ipAddress}). Verwijder record..."
                            Remove-DnsServerResourceRecord -Name $customRecordName -ZoneName $customZone -Force
                            Write-Log "Probeer nieuw A-record toe te voegen voor ${customRecordName} in zone ${customZone} met IP ${ipAddress}"
                            Add-DnsServerResourceRecordA -Name $customRecordName -ZoneName $customZone -IPv4Address $ipAddress
                            Write-Log "Custom DNS-record bijgewerkt: ${customRecordName} -> ${ipAddress} in zone ${customZone}"
                            Ensure-PtrRecord -IpAddress $ipAddress -Hostname $customRecordName -ZoneName $customZone
                        }
                        catch {
                            Write-Log "Fout bij het bijwerken van custom DNS-record ${customRecordName} in zone ${customZone}: $($_.Exception.Message)" -Type "ERROR"
                        }
                    }
                    else {
                        Write-Log "Custom DNS-record is al actueel: ${customRecordName} -> ${ipAddress} in zone ${customZone}"
                        Ensure-PtrRecord -IpAddress $ipAddress -Hostname $customRecordName -ZoneName $customZone
                    }
                }
                else {
                    try {
                        Write-Log "Probeer nieuw custom A-record toe te voegen voor ${customRecordName} in zone ${customZone} met IP ${ipAddress}"
                        Add-DnsServerResourceRecordA -Name $customRecordName -ZoneName $customZone -IPv4Address $ipAddress
                        Write-Log "Nieuw custom DNS-record toegevoegd: ${customRecordName} -> ${ipAddress} in zone ${customZone}"
                        Ensure-PtrRecord -IpAddress $ipAddress -Hostname $customRecordName -ZoneName $customZone
                    }
                    catch {
                        Write-Log "Fout bij het toevoegen van custom DNS-record ${customRecordName} in zone ${customZone}: $($_.Exception.Message)" -Type "ERROR"
                    }
                }
            }
        }
    }
}

Write-Log "Verwerking van DNS-records toevoegen/bijwerken voltooid!"
Write-Host "Logbestand opgeslagen op: $logFilePath" 