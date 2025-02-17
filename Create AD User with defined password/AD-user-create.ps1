# deze powershell script maakt gebruikers aan met een import csv bestand, conversie voor username is voorletter.achternaam@domain.com
# Benodigde modules importeren
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configuratie variabelen
$script:HomeDirectoryPath = "\\pad\to\home\dir # Pas dit aan naar je server pad
$script:HomeDriveLetter = "H:"

# Hoofdform aanmaken
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD User Creation with CSV"
$form.Size = New-Object System.Drawing.Size(600,400)
$form.StartPosition = "CenterScreen"

# Log TextBox
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10,200)
$txtLog.Size = New-Object System.Drawing.Size(560,150)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

# Logging functie
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Information','Warning','Error')]
        [string]$Level = 'Information'
    )
    
    $LogPath = "AD_User_Creation_Logs"
    $LogFile = Join-Path $LogPath ("log_" + (Get-Date -Format "yyyyMMdd") + ".log")
    
    if (-not (Test-Path $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath | Out-Null
    }
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$TimeStamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    
    if ($txtLog.IsHandleCreated) {
        $txtLog.Invoke([Action]{
            $txtLog.AppendText("$LogMessage`r`n")
            $txtLog.ScrollToCaret()
        })
    }
}

# Password TextBox
$lblPassword = New-Object System.Windows.Forms.Label
$lblPassword.Location = New-Object System.Drawing.Point(10,20)
$lblPassword.Size = New-Object System.Drawing.Size(120,20)
$lblPassword.Text = "Default Password:"
$form.Controls.Add($lblPassword)

$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(140,20)
$txtPassword.Size = New-Object System.Drawing.Size(200,20)
$txtPassword.PasswordChar = '*'
$form.Controls.Add($txtPassword)

# OU Selection
$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Location = New-Object System.Drawing.Point(10,50)
$lblOU.Size = New-Object System.Drawing.Size(120,20)
$lblOU.Text = "Target OU:"
$form.Controls.Add($lblOU)

$txtOU = New-Object System.Windows.Forms.TextBox
$txtOU.Location = New-Object System.Drawing.Point(140,50)
$txtOU.Size = New-Object System.Drawing.Size(300,20)
$txtOU.ReadOnly = $true
$form.Controls.Add($txtOU)

$btnBrowseOU = New-Object System.Windows.Forms.Button
$btnBrowseOU.Location = New-Object System.Drawing.Point(450,49)
$btnBrowseOU.Size = New-Object System.Drawing.Size(100,23)
$btnBrowseOU.Text = "Browse OU"
$btnBrowseOU.Add_Click({
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "Select Organizational Unit"
    $ouForm.Size = New-Object System.Drawing.Size(500,600)
    $ouForm.StartPosition = "CenterScreen"

    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Location = New-Object System.Drawing.Point(10,10)
    $treeView.Size = New-Object System.Drawing.Size(460,500)
    $treeView.PathSeparator = "/"
    $ouForm.Controls.Add($treeView)

    function Add-OUNode {
        param ($ParentNode, $DistinguishedName)
        try {
            $OUs = Get-ADOrganizationalUnit -Filter * -SearchBase $DistinguishedName -SearchScope OneLevel | Sort-Object Name
            foreach ($OU in $OUs) {
                $node = New-Object System.Windows.Forms.TreeNode
                $node.Text = $OU.Name
                $node.Tag = $OU.DistinguishedName
                $node.Nodes.Add("")
                if ($ParentNode) {
                    $ParentNode.Nodes.Add($node)
                } else {
                    $treeView.Nodes.Add($node)
                }
            }
        }
        catch {
            Write-Log "Error loading OUs: $_" -Level Error
        }
    }

    $treeView.Add_BeforeExpand({
        $node = $_.Node
        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq "") {
            $node.Nodes.Clear()
            Add-OUNode -ParentNode $node -DistinguishedName $node.Tag
        }
    })

    $domain = Get-ADDomain
    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Text = $domain.DNSRoot
    $rootNode.Tag = $domain.DistinguishedName
    $treeView.Nodes.Add($rootNode)
    Add-OUNode -ParentNode $rootNode -DistinguishedName $domain.DistinguishedName

    $btnSelect = New-Object System.Windows.Forms.Button
    $btnSelect.Location = New-Object System.Drawing.Point(10,520)
    $btnSelect.Size = New-Object System.Drawing.Size(100,23)
    $btnSelect.Text = "Select"
    $btnSelect.Add_Click({
        if ($treeView.SelectedNode -and $treeView.SelectedNode.Tag) {
            $txtOU.Text = $treeView.SelectedNode.Tag
            $ouForm.Close()
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please select an OU", "Selection Required")
        }
    })
    $ouForm.Controls.Add($btnSelect)

    $ouForm.ShowDialog()
})
$form.Controls.Add($btnBrowseOU)

# Import CSV Button
$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Location = New-Object System.Drawing.Point(10,100)
$btnImport.Size = New-Object System.Drawing.Size(120,25)
$btnImport.Text = "Import CSV"
$btnImport.Add_Click({
    if ([string]::IsNullOrEmpty($txtOU.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select an OU first", "OU Required")
        return
    }
    
    if ([string]::IsNullOrEmpty($txtPassword.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a default password", "Password Required")
        return
    }

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.Title = "Select CSV File for Import"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:users = Import-Csv -Path $openFileDialog.FileName
            $script:csvPath = $openFileDialog.FileName
            Write-Log "CSV file loaded: $($openFileDialog.FileName)" -Level Information
            Write-Log "Number of users in CSV: $($script:users.Count)" -Level Information
            $btnStart.Enabled = $true
            [System.Windows.Forms.MessageBox]::Show("CSV file loaded successfully. Click 'Start Creation' to process users.", "CSV Loaded")
        }
        catch {
            Write-Log "Error loading CSV: $_" -Level Error
            [System.Windows.Forms.MessageBox]::Show("Error loading CSV file: $_", "Error")
        }
    }
})
$form.Controls.Add($btnImport)

# Start Button
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Point(10,150)  # Positioned below Import CSV button
$btnStart.Size = New-Object System.Drawing.Size(120,25)
$btnStart.Text = "Start Creation"
$btnStart.Enabled = $false  # Initially disabled until CSV is imported
$form.Controls.Add($btnStart)

# Add Start button click handler
$btnStart.Add_Click({
    if ($null -eq $script:users) {
        [System.Windows.Forms.MessageBox]::Show("Please import a CSV file first", "CSV Required")
        return
    }

    try {
        $successful = 0
        $failed = 0
        $results = @()

        foreach ($user in $script:users) {
            try {
                # Verwijder spaties en converteer naar lowercase
                $firstName = $user.Voornaam.Trim().ToLower()
                $lastName = $user.Achternaam.Trim().ToLower()
                
                # Haal eerste letter van voornaam
                $initial = $firstName.Substring(0,1)
                
                # Verwerk achternaam (verwijder spaties)
                $processedLastName = $lastName -replace '\s+', ''
                
                # Maak basis gebruikersnaam (voorletter.achternaam)
                $userName = "$initial.$processedLastName"
                
                # Vervang speciale karakters
                $userName = $userName -replace '[éèêë]', 'e'
                $userName = $userName -replace '[àáâä]', 'a'
                $userName = $userName -replace '[ìíîï]', 'i'
                $userName = $userName -replace '[òóôö]', 'o'
                $userName = $userName -replace '[ùúûü]', 'u'
                $userName = $userName -replace '[ý¥ÿ]', 'y'
                $userName = $userName -replace '[ñ]', 'n'
                $userName = $userName -replace '[^a-z0-9\.]', ''
                
                # Check voor dubbele gebruikersnamen
                $counter = 1
                $originalUserName = $userName
                while (Get-ADUser -Filter "SamAccountName -eq '$userName'" -ErrorAction SilentlyContinue) {
                    $userName = "$originalUserName$counter"
                    $counter++
                }

                Write-Log "Generated username: $userName" -Level Information

                # Maak email adres
                $email = "$userName@$((Get-ADDomain).DNSRoot)"
                Write-Log "Generated email: $email" -Level Information

                # Parameters voor nieuwe gebruiker
                $newUserParams = @{
                    Name = "$($user.Voornaam) $($user.Achternaam)".Trim()
                    GivenName = $user.Voornaam.Trim()
                    Surname = $user.Achternaam.Trim()
                    SamAccountName = $userName
                    UserPrincipalName = $email
                    EmailAddress = $email
                    AccountPassword = (ConvertTo-SecureString -String $txtPassword.Text -AsPlainText -Force)
                    Enabled = $true
                    Path = $txtOU.Text
                }

                # Debug output
                Write-Log "Creating user with parameters:" -Level Information
                $newUserParams.GetEnumerator() | ForEach-Object {
                    Write-Log "$($_.Key): [$($_.Value)]" -Level Information
                }

                # Maak de gebruiker aan
                New-ADUser @newUserParams

                # Set Home Directory using configured path
                $homePath = Join-Path $script:HomeDirectoryPath $userName
                Set-ADUser -Identity $userName -HomeDrive $script:HomeDriveLetter -HomeDirectory $homePath

                $successful++
                $results += [PSCustomObject]@{
                    Username = $userName
                    Status = "Success"
                    HomePath = $homePath
                }

                Write-Log "Created user: $userName with home directory: $homePath" -Level Information
            }
            catch {
                $failed++
                $results += [PSCustomObject]@{
                    Username = "$($user.Voornaam) $($user.Achternaam)"
                    Status = "Failed: $_"
                    HomePath = "N/A"
                }
                Write-Log "Failed to create user $($user.Voornaam) $($user.Achternaam): $_" -Level Error
            }
        }

        # Export results
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $resultsPath = Join-Path $env:USERPROFILE "Documents\UserCreation_Results_$timestamp.csv"
        $results | Export-Csv -Path $resultsPath -NoTypeInformation

        $summary = @"
Creation completed:
Successful: $successful
Failed: $failed

Results have been exported to:
$resultsPath
"@
        [System.Windows.Forms.MessageBox]::Show($summary, "Creation Complete")
        
        # Reset for next batch
        $btnStart.Enabled = $false
        $script:users = $null
    }
    catch {
        Write-Log "Error during user creation: $_" -Level Error
        [System.Windows.Forms.MessageBox]::Show("Error during user creation: $_", "Error")
    }
})

# Genereer gebruikersnaam functie aanpassen
function Get-UserPrincipalName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FirstName,
        [Parameter(Mandatory=$true)]
        [string]$LastName
    )
    
    try {
        Write-Log "Generating username for: $FirstName $LastName" -Level Information
        
        # Verwijder spaties en converteer naar lowercase
        $cleanFirstName = $FirstName.Trim().ToLower()
        $cleanLastName = $LastName.Trim().ToLower()
        
        # Haal eerste letter van voornaam
        $initial = $cleanFirstName.Substring(0,1)
        
        # Verwerk achternaam (verwijder spaties)
        $processedLastName = $cleanLastName -replace '\s+', ''
        
        # Maak basis gebruikersnaam (voorletter.achternaam)
        [string]$baseUserName = ($initial + "." + $processedLastName).Trim()
        Write-Log "Base username generated: [$baseUserName]" -Level Information
        
        # Vervang speciale karakters
        [string]$normalizedUserName = $baseUserName
        $normalizedUserName = $normalizedUserName -replace '[éèêë]', 'e'
        $normalizedUserName = $normalizedUserName -replace '[àáâä]', 'a'
        $normalizedUserName = $normalizedUserName -replace '[ìíîï]', 'i'
        $normalizedUserName = $normalizedUserName -replace '[òóôö]', 'o'
        $normalizedUserName = $normalizedUserName -replace '[ùúûü]', 'u'
        $normalizedUserName = $normalizedUserName -replace '[ý¥ÿ]', 'y'
        $normalizedUserName = $normalizedUserName -replace '[ñ]', 'n'
        $normalizedUserName = ($normalizedUserName -replace '[^a-z0-9\.]', '').Trim()
        
        Write-Log "Normalized username: [$normalizedUserName]" -Level Information
        
        # Check voor dubbele gebruikersnamen
        [string]$finalUserName = $normalizedUserName
        $counter = 1
        
        while (Get-ADUser -Filter "SamAccountName -eq '$finalUserName'" -ErrorAction SilentlyContinue) {
            Write-Log "Username $finalUserName already exists, trying alternative" -Level Information
            $finalUserName = "$normalizedUserName$counter"
            $counter++
        }
        
        Write-Log "Final username: [$finalUserName]" -Level Information
        return $finalUserName.Trim()
    }
    catch {
        Write-Log "Error in Get-UserPrincipalName: $_" -Level Error
        throw
    }
}

# Start de applicatie
[System.Windows.Forms.Application]::Run($form) 