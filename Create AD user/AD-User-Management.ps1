## Deze powershell script is gemaakt voor het maken van gebruikers in Active Directory. Gebuikers die worden aangemaakt kunnen een wachtwoord krijgen per email.
## SMTP Server settings kunnen worden opgeslagen in een XML configuratie bestand. Het wachtwoord voor de SMTP server wordt opgeslagen in een gehashed wachtwoord.
## Heeft een bulk import functie voor het importeren van gebruikers uit een CSV bestand.
## Het script is gebouwd op de Windows Forms library van PowerShell.
## Het script is getest op PowerShell 5.1 en 7.4
## Geschreven door: Marvin Bock
## Versie: 1.0


# Benodigde modules importeren
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Definieer het pad voor configuratie
$script:ConfigPath = Join-Path $env:USERPROFILE "ADUserManagement.config.xml"

# Hoofdform aanmaken
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD User Management"
$form.Size = New-Object System.Drawing.Size(600,900)
$form.StartPosition = "CenterScreen"

# Log TextBox eerst aanmaken
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10,690)
$txtLog.Size = New-Object System.Drawing.Size(560,100)
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
    
    $LogPath = "AD_User_Management_Logs"
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

# Functie voor het genereren van een random wachtwoord
function Generate-SecurePassword {
    $length = [int]$numPwdLength.Value
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if ($chkNumbers.Checked) { $chars += "0123456789" }
    if ($chkSpecialChars.Checked) { $chars += "!@#$%^&*()_+-=[]{}|;:,.<>?" }
    
    $bytes = New-Object "System.Byte[]" $length
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $rng.GetBytes($bytes)
    
    $result = ""
    for ($i = 0; $i -lt $length; $i++) {
        $result += $chars[$bytes[$i] % $chars.Length]
    }
    
    return $result
}

# Update de Get-UserPrincipalName functie
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
        
        # Maak basis gebruikersnaam zonder extra spaties
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
        
        # Check of gebruikersnaam al bestaat
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

# Controls voor enkele gebruiker
$lblFirstName = New-Object System.Windows.Forms.Label
$lblFirstName.Location = New-Object System.Drawing.Point(10,20)
$lblFirstName.Size = New-Object System.Drawing.Size(100,20)
$lblFirstName.Text = "Voornaam:"
$form.Controls.Add($lblFirstName)

$txtFirstName = New-Object System.Windows.Forms.TextBox
$txtFirstName.Location = New-Object System.Drawing.Point(120,20)
$txtFirstName.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($txtFirstName)

# Achternaam controls
$lblLastName = New-Object System.Windows.Forms.Label
$lblLastName.Location = New-Object System.Drawing.Point(10,50)
$lblLastName.Size = New-Object System.Drawing.Size(100,20)
$lblLastName.Text = "Achternaam:"
$form.Controls.Add($lblLastName)

$txtLastName = New-Object System.Windows.Forms.TextBox
$txtLastName.Location = New-Object System.Drawing.Point(120,50)
$txtLastName.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($txtLastName)

# Alternative email field
$lblAltEmail = New-Object System.Windows.Forms.Label
$lblAltEmail.Location = New-Object System.Drawing.Point(10,80)
$lblAltEmail.Size = New-Object System.Drawing.Size(100,20)
$lblAltEmail.Text = "Notify Email:"
$form.Controls.Add($lblAltEmail)

$txtAltEmail = New-Object System.Windows.Forms.TextBox
$txtAltEmail.Location = New-Object System.Drawing.Point(120,80)
$txtAltEmail.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($txtAltEmail)

# Email checkbox
$chkSendMail = New-Object System.Windows.Forms.CheckBox
$chkSendMail.Location = New-Object System.Drawing.Point(120,110)
$chkSendMail.Size = New-Object System.Drawing.Size(200,20)
$chkSendMail.Text = "Verstuur wachtwoord per email"
$chkSendMail.Checked = $true
$form.Controls.Add($chkSendMail)

# Password Policy GroupBox
$gbPasswordPolicy = New-Object System.Windows.Forms.GroupBox
$gbPasswordPolicy.Location = New-Object System.Drawing.Point(10,140)
$gbPasswordPolicy.Size = New-Object System.Drawing.Size(560,100)
$gbPasswordPolicy.Text = "Password Policy"
$form.Controls.Add($gbPasswordPolicy)

# Password Length
$lblPwdLength = New-Object System.Windows.Forms.Label
$lblPwdLength.Location = New-Object System.Drawing.Point(10,25)
$lblPwdLength.Size = New-Object System.Drawing.Size(100,20)
$lblPwdLength.Text = "Min Length:"
$gbPasswordPolicy.Controls.Add($lblPwdLength)

$numPwdLength = New-Object System.Windows.Forms.NumericUpDown
$numPwdLength.Location = New-Object System.Drawing.Point(120,25)
$numPwdLength.Size = New-Object System.Drawing.Size(60,20)
$numPwdLength.Minimum = 8
$numPwdLength.Maximum = 64
$numPwdLength.Value = 12
$gbPasswordPolicy.Controls.Add($numPwdLength)

# Special Characters Required
$chkSpecialChars = New-Object System.Windows.Forms.CheckBox
$chkSpecialChars.Location = New-Object System.Drawing.Point(200,25)
$chkSpecialChars.Size = New-Object System.Drawing.Size(150,20)
$chkSpecialChars.Text = "Special Characters"
$chkSpecialChars.Checked = $true
$gbPasswordPolicy.Controls.Add($chkSpecialChars)

# Numbers Required
$chkNumbers = New-Object System.Windows.Forms.CheckBox
$chkNumbers.Location = New-Object System.Drawing.Point(360,25)
$chkNumbers.Size = New-Object System.Drawing.Size(150,20)
$chkNumbers.Text = "Numbers"
$chkNumbers.Checked = $true
$gbPasswordPolicy.Controls.Add($chkNumbers)

# Functie om UPN Suffixes op te halen
function Get-AllUPNSuffixes {
    try {
        $forest = Get-ADForest
        $domains = Get-ADDomain
        
        # Verzamel alle UPN suffixes
        $upnSuffixes = @()
        
        # Voeg de default domain suffix toe
        $upnSuffixes += $domains.DNSRoot
        
        # Voeg alternative UPN suffixes toe van het forest
        $upnSuffixes += $forest.UPNSuffixes
        
        Write-Log "Retrieved UPN Suffixes: $($upnSuffixes -join ', ')" -Level Information
        return $upnSuffixes | Sort-Object -Unique
    }
    catch {
        Write-Log "Error retrieving UPN Suffixes: $($_.Exception.Message)" -Level Error
        throw
    }
}

# Domain Settings GroupBox aanpassen (hernoemd naar UPN Settings)
$gbUPNSettings = New-Object System.Windows.Forms.GroupBox
$gbUPNSettings.Location = New-Object System.Drawing.Point(10,400)
$gbUPNSettings.Size = New-Object System.Drawing.Size(560,80)
$gbUPNSettings.Text = "UPN Settings"
$form.Controls.Add($gbUPNSettings)

# UPN Suffix Label en ComboBox
$lblUPNSuffix = New-Object System.Windows.Forms.Label
$lblUPNSuffix.Location = New-Object System.Drawing.Point(10,25)
$lblUPNSuffix.Size = New-Object System.Drawing.Size(100,20)
$lblUPNSuffix.Text = "UPN Suffix:"
$gbUPNSettings.Controls.Add($lblUPNSuffix)

$cmbUPNSuffix = New-Object System.Windows.Forms.ComboBox
$cmbUPNSuffix.Location = New-Object System.Drawing.Point(120,25)
$cmbUPNSuffix.Size = New-Object System.Drawing.Size(200,20)
$cmbUPNSuffix.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$gbUPNSettings.Controls.Add($cmbUPNSuffix)

# Preview Label
$lblAccountPreview = New-Object System.Windows.Forms.Label
$lblAccountPreview.Location = New-Object System.Drawing.Point(330,25)
$lblAccountPreview.Size = New-Object System.Drawing.Size(220,20)
$lblAccountPreview.Text = "Account Preview: "
$gbUPNSettings.Controls.Add($lblAccountPreview)

# Vul de UPN Suffix ComboBox
try {
    Write-Log "Starting to populate UPN suffix dropdown" -Level Information
    
    # Vul UPN suffix dropdown
    $upnSuffixes = Get-AllUPNSuffixes
    foreach ($suffix in $upnSuffixes) {
        $cmbUPNSuffix.Items.Add($suffix)
    }
    
    if ($cmbUPNSuffix.Items.Count -gt 0) {
        $cmbUPNSuffix.SelectedIndex = 0
        Write-Log "Selected default UPN suffix: $($cmbUPNSuffix.SelectedItem)" -Level Information
    }
}
catch {
    Write-Log "Error populating UPN suffixes: $($_.Exception.Message)" -Level Error
    [System.Windows.Forms.MessageBox]::Show("Error loading UPN suffixes: $($_.Exception.Message)", "Error")
}

# Functie voor het updaten van de preview
function UpdateAccountPreview {
    if ($txtFirstName.Text -and $txtLastName.Text -and $cmbUPNSuffix.SelectedItem) {
        try {
            # Verwijder spaties aan begin en eind
            $firstName = $txtFirstName.Text.Trim()
            $lastName = $txtLastName.Text.Trim()
            
            # Haal eerste letter van voornaam
            $initial = $firstName.Substring(0,1)
            
            # Verwerk achternaam (verwijder spaties en maak lowercase)
            $processedLastName = $lastName.ToLower() -replace '\s+', ''
            
            # Maak preview gebruikersnaam
            $previewName = "$initial.$processedLastName"
            
            # Convert naar lowercase en verwijder speciale karakters
            $previewName = $previewName.ToLower()
            $previewName = $previewName -replace '[éèêë]', 'e'
            $previewName = $previewName -replace '[àáâä]', 'a'
            $previewName = $previewName -replace '[ìíîï]', 'i'
            $previewName = $previewName -replace '[òóôö]', 'o'
            $previewName = $previewName -replace '[ùúûü]', 'u'
            $previewName = $previewName -replace '[ý¥ÿ]', 'y'
            $previewName = $previewName -replace '[ñ]', 'n'
            $previewName = $previewName -replace '[^a-z0-9\.]', ''
            
            # Maak complete preview met email
            $previewEmail = "$previewName@$($cmbUPNSuffix.SelectedItem)"
            $lblAccountPreview.Text = "Preview: $previewEmail"
        }
        catch {
            $lblAccountPreview.Text = "Preview: <invalid input>"
            Write-Log "Error generating preview: $_" -Level Error
        }
    }
    else {
        $lblAccountPreview.Text = "Preview: <incomplete>"
    }
}

# Event handlers voor live preview
$txtFirstName.Add_TextChanged({
    UpdateAccountPreview
})

$txtLastName.Add_TextChanged({
    UpdateAccountPreview
})

$cmbUPNSuffix.Add_SelectedIndexChanged({
    UpdateAccountPreview
})

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,800)
$progressBar.Size = New-Object System.Drawing.Size(560,23)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

# Status Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10,830)
$lblStatus.Size = New-Object System.Drawing.Size(560,20)
$lblStatus.Text = "Gereed"
$form.Controls.Add($lblStatus)

# OU Selection GroupBox
$gbOUSettings = New-Object System.Windows.Forms.GroupBox
$gbOUSettings.Location = New-Object System.Drawing.Point(10,480)
$gbOUSettings.Size = New-Object System.Drawing.Size(560,80)
$gbOUSettings.Text = "Organizational Unit Selection"
$form.Controls.Add($gbOUSettings)

# OU Path TextBox en Browse Button
$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Location = New-Object System.Drawing.Point(10,25)
$lblOU.Size = New-Object System.Drawing.Size(100,20)
$lblOU.Text = "OU Path:"
$gbOUSettings.Controls.Add($lblOU)

$txtOU = New-Object System.Windows.Forms.TextBox
$txtOU.Location = New-Object System.Drawing.Point(120,25)
$txtOU.Size = New-Object System.Drawing.Size(320,20)
$txtOU.ReadOnly = $true
$gbOUSettings.Controls.Add($txtOU)

$btnBrowseOU = New-Object System.Windows.Forms.Button
$btnBrowseOU.Location = New-Object System.Drawing.Point(450,24)
$btnBrowseOU.Size = New-Object System.Drawing.Size(100,23)
$btnBrowseOU.Text = "Browse OU"
$btnBrowseOU.Add_Click({
    # Create OU Browser Form
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "Select Organizational Unit"
    $ouForm.Size = New-Object System.Drawing.Size(500,600)
    $ouForm.StartPosition = "CenterScreen"

    # Create TreeView
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Location = New-Object System.Drawing.Point(10,10)
    $treeView.Size = New-Object System.Drawing.Size(460,500)
    $treeView.PathSeparator = "/"
    $ouForm.Controls.Add($treeView)

    # Function to add nodes to TreeView
    function Add-OUNode {
        param (
            $ParentNode,
            $DistinguishedName
        )
        try {
            $OUs = Get-ADOrganizationalUnit -Filter * -SearchBase $DistinguishedName -SearchScope OneLevel | Sort-Object Name
            foreach ($OU in $OUs) {
                $node = New-Object System.Windows.Forms.TreeNode
                $node.Text = $OU.Name
                $node.Tag = $OU.DistinguishedName
                $node.Nodes.Add("") # Add dummy node for expand functionality
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

    # Handle node expansion
    $treeView.Add_BeforeExpand({
        $node = $_.Node
        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq "") {
            $node.Nodes.Clear()
            Add-OUNode -ParentNode $node -DistinguishedName $node.Tag
        }
    })

    # Get domain info and add root node
    $domain = Get-ADDomain
    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Text = $domain.DNSRoot
    $rootNode.Tag = $domain.DistinguishedName
    $treeView.Nodes.Add($rootNode)
    Add-OUNode -ParentNode $rootNode -DistinguishedName $domain.DistinguishedName

    # Add Select button
    $btnSelect = New-Object System.Windows.Forms.Button
    $btnSelect.Location = New-Object System.Drawing.Point(10,520)
    $btnSelect.Size = New-Object System.Drawing.Size(100,23)
    $btnSelect.Text = "Select"
    $btnSelect.Add_Click({
        if ($treeView.SelectedNode -and $treeView.SelectedNode.Tag) {
            $txtOU.Text = $treeView.SelectedNode.Tag
            Write-Log "Selected OU: $($treeView.SelectedNode.Tag)" -Level Information
            $ouForm.Close()
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please select an OU", "Selection Required")
        }
    })
    $ouForm.Controls.Add($btnSelect)

    # Add Cancel button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(120,520)
    $btnCancel.Size = New-Object System.Drawing.Size(100,23)
    $btnCancel.Text = "Cancel"
    $btnCancel.Add_Click({ $ouForm.Close() })
    $ouForm.Controls.Add($btnCancel)

    # Show the form
    $ouForm.ShowDialog()
})
$gbOUSettings.Controls.Add($btnBrowseOU)

# Functie om email configuratie te laden
function Get-EmailConfig {
    try {
        if (Test-Path $script:ConfigPath) {
            $config = Import-Clixml -Path $script:ConfigPath
            Write-Log "Loaded email configuration from $script:ConfigPath" -Level Information
            Write-Log "SMTP Server: $($config.SmtpServer), Port: $($config.Port), From: $($config.FromAddress)" -Level Information
            return $config
        }
        Write-Log "No email configuration found at $script:ConfigPath" -Level Error
        return $null
    }
    catch {
        Write-Log "Error loading email configuration: $_" -Level Error
        return $null
    }
}

# Update de Send-PasswordEmail functie
function Send-PasswordEmail {
    param (
        [string]$ToAddress,
        [string]$UserName,
        [string]$Password,
        [string]$AlternativeEmail
    )
    
    try {
        # Laad email configuratie
        $emailConfig = Get-EmailConfig
        if ($null -eq $emailConfig) {
            throw "Email configuration not found. Please configure email settings first."
        }

        # Debug logging
        Write-Log "Email Configuration:" -Level Information
        Write-Log "SMTP Server: $($emailConfig.SmtpServer)" -Level Information
        Write-Log "Port: $($emailConfig.Port)" -Level Information
        Write-Log "From Address: $($emailConfig.FromAddress)" -Level Information
        Write-Log "Use SSL: $($emailConfig.UseSSL)" -Level Information
        Write-Log "Username configured: $(-not [string]::IsNullOrEmpty($emailConfig.Username))" -Level Information
        
        # Bepaal het doel e-mailadres
        $targetEmail = if (![string]::IsNullOrEmpty($AlternativeEmail)) { 
            $AlternativeEmail.Trim()
        } else { 
            $ToAddress.Trim()
        }

        Write-Log "Sending email to: $targetEmail" -Level Information

        # Email body
        $body = @"
Beste gebruiker,

Er is een nieuw account voor u aangemaakt met de volgende gegevens:

Gebruikersnaam: $UserName
Wachtwoord: $Password
E-mail adres: $ToAddress

Gelieve bij eerste aanmelding uw wachtwoord te wijzigen.

Met vriendelijke groet,
IT Support
"@

        # Email parameters
        $emailParams = @{
            From = $emailConfig.FromAddress
            To = $targetEmail
            Subject = "Nieuwe account gegevens $UserName"
            Body = $body
            SmtpServer = $emailConfig.SmtpServer
            Port = $emailConfig.Port
        }

        # Voeg SSL toe indien geconfigureerd
        if ($emailConfig.UseSSL) {
            $emailParams.UseSsl = $true
            Write-Log "Using SSL for email" -Level Information
        }

        # Voeg credentials toe indien geconfigureerd
        if (-not [string]::IsNullOrEmpty($emailConfig.Username)) {
            Write-Log "Using SMTP authentication" -Level Information
            $securePass = $emailConfig.Password | ConvertTo-SecureString
            $credential = New-Object System.Management.Automation.PSCredential($emailConfig.Username, $securePass)
            $emailParams.Credential = $credential
        }

        Write-Log "Sending email with parameters:" -Level Information
        Write-Log ($emailParams | ConvertTo-Json) -Level Information

        Send-MailMessage @emailParams
        Write-Log "Email sent successfully to $targetEmail" -Level Information
        
        return $true
    }
    catch {
        Write-Log "Error sending email: $_" -Level Error
        throw
    }
}

# Test functie voor email instellingen
function Test-EmailSettings {
    param (
        [string]$SmtpServer,
        [int]$Port,
        [string]$FromAddress,
        [bool]$UseSSL,
        [string]$Username,
        [System.Security.SecureString]$Password
    )
    
    try {
        $emailParams = @{
            From = $FromAddress
            To = $FromAddress  # Test naar jezelf
            Subject = "Test Email"
            Body = "This is a test email from AD User Management tool."
            SmtpServer = $SmtpServer
            Port = $Port
        }

        if ($UseSSL) {
            $emailParams.UseSsl = $true
        }

        if ($Username -and $Password) {
            $credential = New-Object System.Management.Automation.PSCredential($Username, $Password)
            $emailParams.Credential = $credential
        }

        Send-MailMessage @emailParams
        Write-Log "Test email sent successfully using configured settings" -Level Information
        return $true
    }
    catch {
        Write-Log "Test email failed: $_" -Level Error
        throw
    }
}

# GroupBox voor de Tools (verplaatst naar onder Password Policy)
$gbTools = New-Object System.Windows.Forms.GroupBox
$gbTools.Location = New-Object System.Drawing.Point(10,270)  # Positie direct onder Password Policy
$gbTools.Size = New-Object System.Drawing.Size(560,120)  # Hoogte aangepast voor 3 buttons
$gbTools.Text = "Tools"
$form.Controls.Add($gbTools)

# Email Settings knop
$btnEmailSettings = New-Object System.Windows.Forms.Button
$btnEmailSettings.Location = New-Object System.Drawing.Point(10,20)  # Positie binnen de GroupBox
$btnEmailSettings.Size = New-Object System.Drawing.Size(120,25)
$btnEmailSettings.Text = "Email Settings"
$btnEmailSettings.Add_Click({ Show-EmailSettings })
$gbTools.Controls.Add($btnEmailSettings)

# Bulk Import knop
$btnBulkImport = New-Object System.Windows.Forms.Button
$btnBulkImport.Location = New-Object System.Drawing.Point(10,50)  # 30 pixels onder Email Settings
$btnBulkImport.Size = New-Object System.Drawing.Size(120,25)
$btnBulkImport.Text = "Bulk Import"
$btnBulkImport.Add_Click({
    # Validatie checks
    if ([string]::IsNullOrEmpty($txtOU.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select an OU first", "OU Required")
        return
    }
    
    if ([string]::IsNullOrEmpty($cmbUPNSuffix.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a UPN Suffix first", "UPN Suffix Required")
        return
    }

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.Title = "Select CSV File for Bulk Import"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Progress form setup
            $progressForm = New-Object System.Windows.Forms.Form
            $progressForm.Text = "Bulk Import Progress"
            $progressForm.Size = New-Object System.Drawing.Size(400,200)
            $progressForm.StartPosition = "CenterScreen"
            $progressForm.FormBorderStyle = "FixedDialog"
            $progressForm.ControlBox = $false

            $progressBar = New-Object System.Windows.Forms.ProgressBar
            $progressBar.Location = New-Object System.Drawing.Point(10,40)
            $progressBar.Size = New-Object System.Drawing.Size(360,20)
            $progressForm.Controls.Add($progressBar)

            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Location = New-Object System.Drawing.Point(10,70)
            $statusLabel.Size = New-Object System.Drawing.Size(360,40)
            $statusLabel.Text = "Preparing import..."
            $progressForm.Controls.Add($statusLabel)

            $resultsLabel = New-Object System.Windows.Forms.Label
            $resultsLabel.Location = New-Object System.Drawing.Point(10,120)
            $resultsLabel.Size = New-Object System.Drawing.Size(360,40)
            $resultsLabel.Text = "Successful: 0 | Failed: 0"
            $progressForm.Controls.Add($resultsLabel)

            $progressForm.Show()
            $progressForm.Refresh()

            # Import CSV
            $users = Import-Csv -Path $openFileDialog.FileName
            $totalUsers = $users.Count
            $progressBar.Maximum = $totalUsers
            $successful = 0
            $failed = 0
            $results = @()

            foreach ($user in $users) {
                try {
                    $statusLabel.Text = "Processing: $($user.Voornaam) $($user.Achternaam)"
                    $progressForm.Refresh()

                    # Genereer gebruikersnaam
                    [string]$userName = Get-UserPrincipalName -FirstName $user.Voornaam -LastName $user.Achternaam
                    $cleanUserName = $userName.Trim()
                    
                    # Genereer wachtwoord
                    $password = Generate-SecurePassword
                    
                    # Maak email adres zonder spaties
                    $upnSuffix = $cmbUPNSuffix.Text.Trim()
                    $email = "{0}@{1}" -f $cleanUserName, $upnSuffix
                    Write-Log "Generated email/UPN: $email" -Level Information

                    # Parameters voor nieuwe gebruiker
                    $newUserParams = @{
                        Name = "$($user.Voornaam) $($user.Achternaam)".Trim()
                        GivenName = $user.Voornaam.Trim()
                        Surname = $user.Achternaam.Trim()
                        SamAccountName = $cleanUserName
                        UserPrincipalName = $email.Trim()
                        EmailAddress = $email.Trim()
                        AccountPassword = (ConvertTo-SecureString -String $password -AsPlainText -Force)
                        Enabled = $true
                        Path = $txtOU.Text.Trim()
                    }

                    # Debug output
                    Write-Log "Creating user with parameters:" -Level Information
                    $newUserParams.GetEnumerator() | ForEach-Object {
                        Write-Log "$($_.Key): [$($_.Value)]" -Level Information
                    }

                    # Maak nieuwe gebruiker aan
                    New-ADUser @newUserParams

                    # Verstuur email als er een mailadres is opgegeven
                    $emailStatus = "No email requested"
                    if (-not [string]::IsNullOrEmpty($user.'Verstuur mail')) {
                        try {
                            Send-PasswordEmail -ToAddress $user.'Verstuur mail'.Trim() -UserName $userName -Password $password
                            $emailStatus = "Email sent"
                        }
                        catch {
                            $emailStatus = "Email failed: $_"
                            Write-Log "Failed to send email for $userName to $($user.'Verstuur mail'): $_" -Level Error
                        }
                    }

                    $successful++
                    $results += [PSCustomObject]@{
                        Username = $userName
                        Email = $email
                        Status = "Success"
                        EmailStatus = $emailStatus
                    }
                }
                catch {
                    $failed++
                    $results += [PSCustomObject]@{
                        Username = "$($user.Voornaam) $($user.Achternaam)"
                        Email = "N/A"
                        Status = "Failed: $_"
                        EmailStatus = "N/A"
                    }
                    Write-Log "Failed to create user $($user.Voornaam) $($user.Achternaam): $_" -Level Error
                }

                $progressBar.Value++
                $resultsLabel.Text = "Successful: $successful | Failed: $failed"
                $progressForm.Refresh()
            }

            # Export resultaten
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $resultsPath = Join-Path $env:USERPROFILE "Documents\BulkImport_Results_$timestamp.csv"
            $results | Export-Csv -Path $resultsPath -NoTypeInformation

            $progressForm.Close()

            # Toon samenvatting
            $summary = @"
Bulk import completed:
Successful: $successful
Failed: $failed

Results have been exported to:
$resultsPath
"@
            [System.Windows.Forms.MessageBox]::Show($summary, "Bulk Import Complete")
            
        }
        catch {
            Write-Log "Error during bulk import: $_" -Level Error
            [System.Windows.Forms.MessageBox]::Show("Error during bulk import: $_", "Error")
        }
    }
})
$gbTools.Controls.Add($btnBulkImport)

# Create User knop (verplaatst naar Tools GroupBox)
$btnCreateUser = New-Object System.Windows.Forms.Button
$btnCreateUser.Location = New-Object System.Drawing.Point(10,80)  # 30 pixels onder Bulk Import
$btnCreateUser.Size = New-Object System.Drawing.Size(120,25)
$btnCreateUser.Text = "Create User"
$btnCreateUser.Add_Click({
    Test-FormValues -Location "Start of Create User"
    
    # Validatie checks
    if ([string]::IsNullOrEmpty($txtOU.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select an OU first", "OU Required")
        return
    }
    
    if ([string]::IsNullOrEmpty($cmbUPNSuffix.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a UPN Suffix first", "UPN Suffix Required")
        return
    }

    try {
        # Genereer gebruikersnaam
        [string]$userName = Get-UserPrincipalName -FirstName $txtFirstName.Text -LastName $txtLastName.Text
        Write-Log "Generated username: [$userName]" -Level Information
        
        if ([string]::IsNullOrEmpty($userName)) {
            throw "Failed to generate valid username"
        }
        
        # Genereer wachtwoord
        $password = Generate-SecurePassword
        
        # Maak email adres zonder spaties
        $upnSuffix = $cmbUPNSuffix.Text.Trim()
        $cleanUserName = $userName.Trim()
        $email = "{0}@{1}" -f $cleanUserName, $upnSuffix
        Write-Log "Generated email: [$email]" -Level Information

        # Parameters voor nieuwe gebruiker
        $newUserParams = @{
            Name = "$($txtFirstName.Text) $($txtLastName.Text)".Trim()
            GivenName = $txtFirstName.Text.Trim()
            Surname = $txtLastName.Text.Trim()
            SamAccountName = $cleanUserName
            UserPrincipalName = $email.Trim()
            EmailAddress = $email.Trim()
            AccountPassword = (ConvertTo-SecureString -String $password -AsPlainText -Force)
            Enabled = $true
            Path = $txtOU.Text.Trim()
        }

        # Debug output
        Write-Log "Creating user with parameters:" -Level Information
        $newUserParams.GetEnumerator() | ForEach-Object {
            Write-Log "$($_.Key): [$($_.Value)]" -Level Information
        }

        # Maak de gebruiker aan
        New-ADUser @newUserParams
        
        Write-Log "User created successfully" -Level Information

        # Email verzending (als aangevinkt)
        if ($chkSendMail.Checked) {
            try {
                Send-PasswordEmail -ToAddress $email -UserName $userName -Password $password -AlternativeEmail $txtAltEmail.Text
            }
            catch {
                $emailError = $_.Exception.Message
                Write-Log "Failed to send email: $emailError" -Level Error
                [System.Windows.Forms.MessageBox]::Show(
                    "User created successfully but failed to send email: $emailError",
                    "Partial Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        }

        # Toon samenvatting
        $summary = @"
Gebruiker succesvol aangemaakt:
Naam: $($txtFirstName.Text) $($txtLastName.Text)
Gebruikersnaam: $userName
UPN/Email: $email
"@
        [System.Windows.Forms.MessageBox]::Show($summary, "Success")
        
        # Clear form fields
        $txtFirstName.Clear()
        $txtLastName.Clear()
        $txtAltEmail.Clear()
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Error creating user: $errorMessage" -Level Error
        Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Level Error
        [System.Windows.Forms.MessageBox]::Show("Error creating user: $errorMessage", "Error")
    }
})
$gbTools.Controls.Add($btnCreateUser)

# Pad voor de configuratie file
$script:ConfigPath = Join-Path $PSScriptRoot "ADUserManagement.config.xml"

# Functie om email configuratie op te slaan
function Save-EmailConfig {
    param (
        [string]$SmtpServer,
        [int]$Port,
        [string]$FromAddress,
        [bool]$UseSSL,
        [string]$Username,
        [System.Security.SecureString]$Password
    )
    
    try {
        $config = @{
            SmtpServer = $SmtpServer
            Port = $Port
            FromAddress = $FromAddress
            UseSSL = $UseSSL
            Username = $Username
            # Convert SecureString naar versleutelde string voor opslag
            Password = if ($Password) { 
                ConvertFrom-SecureString -SecureString $Password 
            } else { 
                $null 
            }
        }
        
        # Sla configuratie op in XML
        $config | Export-Clixml -Path $script:ConfigPath -Force
        Write-Log "Email configuration saved successfully" -Level Information
    }
    catch {
        Write-Log "Error saving email configuration: $_" -Level Error
        throw
    }
}

# Functie om email configuratie te laden
function Load-EmailConfig {
    try {
        if (Test-Path $script:ConfigPath) {
            $config = Import-Clixml -Path $script:ConfigPath
            
            # Convert opgeslagen password terug naar SecureString
            if ($config.Password) {
                $config.Password = $config.Password | ConvertTo-SecureString
            }
            
            return $config
        }
        else {
            # Return default configuratie als er geen config file is
            return @{
                SmtpServer = ""
                Port = 25
                FromAddress = ""
                UseSSL = $false
                Username = ""
                Password = $null
            }
        }
    }
    catch {
        Write-Log "Error loading email configuration: $_" -Level Error
        throw
    }
}

# Email Settings Form
function Show-EmailSettings {
    $formEmail = New-Object System.Windows.Forms.Form
    $formEmail.Text = "Email Settings"
    $formEmail.Size = New-Object System.Drawing.Size(400,350)
    $formEmail.StartPosition = "CenterScreen"
    
    # Load huidige configuratie
    $currentConfig = Load-EmailConfig
    
    # SMTP Server
    $lblSmtp = New-Object System.Windows.Forms.Label
    $lblSmtp.Location = New-Object System.Drawing.Point(10,20)
    $lblSmtp.Size = New-Object System.Drawing.Size(100,20)
    $lblSmtp.Text = "SMTP Server:"
    $formEmail.Controls.Add($lblSmtp)
    
    $txtSmtp = New-Object System.Windows.Forms.TextBox
    $txtSmtp.Location = New-Object System.Drawing.Point(120,20)
    $txtSmtp.Size = New-Object System.Drawing.Size(250,20)
    $txtSmtp.Text = $currentConfig.SmtpServer
    $formEmail.Controls.Add($txtSmtp)
    
    # Port
    $lblPort = New-Object System.Windows.Forms.Label
    $lblPort.Location = New-Object System.Drawing.Point(10,50)
    $lblPort.Size = New-Object System.Drawing.Size(100,20)
    $lblPort.Text = "Port:"
    $formEmail.Controls.Add($lblPort)
    
    $txtPort = New-Object System.Windows.Forms.TextBox
    $txtPort.Location = New-Object System.Drawing.Point(120,50)
    $txtPort.Size = New-Object System.Drawing.Size(100,20)
    $txtPort.Text = $currentConfig.Port
    $formEmail.Controls.Add($txtPort)
    
    # From Address
    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Location = New-Object System.Drawing.Point(10,80)
    $lblFrom.Size = New-Object System.Drawing.Size(100,20)
    $lblFrom.Text = "From Address:"
    $formEmail.Controls.Add($lblFrom)
    
    $txtFrom = New-Object System.Windows.Forms.TextBox
    $txtFrom.Location = New-Object System.Drawing.Point(120,80)
    $txtFrom.Size = New-Object System.Drawing.Size(250,20)
    $txtFrom.Text = $currentConfig.FromAddress
    $formEmail.Controls.Add($txtFrom)
    
    # Use SSL
    $chkSSL = New-Object System.Windows.Forms.CheckBox
    $chkSSL.Location = New-Object System.Drawing.Point(120,110)
    $chkSSL.Size = New-Object System.Drawing.Size(250,20)
    $chkSSL.Text = "Use SSL"
    $chkSSL.Checked = $currentConfig.UseSSL
    $formEmail.Controls.Add($chkSSL)
    
    # Username
    $lblUsername = New-Object System.Windows.Forms.Label
    $lblUsername.Location = New-Object System.Drawing.Point(10,140)
    $lblUsername.Size = New-Object System.Drawing.Size(100,20)
    $lblUsername.Text = "Username:"
    $formEmail.Controls.Add($lblUsername)
    
    $txtUsername = New-Object System.Windows.Forms.TextBox
    $txtUsername.Location = New-Object System.Drawing.Point(120,140)
    $txtUsername.Size = New-Object System.Drawing.Size(250,20)
    $txtUsername.Text = $currentConfig.Username
    $formEmail.Controls.Add($txtUsername)
    
    # Password
    $lblPassword = New-Object System.Windows.Forms.Label
    $lblPassword.Location = New-Object System.Drawing.Point(10,170)
    $lblPassword.Size = New-Object System.Drawing.Size(100,20)
    $lblPassword.Text = "Password:"
    $formEmail.Controls.Add($lblPassword)
    
    $txtPassword = New-Object System.Windows.Forms.MaskedTextBox
    $txtPassword.Location = New-Object System.Drawing.Point(120,170)
    $txtPassword.Size = New-Object System.Drawing.Size(250,20)
    $txtPassword.PasswordChar = '*'
    if ($currentConfig.Password) {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($currentConfig.Password)
        $txtPassword.Text = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }
    $formEmail.Controls.Add($txtPassword)
    
    # Save Button
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(120,220)
    $btnSave.Size = New-Object System.Drawing.Size(100,23)
    $btnSave.Text = "Save"
    $btnSave.Add_Click({
        try {
            $securePass = if ($txtPassword.Text) { 
                ConvertTo-SecureString -String $txtPassword.Text -AsPlainText -Force 
            } else { 
                $null 
            }
            
            Save-EmailConfig `
                -SmtpServer $txtSmtp.Text `
                -Port ([int]$txtPort.Text) `
                -FromAddress $txtFrom.Text `
                -UseSSL $chkSSL.Checked `
                -Username $txtUsername.Text `
                -Password $securePass
                
            [System.Windows.Forms.MessageBox]::Show("Settings saved successfully", "Success")
            $formEmail.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving settings: $_", "Error")
        }
    })
    $formEmail.Controls.Add($btnSave)
    
    # Test Button
    $btnTest = New-Object System.Windows.Forms.Button
    $btnTest.Location = New-Object System.Drawing.Point(230,220)
    $btnTest.Size = New-Object System.Drawing.Size(100,23)
    $btnTest.Text = "Test Email"
    $btnTest.Add_Click({
        try {
            $testParams = @{
                SmtpServer = $txtSmtp.Text
                Port = [int]$txtPort.Text
                From = $txtFrom.Text
                To = $txtFrom.Text  # Test naar jezelf
                Subject = "Test Email"
                Body = "This is a test email from AD User Management tool."
                UseSsl = $chkSSL.Checked
            }
            
            if ($txtUsername.Text -and $txtPassword.Text) {
                $securePass = ConvertTo-SecureString -String $txtPassword.Text -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential ($txtUsername.Text, $securePass)
                $testParams.Credential = $credential
            }
            
            Send-MailMessage @testParams
            [System.Windows.Forms.MessageBox]::Show("Test email sent successfully", "Success")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error sending test email: $_", "Error")
        }
    })
    $formEmail.Controls.Add($btnTest)
    
    $formEmail.ShowDialog()
}

# Voeg deze functie toe aan het begin van je script
function Test-FormValues {
    param (
        [string]$Location
    )
    
    Write-Log "=== Form Values Test at $Location ===" -Level Information
    Write-Log "OU Text: [$($txtOU.Text)]" -Level Information
    Write-Log "UPN Suffix Text: [$($cmbUPNSuffix.Text)]" -Level Information
    Write-Log "UPN Suffix SelectedItem: [$($cmbUPNSuffix.SelectedItem)]" -Level Information
    Write-Log "First Name: [$($txtFirstName.Text)]" -Level Information
    Write-Log "Last Name: [$($txtLastName.Text)]" -Level Information
    Write-Log "=== End Form Values Test ===" -Level Information
}

# Update de Initialize-UPNSuffixComboBox functie
function Initialize-UPNSuffixComboBox {
    Write-Log "Starting UPN Suffix initialization" -Level Information
    $cmbUPNSuffix.Items.Clear()
    
    try {
        $upnSuffixes = Get-ADForest | Select-Object -ExpandProperty UPNSuffixes
        Write-Log "Found UPN Suffixes: [$($upnSuffixes -join ', ')]" -Level Information
        
        if ($upnSuffixes.Count -eq 0) {
            Write-Log "No UPN Suffixes found in AD Forest" -Level Warning
            return
        }

        foreach ($suffix in $upnSuffixes) {
            $cmbUPNSuffix.Items.Add($suffix)
            Write-Log "Added UPN Suffix: [$suffix]" -Level Information
        }

        if ($cmbUPNSuffix.Items.Count -gt 0) {
            $cmbUPNSuffix.SelectedIndex = 0
            Write-Log "Set default UPN Suffix: [$($cmbUPNSuffix.Text)]" -Level Information
        }
    }
    catch {
        Write-Log "Error loading UPN Suffixes: $_" -Level Error
    }
}

# Start de applicatie
[System.Windows.Forms.Application]::Run($form)