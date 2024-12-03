# Import benodigde modules
Import-Module ActiveDirectory

# GUI Form maken met Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
            Password = if ($Password) { ConvertFrom-SecureString -SecureString $Password } else { $null }
        }
        
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
            
            if ($config.Password) {
                $config.Password = $config.Password | ConvertTo-SecureString
            }
            
            return $config
        }
        else {
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
                To = $txtFrom.Text
                Subject = "Test Email"
                Body = "This is a test email."
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

# Functie voor het versleutelen van het wachtwoord
function Protect-Password {
    param([string]$PlainPassword)
    $securePassword = ConvertTo-SecureString -String $PlainPassword -AsPlainText -Force
    return ConvertFrom-SecureString $securePassword
}

# Functie voor het ontsleutelen van het wachtwoord
function Unprotect-Password {
    param([string]$EncryptedPassword)
    $securePassword = ConvertTo-SecureString -String $EncryptedPassword
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
    return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}

# Functie voor het laden van de SMTP-instellingen
function Load-SmtpSettings {
    $settingsPath = ".\smtp-settings.xml"
    if (Test-Path $settingsPath) {
        $settings = Import-Clixml -Path $settingsPath
        return $settings
    }
    return @{
        SmtpServer = ""
        FromEmail = ""
        Username = ""
        Password = ""
        Port = "25"
        UseSSL = $false
    }
}

# Functie voor het opslaan van de SMTP-instellingen
function Save-SmtpSettings {
    param(
        $SmtpServer,
        $FromEmail,
        $Username,
        $Password,
        $Port,
        $UseSSL
    )
    
    $settings = @{
        SmtpServer = $SmtpServer
        FromEmail = $FromEmail
        Username = $Username
        Password = Protect-Password -PlainPassword $Password
        Port = $Port
        UseSSL = $UseSSL
    }
    
    $settings | Export-Clixml -Path ".\smtp-settings.xml"
}

# Logging functie
function Write-Log {
    param($Message)
    
    $LogPath = ".\ADPasswordReset.log"
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$TimeStamp - $Message" | Out-File -FilePath $LogPath -Append
}

# Functie voor het genereren van een willekeurig wachtwoord
function New-RandomPassword {
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
    
    # Zorg ervoor dat het wachtwoord aan de minimale vereisten voldoet
    if ($chkNumbers.Checked -and $result -notmatch '\d') {
        $pos = Get-Random -Maximum $length
        $result = $result.Remove($pos, 1).Insert($pos, (Get-Random -Minimum 0 -Maximum 9))
    }
    
    if ($chkSpecialChars.Checked -and $result -notmatch '[^a-zA-Z0-9]') {
        $specialChars = "!@#$%^&*()_+-=[]{}|;:,.<>?"
        $pos = Get-Random -Maximum $length
        $result = $result.Remove($pos, 1).Insert($pos, $specialChars[(Get-Random -Maximum $specialChars.Length)])
    }
    
    return $result
}

# Functie voor het versturen van e-mail
function Send-PasswordEmail {
    param(
        $ToEmail,
        $UserName,
        $Password
    )

    try {
        # Laad de SMTP instellingen
        $settings = Load-EmailConfig
        Write-Log "Loaded SMTP settings: Server=$($settings.SmtpServer), Port=$($settings.Port), From=$($settings.FromAddress), SSL=$($settings.UseSSL)"

        if ([string]::IsNullOrEmpty($settings.SmtpServer)) {
            Write-Log "SMTP Server is not configured"
            throw "SMTP Server is niet geconfigureerd. Configureer eerst de SMTP instellingen."
        }

        if ([string]::IsNullOrEmpty($settings.FromAddress)) {
            Write-Log "From Address is not configured"
            throw "From Address is niet geconfigureerd. Configureer eerst de SMTP instellingen."
        }

        $emailPassword = ""
        if ($settings.Password) {
            $emailPassword = $settings.Password | ConvertTo-SecureString
            Write-Log "SMTP authentication will be used"
        }

        $Subject = "Nieuw wachtwoord voor je account"
        $Body = @"
Beste $UserName,

Je wachtwoord is gereset. Hieronder vind je je nieuwe inloggegevens:

Gebruikersnaam: $UserName
Wachtwoord: $Password

Verander dit wachtwoord bij je eerste aanmelding.

Met vriendelijke groet,
IT Support
"@

        $mailParams = @{
            From = $settings.FromAddress
            To = $ToEmail
            Subject = $Subject
            Body = $Body
            SmtpServer = $settings.SmtpServer
            Port = $settings.Port
            UseSsl = $settings.UseSSL
        }

        # Voeg credentials toe indien geconfigureerd
        if ($settings.Username -and $settings.Password) {
            $credentials = New-Object System.Management.Automation.PSCredential($settings.Username, $emailPassword)
            $mailParams.Add("Credential", $credentials)
            Write-Log "Added credentials for user: $($settings.Username)"
        }

        Write-Log "Attempting to send email to $ToEmail using server $($settings.SmtpServer):$($settings.Port)"
        Send-MailMessage @mailParams
        Write-Log "Email sent successfully"
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        Write-Log "Email error details: $errorDetails"
        throw $errorDetails
    }
}

# Functie voor het ophalen van AD OUs
function Get-ADOUList {
    $domainRoot = (Get-ADDomain).DistinguishedName
    $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $domainRoot | 
           Select-Object Name, DistinguishedName |
           Sort-Object Name
    return $ous
}

# Functie voor het tonen van het gebruikers selectie venster
function Show-UserBrowser {
    $browserForm = New-Object System.Windows.Forms.Form
    $browserForm.Text = "Gebruiker Selecteren"
    $browserForm.Size = New-Object System.Drawing.Size(600,500)
    $browserForm.StartPosition = "CenterScreen"

    # Zoekbalk
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(20,20)
    $lblSearch.Size = New-Object System.Drawing.Size(100,20)
    $lblSearch.Text = "Zoeken:"
    $browserForm.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(120,20)
    $txtSearch.Size = New-Object System.Drawing.Size(200,20)
    $browserForm.Controls.Add($txtSearch)

    # ListView voor gebruikers
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(20,50)
    $listView.Size = New-Object System.Drawing.Size(540,350)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.MultiSelect = $false
    $listView.Columns.Add("Naam", 200)
    $listView.Columns.Add("Gebruikersnaam", 150)
    $listView.Columns.Add("OU", 190)
    $browserForm.Controls.Add($listView)

    # Functie om gebruikers te laden
    function Load-Users {
        param($SearchText)
        $listView.Items.Clear()
        
        $filter = "*"
        if ($SearchText) {
            $filter = "*$SearchText*"
        }

        Get-ADUser -Filter "DisplayName -like '$filter' -or SamAccountName -like '$filter'" -Properties DisplayName, DistinguishedName | 
        Sort-Object DisplayName | 
        ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem($_.DisplayName)
            $item.SubItems.Add($_.SamAccountName)
            $ou = ($_.DistinguishedName -split ',',2)[1]
            $item.SubItems.Add($ou)
            $item.Tag = $_
            $listView.Items.Add($item)
        }
    }

    # Zoek knop
    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Location = New-Object System.Drawing.Point(330,20)
    $btnSearch.Size = New-Object System.Drawing.Size(100,20)
    $btnSearch.Text = "Zoeken"
    $btnSearch.Add_Click({ Load-Users $txtSearch.Text })
    $browserForm.Controls.Add($btnSearch)

    # Enter toets in zoekbalk
    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            Load-Users $txtSearch.Text
            $_.SuppressKeyPress = $true
        }
    })

    # Selecteer knop
    $btnSelect = New-Object System.Windows.Forms.Button
    $btnSelect.Location = New-Object System.Drawing.Point(20,410)
    $btnSelect.Size = New-Object System.Drawing.Size(200,30)
    $btnSelect.Text = "Selecteer Gebruiker"
    $btnSelect.Add_Click({
        if ($listView.SelectedItems.Count -eq 1) {
            $script:selectedUser = $listView.SelectedItems[0].Tag
            $browserForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $browserForm.Close()
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Selecteer een gebruiker.", "Waarschuwing")
        }
    })
    $browserForm.Controls.Add($btnSelect)

    # Laad initiële lijst
    Load-Users ""

    # Toon het venster en wacht op resultaat
    $result = $browserForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $script:selectedUser
    }
    return $null
}

# Maak het hoofdvenster
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Wachtwoord Reset Tool"
$form.Size = New-Object System.Drawing.Size(500,450)
$form.StartPosition = "CenterScreen"

# Browse knop (verplaatst naar boven)
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(20,20)
$btnBrowse.Size = New-Object System.Drawing.Size(200,30)
$btnBrowse.Text = "Zoek Gebruiker..."
$btnBrowse.Add_Click({
    $selectedUser = Show-UserBrowser
    if ($selectedUser) {
        $script:currentUser = $selectedUser
        $lblSelectedUser.Text = "Geselecteerde gebruiker: $($selectedUser.DisplayName) ($($selectedUser.SamAccountName))"
    }
})
$form.Controls.Add($btnBrowse)

# Geselecteerde gebruiker label (verplaatst)
$lblSelectedUser = New-Object System.Windows.Forms.Label
$lblSelectedUser.Location = New-Object System.Drawing.Point(20,60)
$lblSelectedUser.Size = New-Object System.Drawing.Size(400,20)
$lblSelectedUser.Text = "Geselecteerde gebruiker: Geen"
$form.Controls.Add($lblSelectedUser)

# Extra e-mailadres label en tekstveld (verplaatst)
$lblExtraEmail = New-Object System.Windows.Forms.Label
$lblExtraEmail.Location = New-Object System.Drawing.Point(20,90)
$lblExtraEmail.Size = New-Object System.Drawing.Size(100,20)
$lblExtraEmail.Text = "Extra E-mail:"
$form.Controls.Add($lblExtraEmail)

$txtExtraEmail = New-Object System.Windows.Forms.TextBox
$txtExtraEmail.Location = New-Object System.Drawing.Point(120,90)
$txtExtraEmail.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($txtExtraEmail)

# Reset knop (verplaatst)
$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Location = New-Object System.Drawing.Point(20,120)
$btnReset.Size = New-Object System.Drawing.Size(200,30)
$btnReset.Text = "Reset Wachtwoord"
$btnReset.Add_Click({
    if (-not $script:currentUser) {
        [System.Windows.Forms.MessageBox]::Show("Selecteer eerst een gebruiker.", "Fout")
        return
    }

    $username = $script:currentUser.SamAccountName
    $extraEmail = $txtExtraEmail.Text

    try {
        # Zoek de AD-gebruiker met mail property
        $adUser = Get-ADUser -Identity $username -Properties mail
        Write-Log "AD Email address found: $($adUser.mail)"  # Log het gevonden e-mailadres
        
        $newPassword = New-RandomPassword
        
        # Reset het wachtwoord
        Set-ADAccountPassword -Identity $adUser -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force)
        
        # Pas de account opties toe
        Set-ADUser -Identity $adUser -ChangePasswordAtLogon $chkResetRequired.Checked
        Set-ADUser -Identity $adUser -PasswordNeverExpires $chkNeverExpires.Checked
        
        Write-Log "Wachtwoord reset voor gebruiker: $username"
        Write-Log "Reset Required: $($chkResetRequired.Checked), Never Expires: $($chkNeverExpires.Checked)"

        $emailsSent = 0
        $emailErrors = @()

        # Verstuur e-mail naar het AD e-mailadres
        if ($adUser.mail) {
            Write-Log "Attempting to send email to AD email: $($adUser.mail)"
            try {
                $emailSent = Send-PasswordEmail -ToEmail $adUser.mail -UserName $username -Password $newPassword
                $emailsSent++
                Write-Log "E-mail verzonden naar AD e-mailadres: $($adUser.mail)"
            } catch {
                $emailErrors += "AD e-mail ($($adUser.mail)): $($_.Exception.Message)"
                Write-Log "Fout bij verzenden e-mail naar AD e-mailadres: $($_.Exception.Message)"
            }
        } else {
            Write-Log "Geen e-mailadres gevonden in AD voor gebruiker $username"
        }

        # Verstuur e-mail naar extra e-mailadres indien ingevuld
        if (-not [string]::IsNullOrEmpty($extraEmail)) {
            Write-Log "Attempting to send email to extra email: $extraEmail"
            try {
                $emailSent = Send-PasswordEmail -ToEmail $extraEmail -UserName $username -Password $newPassword
                $emailsSent++
                Write-Log "E-mail verzonden naar extra e-mailadres: $extraEmail"
            } catch {
                $emailErrors += "Extra e-mail ($extraEmail): $($_.Exception.Message)"
                Write-Log "Fout bij verzenden e-mail naar extra e-mailadres: $($_.Exception.Message)"
            }
        }

        # Toon een gedetailleerd resultaat
        $message = "Wachtwoord succesvol gereset.`n`n"
        if ($emailsSent -gt 0) {
            $message += "E-mails verzonden: $emailsSent`n"
        } else {
            $message += "Geen e-mails verzonden.`n"
        }
        
        if ($emailErrors.Count -gt 0) {
            $message += "`nFouten bij verzenden naar:`n"
            $message += $emailErrors -join "`n"
        }

        if (-not $adUser.mail) {
            $message += "`n`nLet op: Geen e-mailadres gevonden in AD voor deze gebruiker."
        }

        $message += "`n`nNieuw wachtwoord: $newPassword"
        
        [System.Windows.Forms.MessageBox]::Show($message, "Reset Resultaat")
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Fout bij wachtwoord reset voor $username : $errorMessage"
        [System.Windows.Forms.MessageBox]::Show("Fout bij het resetten van het wachtwoord: $errorMessage", "Fout")
    }
})
$form.Controls.Add($btnReset)

# Password Policy GroupBox
$gbPasswordPolicy = New-Object System.Windows.Forms.GroupBox
$gbPasswordPolicy.Location = New-Object System.Drawing.Point(20,160)
$gbPasswordPolicy.Size = New-Object System.Drawing.Size(440,100)
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

# Password Options GroupBox
$gbPasswordOptions = New-Object System.Windows.Forms.GroupBox
$gbPasswordOptions.Location = New-Object System.Drawing.Point(20,270)  # Positie onder Password Policy
$gbPasswordOptions.Size = New-Object System.Drawing.Size(440,80)
$gbPasswordOptions.Text = "Password Options"
$form.Controls.Add($gbPasswordOptions)

# Reset Password at Next Logon
$chkResetRequired = New-Object System.Windows.Forms.CheckBox
$chkResetRequired.Location = New-Object System.Drawing.Point(10,25)
$chkResetRequired.Size = New-Object System.Drawing.Size(200,20)
$chkResetRequired.Text = "Reset Required at Next Logon"
$chkResetRequired.Checked = $true  # Standaard aangevinkt
$gbPasswordOptions.Controls.Add($chkResetRequired)

# Password Never Expires
$chkNeverExpires = New-Object System.Windows.Forms.CheckBox
$chkNeverExpires.Location = New-Object System.Drawing.Point(220,25)
$chkNeverExpires.Size = New-Object System.Drawing.Size(200,20)
$chkNeverExpires.Text = "Password Never Expires"
$chkNeverExpires.Checked = $false  # Standaard uitgevinkt
$gbPasswordOptions.Controls.Add($chkNeverExpires)

# SMTP instellingen knop (verplaatst naar onder)
$btnSmtpSettings = New-Object System.Windows.Forms.Button
$btnSmtpSettings.Location = New-Object System.Drawing.Point(20,360)  # Nieuwe Y-positie
$btnSmtpSettings.Size = New-Object System.Drawing.Size(200,30)
$btnSmtpSettings.Text = "SMTP Instellingen"
$btnSmtpSettings.Add_Click({ Show-EmailSettings })
$form.Controls.Add($btnSmtpSettings)

# Status label (verplaatst)
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(230,365)  # Nieuwe Y-positie
$lblStatus.Size = New-Object System.Drawing.Size(230,20)
$lblStatus.Text = "Status: Gereed"
$form.Controls.Add($lblStatus)

# Start de applicatie
$form.ShowDialog()
