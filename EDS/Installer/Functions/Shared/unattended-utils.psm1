function Save-UnattendedXml {
    param(
        [Parameter(Mandatory=$true)]
        [xml]$xmlDoc,
        [Parameter(Mandatory=$true)]
        [string]$filePath
    )
    # Use UTF-8 without BOM to avoid encoding issues
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $xmlWriterSettings = New-Object System.Xml.XmlWriterSettings
    $xmlWriterSettings.Indent = $true
    $xmlWriterSettings.Encoding = $utf8NoBom
    $xmlWriterSettings.OmitXmlDeclaration = $false
    $xmlWriter = [System.Xml.XmlWriter]::Create($filePath, $xmlWriterSettings)
    try {
        $xmlDoc.save($xmlWriter)
    } finally {
        $xmlWriter.Close()
    }
}

function Set-DefaultUnattendedXML {
    param(
        [string]$EDSFolderName = "EDS",
        [string]$WinPeDrive = "X:"
    )
    [OutputType([xml])]
    
    [xml]$xmlDoc;

    try {
        $installDrive = Get-InstallationDrive -EDSFolderName $EDSFolderName
        if (-not $installDrive) {
            throw "Installation media not found. Please ensure the USB drive is properly connected."
        }

        # Create TEMP directory if it doesn't exist
        $tempDir = Join-Path $WinPeDrive "Temp" 
        $tempXmlPath = Join-Path $tempDir "unattended.xml"
        $script:unattendPath = $tempXmlPath
        $unattendPath = $tempXmlPath
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }
        $sourceXmlPath = Join-Path $installDrive "$EDSFolderName\Installer\unattended.xml"
        Write-Host "Looking for unattended.xml in $sourceXmlPath"

        if (Test-Path $sourceXmlPath) {
            Write-Host "Existing unattended.xml found, modifying it..."
            $xmlDoc = [xml](Get-Content -Path $sourceXmlPath)
        } else {
            Write-Host "No existing unattended.xml found, creating default one..."
            $xmlDoc = [xml](New-Object System.Xml.XmlDocument)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error updating unattended.xml: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        throw "Could not load unattended XML."
    }

    # Create basic unattend structure if it doesn't exist
    if (-not $xmlDoc.DocumentElement) {
        $root = $xmlDoc.CreateElement("unattend", "urn:schemas-microsoft-com:unattend")
        # Add wcm namespace
        $wcmAttr = $xmlDoc.CreateAttribute("xmlns:wcm")
        $wcmAttr.Value = "http://schemas.microsoft.com/WMIConfig/2002/State"
        $root.Attributes.Append($wcmAttr) | Out-Null
        $xmlDoc.AppendChild($root) | Out-Null
    }

    $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsMgr.AddNamespace("u", "urn:schemas-microsoft-com:unattend")

    # Use correct XPath with namespace
    $settingsList = $xmlDoc.SelectNodes("//u:settings[@pass='specialize']", $nsMgr)
    if ($settingsList.Count -gt 0) {
        $settings = $settingsList[0]
    } else {
        $settings = $xmlDoc.CreateElement("settings", "urn:schemas-microsoft-com:unattend")
        $settings.SetAttribute("pass", "specialize")
        $xmlDoc.DocumentElement.AppendChild($settings) | Out-Null  # FIX: use DocumentElement
    }

    # Create component if it doesn't exist
    $component = $settings.SelectSingleNode("u:component[@name='Microsoft-Windows-Shell-Setup']", $nsMgr)
    if (-not $component) {
        $component = $xmlDoc.CreateElement("component", "urn:schemas-microsoft-com:unattend")
        $component.SetAttribute("name", "Microsoft-Windows-Shell-Setup")
        $component.SetAttribute("processorArchitecture", "amd64")
        $component.SetAttribute("publicKeyToken", "31bf3856ad364e35")
        $component.SetAttribute("language", "neutral")
        $component.SetAttribute("versionScope", "nonSxS")
        $settings.AppendChild($component) | Out-Null
    }


    # Add RunSynchronous commands component if it doesn't exist
    $runComponent = $settings.SelectSingleNode("u:component[@name='Microsoft-Windows-Deployment']", $nsMgr)
    if (-not $runComponent) {
        $runComponent = $xmlDoc.CreateElement("component", "urn:schemas-microsoft-com:unattend")
        $runComponent.SetAttribute("name", "Microsoft-Windows-Deployment")
        $runComponent.SetAttribute("processorArchitecture", "amd64")
        $runComponent.SetAttribute("publicKeyToken", "31bf3856ad364e35")
        $runComponent.SetAttribute("language", "neutral")
        $runComponent.SetAttribute("versionScope", "nonSxS")
        $settings.AppendChild($runComponent) | Out-Null
    }

    # Create RunSynchronous element if it doesn't exist
    $runSync = $runComponent.SelectSingleNode("u:RunSynchronous", $nsMgr)
    if (-not $runSync) {
        $runSync = $xmlDoc.CreateElement("RunSynchronous", "urn:schemas-microsoft-com:unattend")
        $runComponent.AppendChild($runSync) | Out-Null
    }

    $existingOrders = @()
    foreach ($cmd in $runSync.SelectNodes("u:RunSynchronousCommand", $nsMgr)) {
        $orderNode = $cmd.SelectSingleNode("u:Order", $nsMgr)
        if ($orderNode -and $orderNode.InnerText -match '^\\d+$') {
            $existingOrders += [int]$orderNode.InnerText
        }
    }
    if ($existingOrders.Count -eq 0) { $maxOrder = 0 } else { $maxOrder = ($existingOrders | Measure-Object -Maximum).Maximum }

    $wcmNamespaceUri = "http://schemas.microsoft.com/WMIConfig/2002/State"
    $wcmAttr = $xmlDoc.CreateAttribute("wcm", "action", $wcmNamespaceUri)
    $wcmAttr.Value = "add"

    # Find the next available <Order> value for RunSynchronousCommand
    $existingOrders = @()
    $runSyncCmds = $runSync.SelectNodes("u:RunSynchronousCommand/u:Order", $nsMgr)
    if ($runSyncCmds) {
        foreach ($orderNode in $runSyncCmds) {
            [int]$val = 0
            if ([int]::TryParse($orderNode.InnerText, [ref]$val)) {
                $existingOrders += $val
            }
        }
    }
    if ($existingOrders.Count -eq 0) { $existingOrders = @(0) }
    $nextOrder = (($existingOrders | Measure-Object -Maximum).Maximum) + 1

    Write-Host "Next available order for RunSynchronousCommand is $nextOrder"

    # Add command to extract Specialize.ps1 (keep this in specialize)
    $extractCommand = $xmlDoc.CreateElement("RunSynchronousCommand", "urn:schemas-microsoft-com:unattend")
    $extractCommand.SetAttributeNode($wcmAttr)
    $path = $xmlDoc.CreateElement("Path", "urn:schemas-microsoft-com:unattend")
    $path.InnerText = "powershell.exe -WindowStyle Normal -NoProfile -Command `"`$xml = [xml]::new(); `$xml.Load('C:\Windows\Panther\unattend.xml'); `$sb = [scriptblock]::Create( `$xml.unattend.EDS.CopyScript ); Invoke-Command -ScriptBlock `$sb -ArgumentList $EDSFolderName;`""
    $description = $xmlDoc.CreateElement("Description", "urn:schemas-microsoft-com:unattend")
    $description.InnerText = "Execute CopySpecialize Script embedded inside unattend.xml"
    $order = $xmlDoc.CreateElement("Order", "urn:schemas-microsoft-com:unattend")
    $order.InnerText = "$nextOrder"
    $extractCommand.AppendChild($path) | Out-Null
    $extractCommand.AppendChild($description) | Out-Null
    $extractCommand.AppendChild($order) | Out-Null
    $runSync.AppendChild($extractCommand) | Out-Null

    # DO NOT add Specialize.ps1 execution to specialize RunSynchronous anymore!
    # Instead, add it to oobeSystem/FirstLogonCommands
    # Find oobeSystem settings/component
    $oobeSettings = $xmlDoc.SelectSingleNode("//u:settings[@pass='oobeSystem']", $nsMgr)
    if (-not $oobeSettings) {
        $oobeSettings = $xmlDoc.CreateElement("settings", "urn:schemas-microsoft-com:unattend")
        $oobeSettings.SetAttribute("pass", "oobeSystem")
        $xmlDoc.DocumentElement.AppendChild($oobeSettings) | Out-Null
    }
    $shellComponent = $oobeSettings.SelectSingleNode("u:component[@name='Microsoft-Windows-Shell-Setup']", $nsMgr)
    if (-not $shellComponent) {
        $shellComponent = $xmlDoc.CreateElement("component", "urn:schemas-microsoft-com:unattend")
        $shellComponent.SetAttribute("name", "Microsoft-Windows-Shell-Setup")
        $shellComponent.SetAttribute("processorArchitecture", "amd64")
        $shellComponent.SetAttribute("publicKeyToken", "31bf3856ad364e35")
        $shellComponent.SetAttribute("language", "neutral")
        $shellComponent.SetAttribute("versionScope", "nonSxS")
        $oobeSettings.AppendChild($shellComponent) | Out-Null
    }
    $firstLogon = $shellComponent.SelectSingleNode("u:FirstLogonCommands", $nsMgr)
    if (-not $firstLogon) {
        $firstLogon = $xmlDoc.CreateElement("FirstLogonCommands", "urn:schemas-microsoft-com:unattend")
        $shellComponent.AppendChild($firstLogon) | Out-Null
    }
    # Find max order
    $orders = @()
    foreach ($cmd in $firstLogon.SelectNodes("u:SynchronousCommand/u:Order", $nsMgr)) {
        if ($cmd.InnerText -match '^\d+$') { $orders += [int]$cmd.InnerText }
    }
    $nextOrderFL = ($orders | Measure-Object -Maximum).Maximum
    if (-not $nextOrderFL) { $nextOrderFL = 0 }
    $nextOrderFL++
    # Add new SynchronousCommand for Specialize.ps1
    $syncCmd = $xmlDoc.CreateElement("SynchronousCommand", "urn:schemas-microsoft-com:unattend")
    $syncCmd.SetAttribute("action", "http://schemas.microsoft.com/WMIConfig/2002/State", "add")
    $order = $xmlDoc.CreateElement("Order", "urn:schemas-microsoft-com:unattend")
    $order.InnerText = "$nextOrderFL"
    $cmdLine = $xmlDoc.CreateElement("CommandLine", "urn:schemas-microsoft-com:unattend")
    $cmdLine.InnerText = "powershell.exe -ExecutionPolicy Bypass -File C:\Windows\Setup\$EDSFolderName\Specialize.ps1"
    $syncCmd.AppendChild($order) | Out-Null
    $syncCmd.AppendChild($cmdLine) | Out-Null
    $firstLogon.AppendChild($syncCmd) | Out-Null

    # Create EDS element if it doesn't exist
    $eds = $xmlDoc.DocumentElement.SelectSingleNode("EDS")  # FIX: use DocumentElement
    if (-not $eds) {
        $eds = $xmlDoc.CreateElement("EDS", "https://eds.cwi.at")
        $xmlDoc.DocumentElement.AppendChild($eds) | Out-Null
    }

    # Add CopyScript section inside EDS
    $copyScript = $xmlDoc.CreateElement("CopyScript",$eds.NamespaceURI)
    $copyScript.InnerText = Get-Content -Path "$installDrive\$EDSFolderName\Installer\Functions\CopySpecialize.ps1" -Raw
    $eds.AppendChild($copyScript) | Out-Null

    # When saving, ensure XML declaration is present
    Save-UnattendedXml -xmlDoc $xmlDoc -filePath $tempXmlPath
    [xml]$script:unattendXml = Get-Content -Path $tempXmlPath -Raw
}

function Set-UnattendedDeviceName {
    param (
        [Parameter(Mandatory=$true)]
        [xml]$xmlDoc,
        [Parameter(Mandatory=$true)]
        [string]$deviceName
    )

    $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsMgr.AddNamespace("u", "urn:schemas-microsoft-com:unattend") | Out-Null

    # Get the component from the XML document
    $component = $xmlDoc.SelectSingleNode("//u:settings[@pass='specialize']/u:component[@name='Microsoft-Windows-Shell-Setup']", $nsMgr)
    if (-not $component) {
        Write-Warning "Required component not found in XML"
        return $false
    }

    # Create or update ComputerName element
    $computerName = $component.SelectSingleNode("u:ComputerName", $nsMgr)
    if (-not $computerName) {
        $computerName = $xmlDoc.CreateElement("ComputerName", "urn:schemas-microsoft-com:unattend")
        $component.AppendChild($computerName) | Out-Null
    }
    $computerName.InnerText = $deviceName

    # Use helper function to maintain encoding
    Save-UnattendedXml -xmlDoc $xmlDoc -filePath $script:unattendPath
    return $true
}


function Set-UnattendedUserInput {
    param (
        [Parameter(Mandatory=$true)]
        [xml]$xmlDoc,
        [Parameter(Mandatory=$true)]
        [Hashtable]$UserInput
    )

    Write-Host "Saving userInput..."

    $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsMgr.AddNamespace("u", "urn:schemas-microsoft-com:unattend")
    $nsMgr.AddNamespace("e", "https://eds.cwi.at")

    # Query the <EDS> node using XPath with the correct namespace
    $eds = $xmlDoc.SelectSingleNode("//e:EDS", $nsMgr)

    # Add or update settings under EDS
    $userInputBlock = $eds.SelectSingleNode("UserInput")
    if (-not $userInputBlock) {
        $userInputBlock = $xmlDoc.CreateElement("UserInput", $eds.NamespaceURI)
        $eds.AppendChild($userInputBlock)  | Out-Null
    }

    if ($UserInput.ContainsKey('localPassword')) {
        $UserInput.Remove('localPassword')
    }

    # Iterate through the UserInput hashtable and create elements
    foreach ($key in $UserInput.Keys) {
        $value = $UserInput[$key]

        # Check if the element already exists
        $element = $userInputBlock.SelectSingleNode($key)
        if (-not $element) {
            $element = $xmlDoc.CreateElement($key, $eds.NamespaceURI)
            $userInputBlock.AppendChild($element)  | Out-Null
        }
        $element.InnerText = $value
    }

    try {
        # Use helper function to maintain encoding
        Save-UnattendedXml -xmlDoc $xmlDoc -filePath $script:unattendPath
    } catch {
        Write-Warning "Failed to save XML: $_"
        return $false
    }
}

function Set-LocalAccount {
    param(
        [Parameter(Mandatory=$true)]
        [xml]$xmlDoc,
        [Parameter(Mandatory=$true)]
        [string]$UserName,
        [Parameter(Mandatory=$true)]
        [string]$PasswordBase64
    )

    $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsMgr.AddNamespace("u", "urn:schemas-microsoft-com:unattend") | Out-Null
    $nsMgr.AddNamespace("wcm", "http://schemas.microsoft.com/WMIConfig/2002/State") | Out-Null

    # Find the oobeSystem settings/component
    $settings = $xmlDoc.SelectSingleNode("//u:settings[@pass='oobeSystem']", $nsMgr)
    if (-not $settings) {
        $settings = $xmlDoc.CreateElement("settings", "urn:schemas-microsoft-com:unattend")
        $settings.SetAttribute("pass", "oobeSystem")
        $xmlDoc.DocumentElement.AppendChild($settings) | Out-Null
    }
    $component = $settings.SelectSingleNode("u:component[@name='Microsoft-Windows-Shell-Setup']", $nsMgr)
    if (-not $component) {
        $component = $xmlDoc.CreateElement("component", "urn:schemas-microsoft-com:unattend")
        $component.SetAttribute("name", "Microsoft-Windows-Shell-Setup")
        $component.SetAttribute("processorArchitecture", "amd64")
        $component.SetAttribute("publicKeyToken", "31bf3856ad364e35")
        $component.SetAttribute("language", "neutral")
        $component.SetAttribute("versionScope", "nonSxS")
        $settings.AppendChild($component) | Out-Null
    }
    # Find or create <UserAccounts>/<LocalAccounts>
    $userAccounts = $component.SelectSingleNode("u:UserAccounts", $nsMgr)
    if (-not $userAccounts) {
        $userAccounts = $xmlDoc.CreateElement("UserAccounts", "urn:schemas-microsoft-com:unattend")
        $component.AppendChild($userAccounts) | Out-Null
    }
    $localAccounts = $userAccounts.SelectSingleNode("u:LocalAccounts", $nsMgr)
    if (-not $localAccounts) {
        $localAccounts = $xmlDoc.CreateElement("LocalAccounts", "urn:schemas-microsoft-com:unattend")
        $userAccounts.AppendChild($localAccounts) | Out-Null
    }
    # Check for existing LocalAccount with the username
    $existingAccount = $localAccounts.SelectSingleNode("u:LocalAccount[u:Name='$UserName']", $nsMgr)
    if ($existingAccount) {
        # Update username and password
        $existingAccount.SelectSingleNode("u:Name", $nsMgr).InnerText = $UserName
        $pwNode = $existingAccount.SelectSingleNode("u:Password/u:Value", $nsMgr)
        if ($pwNode) {
            $pwNode.InnerText = $PasswordBase64
        } else {
            # Add password node if missing
            $pwParent = $xmlDoc.CreateElement("Password", "urn:schemas-microsoft-com:unattend")
            $pwValue = $xmlDoc.CreateElement("Value", "urn:schemas-microsoft-com:unattend")
            $pwValue.InnerText = $PasswordBase64
            $pwPlain = $xmlDoc.CreateElement("PlainText", "urn:schemas-microsoft-com:unattend")
            $pwPlain.InnerText = "false"
            $pwParent.AppendChild($pwValue) | Out-Null
            $pwParent.AppendChild($pwPlain) | Out-Null
            $existingAccount.AppendChild($pwParent) | Out-Null
        }
    } else {
        # Create new LocalAccount
        $localAccount = $xmlDoc.CreateElement("LocalAccount", "urn:schemas-microsoft-com:unattend")
        $wcmAttr = $xmlDoc.CreateAttribute("wcm", "action", "http://schemas.microsoft.com/WMIConfig/2002/State")
        $wcmAttr.Value = "add"
        $localAccount.SetAttributeNode($wcmAttr)
        $name = $xmlDoc.CreateElement("Name", "urn:schemas-microsoft-com:unattend")
        $name.InnerText = $UserName
        $displayName = $xmlDoc.CreateElement("DisplayName", "urn:schemas-microsoft-com:unattend")
        $displayName.InnerText = $UserName
        $group = $xmlDoc.CreateElement("Group", "urn:schemas-microsoft-com:unattend")
        $group.InnerText = "Administrators"
        $pwParent = $xmlDoc.CreateElement("Password", "urn:schemas-microsoft-com:unattend")
        $pwValue = $xmlDoc.CreateElement("Value", "urn:schemas-microsoft-com:unattend")
        $pwValue.InnerText = $PasswordBase64
        $pwPlain = $xmlDoc.CreateElement("PlainText", "urn:schemas-microsoft-com:unattend")
        $pwPlain.InnerText = "false"
        $pwParent.AppendChild($pwValue) | Out-Null
        $pwParent.AppendChild($pwPlain) | Out-Null
        $localAccount.AppendChild($name) | Out-Null
        $localAccount.AppendChild($displayName) | Out-Null
        $localAccount.AppendChild($group) | Out-Null
        $localAccount.AppendChild($pwParent) | Out-Null
        $localAccounts.AppendChild($localAccount) | Out-Null
    }
    
    # Use helper function to maintain encoding
    Save-UnattendedXml -xmlDoc $xmlDoc -filePath $script:unattendPath
    return $true
}

Export-ModuleMember -Function *