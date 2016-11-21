<#
    .SYNOPSIS
        Active Directory User Creation Tool v1.4

    .DESCRIPTION
        A Powershell-GUI tool for creating AD user accounts, single or bulk.

    .INPUTS
        ANUC.Options.XML: Settings and default options file.
        Form input

    .OUTPUTS
        ANUC.log: Log file created under \logs folder.

    .LINK
        https://github.com/feldrim/ANUC

    .NOTES
        Mainly based on the script by Rich Prescott, this script includes updates by 
        Gabriel Jensen and Jim Smith.

    .REQUIREMENTS
        - PSLogging Module - http://9to5it.com/powershell-logging-v2-easily-create-log-files/
        - Active Directory Module
        - Exchange 2010 Snapin

    #> #Version 1.x


#For debugging
$oldVerbosePreference = $VerbosePreference
#$VerbosePreference = "Continue";

$XMLOptions = "ANUC.Options.xml" #change to desired XML file name

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#For the PSScriptRoot command not to take the root as PSLogging module but the ANUC script
$initialLogPath = $PSScriptRoot + '\logs\' #Version 1.x

#----------------------------------------------
#region Import Main Assemblies
#----------------------------------------------
LoadModule ActiveDirectory #Version 1.x
LoadSnapin Microsoft.Exchange.Management.PowerShell.E2010 #Version 1.x
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
#endregion Import Main Assemblies
#----------------------------------------------

#----------------------------------------------
#region Import Logging Assemblies "1.x" #Version 1.x
#----------------------------------------------
LoadModule PSLogging
$sScriptVersion = $XML.Options.Version
$sLogPath = $initialLogPath
$sLogName = 'ANUC.log'
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#Create log folder if not exist
New-Item -ItemType Directory -Force -Path $initialLogPath #Version 1.x

#endregion Import Main Assemblies
#----------------------------------------------

#----------------------------------------------
#region Application Functions
#----------------------------------------------
function LoadModule($moduleName){ #Version 1.x
    if(!(Get-Module -List $moduleName) ) {
        Write-LogWarning -LogPath $sLogFile -Message "Couldn't locate $moduleName Module."
        } else{
            Import-Module $moduleName
        }
    }

    function LoadSnapin($snapinName){ #Version 1.x
        if((Get-PSSnapin -Name $snapinName) -eq $null) {
            Add-PSSnapin $snapinName 
            } else{
                Write-LogWarning -LogPath $sLogFile -Message "$snapinName Snapin already exists."
            }
        }
        function OnApplicationLoad {
            
            Start-Log -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion #Version 1.x

$CreateXML = @"
<?xml version="1.0" standalone="no"?>
<OPTIONS Product="AD New User Creation" Version="1.4">
  <Settings>
    <sAMAccountName Generate="True">
      <Style Format="FirstName.LastName" Enabled="False" />
      <Style Format="FirstInitialLastName" Enabled="True" /> 
      <Style Format="LastNameFirstInitial" Enabled="False" />
    </sAMAccountName>
    <UPN Generate="True">
      <Style Format="FirstName.LastName" Enabled="False" />
      <Style Format="FirstInitialLastName" Enabled="True" />
      <Style Format="LastNameFirstInitial" Enabled="False" />
    </UPN>
    <DisplayName Generate="True">
      <Style Format="FirstName LastName" Enabled="True" />
      <Style Format="LastName, FirstName" Enabled="False" />
    </DisplayName>
    <LowerCaseUserNames Enabled="True" />
    <AccountStatus Enabled="True" />
    <Password ChangeAtLogon="True" />
    <DomainController>YOUR.DOMAIN-CONTROLLER.COM</DomainController>
    <DomainNS>YOURDOMAIN</DomainNS>
    <ScriptPath>LOGIN_SCRIPT.bat</ScriptPath>
    <UserDirectory>\\SERVER\Users\</UserDirectory>
    <Subfolders>
      <Subfolder>Temp</Subfolder>
      <Subfolder>Business</Subfolder>
    </Subfolders>
    <HomeDrive>U:</HomeDrive>
    <HomePage>http://your.homepage.com/</HomePage>
    <Email EnableHtml="True">
    	<Administrator>admin@anuc.com</Administrator>
    	<Supervisor>supervisor@anuc.com</Supervisor>
    	<SubjectToSupervisor>New User Created</SubjectToSupervisor>
    	<SubjectToUser>Welcome</SubjectToUser>
    </Email>
    <SMTPServer>SMTP.awesome.local</SMTPServer>
  </Settings>
  <Default>
    <Domain>awesome.local</Domain>
    <Path>OU=MyOU,DC=awesome,DC=local</Path>
    <FirstName></FirstName>
    <LastName></LastName>
    <Office></Office>
    <Title></Title>
    <Description>Full-Time Employee</Description>
    <Department>IT</Department>
    <Company>Awesome Inc.</Company>
    <Site>TN</Site>
    <Country>US</Country>
    <Password>P@ssw0rd</Password>
    <Group>Normal User</Group>
  </Default>
  <Locations>
    <Location Site="TN">
      <StreetAddress>1 Main Street</StreetAddress>
      <City>Nashville</City>
      <State>TN</State>
      <PostalCode>10001</PostalCode>
      <Phone>888-555-0000</Phone>
      <Fax>888-555-0000</Fax>
    </Location>
    <Location Site="Custom">
      <StreetAddress></StreetAddress>
      <City></City>
      <State></State>
      <PostalCode></PostalCode>
      <Phone></Phone>
      <Fax></Fax>
    </Location>
  </Locations>
  <Domains>
    <Domain Name="awesome.local">
      <Path>OU=MyOU,DC=awesome,DC=local</Path>
      <Path>CN=Users,DC=awesome,DC=local</Path>
    </Domain>
    <Domain Name="awesome.lab">
      <Path>OU=RPUsers1,DC=awesome,DC=lab</Path>
      <Path>OU=RPUsers2,DC=awesome,DC=lab</Path>
      <Path>OU=RPUsers3,DC=awesome,DC=lab</Path>
    </Domain>
  </Domains>
  <Descriptions>
    <Description>Full-Time Employee</Description>
    <Description>Part-Time Employee</Description>
    <Description>Consultant</Description>
    <Description>Intern</Description>
    <Description>Service Account</Description>
    <Description>Temp</Description>
    <Description>Freelancer</Description>
  </Descriptions>
  <Departments>
    <Department>Finance</Department>
    <Department>IT</Department>
    <Department>Marketing</Department>
    <Department>Sales</Department>
    <Department>Executive</Department>
    <Department>Human Resources</Department>
    <Department>Security</Department>
  </Departments>
  <JobTitles>
    <JobTitle>Accountant</JobTitle>
    <JobTitle>Project Manager</JobTitle>
    <JobTitle>Intern</JobTitle>
    <JobTitle>Office Administrator</JobTitle>
  </JobTitles>
  <Groups>
    <Group Name="Normal User">
      <List Type="SecurityGroup">Awesome Users</List>
      <List Type="ComboGroup">Security and Distribution List</List>
      <List Type="DistributionList">Awesome List</List>
    </Group>
    <Group Name="Administrator">
      <List Type="SecurityGroup">Admin Users</List>
      <List Type="SecurityGroup">Awesome Users</List>
      <List Type="ComboGroup">Security and Distribution List</List>
      <List Type="DistributionList">Awesome List</List>
    </Group>
  </Groups>
  <SecurityGroups>
    <SecurityGroup>Awesome Users</SecurityGroup>
    <SecurityGroup>Admin Users</SecurityGroup>
  </SecurityGroups>
  <ComboGroups>
    <ComboGroup>Security and Distribution List</ComboGroup>
  </ComboGroups>
  <DistributionLists>
    <DistributionList>Awesome List</DistributionList>
  </DistributionLists>
</OPTIONS>

"@
            
            $Script:ParentFolder = Split-Path (Get-Variable MyInvocation -scope 1 -ValueOnly).MyCommand.Definition
            $XMLFile = Join-Path $ParentFolder $XMLOptions
            
            $XMLMsg = "Configuration file $XMLOptions not detected in folder $ParentFolder.  Would you like to create one now?"
            if(!(Test-Path $XMLFile)){
             if([System.Windows.Forms.MessageBox]::Show($XMLMsg,"Warning",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "Yes")
             {
                $CreateXML | Out-File $XMLFile
                $TemplateMsg = "Opening XML configuration file for editing ($XMLFile).  Please relaunch the script when the configuration is complete."
                [System.Windows.Forms.MessageBox]::Show($TemplateMsg,"Information",[System.Windows.Forms.MessageBoxButtons]::Ok) | Out-Null
                notepad $XMLFile
                Write-LogInfo -LogPath $sLogFile -Message "New XML configuration file created." #Version 1.x
                Exit
            }
            else{
                Write-LogInfo -LogPath $sLogFile -Message "Options file found." #Version 1.x
                Exit
            }
        }
        else{[XML]$Script:XML = Get-Content $XMLFile}
        if($XML.Options.Version -ne ([XML]$CreateXML).Options.Version)
        {
            $VersionMsg = "You are using an older version of the Options file.  Please generate a new Options file and transfer your settings.`r`nIn Use: $($XML.Options.Version) `r`nLatest: $(([xml]$CreateXML).Options.Version)"
            [System.Windows.Forms.MessageBox]::Show($VersionMsg,"Warning",[System.Windows.Forms.MessageBoxButtons]::Ok)
            Write-LogWarning -LogPath $sLogFile -Message "An older version of the Options file is detected." #Version 1.x
        }
        else{
            Write-LogInfo -LogPath $sLogFile -Message "Correct version of the Options file is found." #Version 1.x 
        }
        return $true #return true for success or false for failure
    }

    function OnApplicationExit {
        Remove-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010
        Remove-Module -Name ActiveDirectory
        Write-LogInfo -LogPath $sLogFile -Message "Exited succesfully." #Version 1.x
        Stop-Log -LogPath $sLogFile #Version 1.x
        $script:ExitCode = 0 #Set the exit code for the Packager
    }

    Function Set-sAMAccountName {
        Param([Switch]$Csv=$false)
        if(!$Csv)
        {
            $GivenName = $txtFirstName.text
            $SurName = $txtLastName.text
        }
        else{}
        if($XML.Options.Settings.LowerCaseUserNames.Enabled -eq "True")
        {
            $GivenName = $GivenName.ToLower()
            $SurName = $SurName.ToLower()
        }
        Switch($XML.Options.Settings.sAMAccountName.Style | Where-Object {$_.Enabled -eq $True} | Select-Object -ExpandProperty Format)
        {
            "FirstName.LastName"    {"{0}.{1}" -f $GivenName,$Surname}
            "FirstInitialLastName"  {"{0}{1}" -f ($GivenName)[0],$SurName}
            "LastNameFirstInitial"  {"{0}{1}" -f $SurName,($GivenName)[0]}
            Default                 {"{0}.{1}" -f $GivenName,$Surname}
        }
    }

    Function Set-UPN {
        Param([Switch]$Csv=$false)
        if(!$Csv)
        {
            $GivenName = $txtFirstName.text
            $SurName = $txtLastName.text
            $Domain = $cboDomain.Text
        }
        else{}
        if($XML.Options.Settings.LowerCaseUserNames.Enabled -eq "True")
        {
            $GivenName = $GivenName.ToLower()
            $SurName = $SurName.ToLower()
        }
        Switch($XML.Options.Settings.UPN.Style | Where-Object {$_.Enabled -eq $True} | Select-Object -ExpandProperty Format)
        {
            "FirstName.LastName"    {"{0}.{1}@{2}" -f $GivenName,$Surname,$Domain}
            "FirstInitialLastName"  {"{0}{1}@{2}" -f ($GivenName)[0],$SurName,$Domain}
            "LastNameFirstInitial"  {"{0}{1}@{2}" -f $SurName,($GivenName)[0],$Domain}
            Default                 {"{0}.{1}@{2}" -f $GivenName,$Surname,$Domain}
        }
    }

    Function Set-DisplayName {
        Param([Switch]$Csv=$false)
        if(!$Csv)
        {
            $GivenName = $txtFirstName.text
            $SurName = $txtLastName.text
        }
        else{}
        Switch($XML.Options.Settings.DisplayName.Style | Where-Object {$_.Enabled -eq $True} | Select-Object -ExpandProperty Format)
        {
            "FirstName LastName"    {"{0} {1}" -f $GivenName,$Surname}
            "LastName, FirstName"   {"{0}, {1}" -f $SurName, $GivenName}
            Default                 {"{0} {1}" -f $GivenName,$Surname}
        }
    }
    #endregion Application Functions
    #----------------------------------------------

    #region Email functions
    #----------------------------------------------

    Function Send-Email{
      <#
      .SYNOPSIS
      Used to send data as an email to a list of addresses
      .DESCRIPTION
      This function is used to send an email to a list of addresses. The body can be provided in HTML or plain-text
      
      .PARAMETER EmailFrom
      Mandatory. The email addresses of who you want to send the email from. Example: "admin@9to5IT.com"
      .PARAMETER EmailTo
      Mandatory. The email addresses of where to send the email to. Seperate multiple emails by ",". Example: "admin@9to5IT.com, test@test.com"
      
      .PARAMETER EmailSubject
      Mandatory. The subject of the email you want to send. Example: "Cool Script - [" + (Get-Date).ToShortDateString() + "]"
      .PARAMETER EmailBody
      Mandatory. The body of the email in plain-text or HTML format."
      .PARAMETER EmailHTML
      Mandatory. Boolean. True = email in HTML format (therefore body must be in HTML code). False = email in plain-text format"
      
      .INPUTS
      None - other than parameters above
      .OUTPUTS
      Email sent to the list of addresses specified
      .NOTES
      Version:        1.0
      Author:         Luca Sturlese
      Creation Date:  18/09/14
      Purpose/Change: Initial function development
      
      .EXAMPLE
      Send-Email -EmailFrom "admin@9to5IT.com" -EmailTo "admin@9to5IT.com, test@test.com" -EmailSubject "Cool Script - [" + (Get-Date).ToShortDateString() + "]" -EmailBody $sHTMLBody -EmailHTML $True
      .EXAMPLE
      Send-Email -EmailFrom "admin@9to5IT.com" -EmailTo "admin@9to5IT.com, test@test.com" -EmailSubject "Cool Script - [" + (Get-Date).ToShortDateString() + "]" -EmailBody "This is a test" -EmailHTML $False
      #>
      
      [CmdletBinding()]
      
      Param ([Parameter(Mandatory=$true)][string]$EmailFrom, [Parameter(Mandatory=$true)][string]$EmailTo, [Parameter(Mandatory=$true)][string]$EmailSubject, [Parameter(Mandatory=$true)][string]$EmailBody, [Parameter(Mandatory=$true)][boolean]$EmailHTML)
      
      Begin{}
      
      Process{
        Try{
          #SMTP Settings
          $sSMTPServer = $XML.Options.Settings.SMTPServer

          #Create Embedded HTML Email Message
          $oMessage = New-Object System.Net.Mail.MailMessage $EmailFrom, $EmailTo
          $oMessage.Subject = $EmailSubject
          $oMessage.IsBodyHtml = $EmailHTML
          $oMessage.Body = $EmailBody
          
          #Create SMTP object and send email
          $oSMTP = New-Object Net.Mail.SmtpClient($sSMTPServer)
          $oSMTP.Send($oMessage)

          Exit 0
      }
      
      Catch{
          Exit 1
      }
  }
  
  End{}
}

#----------------------------------------------
#endregion Email functions

#region Form Functions
#----------------------------------------------

function Call-ANUC_pff {

    #----------------------------------------------
    #region Import Form Assemblies
    #----------------------------------------------

    [void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    [void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    [void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")

    Write-LogInfo -LogPath $sLogFile -Message "Form Assemblies loaded." #Version 1.x
    #endregion Import Form Assemblies
    #----------------------------------------------
    #region Generated Form Objects
    #----------------------------------------------

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $formMain = New-Object System.Windows.Forms.Form
    $btnSubmitAll = New-Object System.Windows.Forms.Button
    $btnLast = New-Object System.Windows.Forms.Button
    $btnNext = New-Object System.Windows.Forms.Button
    $btnPrev = New-Object System.Windows.Forms.Button
    $btnFirst = New-Object System.Windows.Forms.Button
    $btnImportCSV = New-Object System.Windows.Forms.Button
    $lvCSV = New-Object System.Windows.Forms.ListView

    $cboGroup = New-Object System.Windows.Forms.ComboBox #20141120
    $lblGroup = New-Object System.Windows.Forms.Label #20141120
    $lblGroups = New-Object System.Windows.Forms.Label #20141114
    $clbGroups = New-Object System.Windows.Forms.CheckedListBox #20141114
    $lblLists = New-Object System.Windows.Forms.Label #20141114
    $clbLists = New-Object System.Windows.Forms.CheckedListBox #20141114
    $lblCombo = New-Object System.Windows.Forms.Label #20141120
    $clbCombo = New-Object System.Windows.Forms.CheckedListBox #20141120

    $txtUPN = New-Object System.Windows.Forms.TextBox
    $lblUserPrincipalName = New-Object System.Windows.Forms.Label

    $txtsAM = New-Object System.Windows.Forms.TextBox
    $lblSamAccountName = New-Object System.Windows.Forms.Label

    $txtDN = New-Object System.Windows.Forms.TextBox
    $lblDisplayName = New-Object System.Windows.Forms.Label

    $cboSite = New-Object System.Windows.Forms.ComboBox
    $lblSite = New-Object System.Windows.Forms.Label

    $cboDescription = New-Object System.Windows.Forms.ComboBox
    $lblDescription = New-Object System.Windows.Forms.Label

    $txtPassword = New-Object System.Windows.Forms.TextBox
    $lblPassword = New-Object System.Windows.Forms.Label

    $cboDomain = New-Object System.Windows.Forms.ComboBox
    $lblCurrentDomain = New-Object System.Windows.Forms.Label

    $txtOfficePhone = New-Object System.Windows.Forms.TextBox
    $lblOfficePhone = New-Object System.Windows.Forms.Label

    $txtFax = New-Object System.Windows.Forms.TextBox
    $lblFax = New-Object System.Windows.Forms.Label

    $txtMobilePhone = New-Object System.Windows.Forms.TextBox
    $lblMobilePhone = New-Object System.Windows.Forms.Label

    $txtLastName = New-Object System.Windows.Forms.TextBox
    $lblLastName = New-Object System.Windows.Forms.Label

    $cboPath = New-Object System.Windows.Forms.ComboBox
    $lblOU = New-Object System.Windows.Forms.Label

    $txtFirstName = New-Object System.Windows.Forms.TextBox
    $lblFirstName = New-Object System.Windows.Forms.Label

    $txtPostalCode = New-Object System.Windows.Forms.TextBox
    $lblPostalCode = New-Object System.Windows.Forms.Label
    
    $txtState = New-Object System.Windows.Forms.TextBox
    $lblState = New-Object System.Windows.Forms.Label
    
    $txtCity = New-Object System.Windows.Forms.TextBox
    $lblCity = New-Object System.Windows.Forms.Label
    
    $txtStreetAddress = New-Object System.Windows.Forms.TextBox
    $lblStreetAddress = New-Object System.Windows.Forms.Label
    
    $txtOffice = New-Object System.Windows.Forms.TextBox
    $lblOffice = New-Object System.Windows.Forms.Label
    
    $txtCompany = New-Object System.Windows.Forms.TextBox
    $lblCompany = New-Object System.Windows.Forms.Label
    
    $cboDepartment = New-Object System.Windows.Forms.ComboBox
    $lblDepartment = New-Object System.Windows.Forms.Label

    $cboTitle = New-Object System.Windows.Forms.ComboBox
    $lblTitle = New-Object System.Windows.Forms.Label

    $SB = New-Object System.Windows.Forms.StatusBar
    $btnSubmit = New-Object System.Windows.Forms.Button
    $System_Windows_Forms_MenuStrip_1 = New-Object System.Windows.Forms.MenuStrip
    $menustrip1 = New-Object System.Windows.Forms.MenuStrip
    $fileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $formMode = New-Object System.Windows.Forms.ToolStripMenuItem
    $CSVTemplate = New-Object System.Windows.Forms.SaveFileDialog
    $OFDImportCSV = New-Object System.Windows.Forms.OpenFileDialog
    $CreateCSVTemplate = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuExit = New-Object System.Windows.Forms.ToolStripMenuItem
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    #endregion Generated Form Objects
    #----------------------------------------------
    #region Generate form action functions
    #----------------------------------------------
    
    $formMain_Load={
        
        $formMain.Text = $formMain.Text + " " + $XML.Options.Version + " (" + $env:UserDomainName + "\" + $env:username + " on " + $env:ComputerName + ")"
        
        Write-Verbose "Adding domains to combo box"
        $XML.Options.Domains.Domain | ForEach-Object{$cboDomain.Items.Add($_.Name)}
        
        Write-Verbose "Adding OUs to combo box"
        $XML.Options.Domains.Domain | Where-Object{$_.Name -match $cboDomain.Text} | Select-Object -ExpandProperty Path | ForEach-Object{$cboPath.Items.Add($_)}
        
        Write-Verbose "Adding descriptions to combo box"
        $XML.Options.Descriptions.Description | ForEach-Object{$cboDescription.Items.Add($_)}
        
        Write-Verbose "Adding titles to combo box"
        $XML.Options.JobTitles.JobTitle | ForEach-Object{$cboTitle.Items.Add($_)}
        
        Write-Verbose "Adding sites to combo box"
        $XML.Options.Locations.Location | ForEach-Object{$cboSite.Items.Add($_.Site)}
        
        Write-Verbose "Adding departments to combo box"
        $XML.Options.Departments.Department | ForEach-Object{$cboDepartment.Items.Add($_)}
        
        Write-Verbose "Adding groups to combo box"
        $XML.Options.Groups.Group | ForEach-Object{$cboGroup.Items.Add($_.Name)} #20141120

        Write-Verbose "Adding groups to checked list box"
        $XML.Options.SecurityGroups.SecurityGroup | ForEach-Object{$clbGroups.Items.Add($_)} #20141114
        
        Write-Verbose "Adding lists to checked list box"
        $XML.Options.DistributionLists.DistributionList | ForEach-Object{$clbLists.Items.Add($_)} #20141114
        
        Write-Verbose "Adding combo to checked list box"
        $XML.Options.ComboGroups.ComboGroup | ForEach-Object{$clbCombo.Items.Add($_)} #20141120
        
        Write-Verbose "Setting default fields"
        $cboDomain.SelectedItem = $XML.Options.Default.Domain
        $cboPath.SelectedItem = $XML.Options.Default.Path
        $txtFirstName.Text = $XML.Options.Default.FirstName
        $txtLastName.Text = $XML.Options.Default.LastName
        $txtOffice.Text = $XML.Options.Default.Office
        $cboTitle.SelectedItem = $XML.Options.Default.Title
        $cboDescription.SelectedItem = $XML.Options.Default.Description
        $cboDepartment.SelectedItem = $XML.Options.Default.Department
        $txtCompany.Text = $XML.Options.Default.Company
        $cboSite.SelectedItem = $XML.Options.Default.Site
        #$txtStreetAddress.Text = $XML.Options.Default.StreetAddress
        #$txtCity.Text = $XML.Options.Default.City
        #$txtState.Text = $XML.Options.Default.State
        #$txtPostalCode.Text = $XML.Options.Default.PostalCode
        #$txtOfficePhone.Text = $XML.Options.Default.Phone
        #$txtFax.Text = $XML.Options.Default.Fax
        $txtPassword.Text = $XML.Options.Default.Password
        $cboGroup.SelectedItem = $XML.Options.Default.Group #20141120

        Write-Verbose "Creating CSV Headers"
        $Headers = @('ID','Domain','Path','FirstName','LastName','Office','Title','Description','Department','Company','Phone','Fax','Mobile','StreetAddress','City','State','PostalCode','Password','sAMAccountName','userPrincipalName','DisplayName')
        $Headers| ForEach-Object{[Void]$lvCSV.Columns.Add($_)}
        
        Write-LogInfo -LogPath $sLogFile -Message "Form created." #Version 1.x
    }
    
    $btnSubmit_Click={
        
        # Load properties from the form
        $Domain=$cboDomain.Text
        $Path=$cboPath.Text
        $GivenName = $txtFirstName.Text
        $Surname = $txtLastName.Text
        $OfficePhone = $txtOfficePhone.Text
        $Fax = $txtFax.Text
        $MobilePhone = $txtMobilePhone.Text
        $Description = $cboDescription.Text
        $Title = $cboTitle.Text
        $Department = $cboDepartment.Text
        $Company = $txtCompany.Text
        $Office = $txtOffice.Text
        $StreetAddress = $txtStreetAddress.Text
        $City = $txtCity.Text
        $State = $txtState.Text
        $PostalCode = $txtPostalCode.Text
        $Country = $XML.Options.Default.Country #20141120
        $AccountPassword = $txtPassword.text | ConvertTo-SecureString -AsPlainText -Force
        $UserGroups = $clbGroups.CheckedItems #20141114
        $UserLists = $clbLists.CheckedItems #20141114
        $UserCombo = $clbCombo.CheckedItems #20141120
        
        # Load default read-only properties from the XML
        if($XML.Options.Settings.Password.ChangeAtLogon -eq "True"){$ChangePasswordAtLogon = $True}
        else{$ChangePasswordAtLogon = $false}
        
        if($XML.Options.Settings.AccountStatus.Enabled -eq "True"){$Enabled = $True}
        else{$Enabled = $false}
        
        $Name="$GivenName $Surname"
        
        if($XML.Options.Settings.sAMAccountName.Generate -eq $True){$sAMAccountName = Set-sAMAccountName}
        else{$sAMAccountName = $txtsAM.Text}

        if($XML.Options.Settings.uPN.Generate -eq $True){$userPrincipalName = Set-UPN}
        else{$userPrincipalName = $txtuPN.Text}
        
        if($XML.Options.Settings.DisplayName.Generate -eq $True){$DisplayName = Set-DisplayName}
        else{$DisplayName = $txtDN.Text}

        $DomainController = $XML.Options.Settings.DomainController #20141117
        $DomainNS = $XML.Options.Settings.DomainNS
        $HomePage = $XML.Options.Settings.HomePage
        $ScriptPath = $XML.Options.Settings.ScriptPath
        $UserDirectory = $XML.Options.Settings.UserDirectory
        $HomeDirectory = $UserDirectory+$samAccountName
        $HomeDrive = $XML.Options.Settings.HomeDrive
        $Subfolders = $XML.Options.Settings.Subfolders
        $Subfolder = $XML.Options.Settings.Subfolders.Subfolder
        
        $User = @{
          Name = $Name
          GivenName = $GivenName
          Surname = $Surname
          Path = $Path
          samAccountName = $samAccountName
          userPrincipalName = $userPrincipalName
          DisplayName = $DisplayName
          AccountPassword = $AccountPassword
          ChangePasswordAtLogon = $ChangePasswordAtLogon
          Enabled = $Enabled
          OfficePhone = $OfficePhone
          Fax = $Fax
          Mobile = $MobilePhone
          Description = $Description
          Title = $Title
          Department = $Department
          Company = $Company
          Office = $Office
          StreetAddress = $StreetAddress
          City = $City
          State = $State
          PostalCode = $PostalCode
          Country = $Country
          HomePage = $HomePage
          ScriptPath = $ScriptPath
          HomeDirectory = $HomeDirectory
          HomeDrive = $HomeDrive
      }

      #create new user account
      $SB.Text = "Creating new user $sAMAccountName"
      $ADError = $Null
      New-ADUser @User -ErrorVariable ADError
      if ($ADerror){
        $SB.Text = "[$sAMAccountName] $ADError"
        Write-LogError -LogPath $sLogFile -Message $sAMAccountName , $ADError -ExitGracefully $False #Version 1.x
    }
    else{
        $SB.Text = "$sAMAccountName created successfully."
        Write-LogInfo -LogPath $sLogFile -Message "User [$sAMAccountName] created by $env:UserDomainName \ $env:username on $env:ComputerName" #Version 1.x
    }

    #create user folder and set permissions
    $SB.Text = "Creating user folder and setting permissions"
    New-Item -path $UserDirectory -name $sAMAccountName -type directory
    Write-LogInfo -LogPath $sLogFile -Message "User folder created." #Version 1.x
    ForEach($Subfolder in $Subfolders.Subfolder){       #Version 1.x
        New-Item -path $HomeDirectory -name $Subfolder -type directory
        Write-LogInfo -LogPath $sLogFile -Message "User subfolder ($Subfolder) created." #Version 1.x
    }

    $ACL = Get-Acl $HomeDirectory
    $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($userPrincipalName, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
    $ACL.AddAccessRule($Ar)
    Set-Acl $HomeDirectory $Acl
    Write-LogInfo -LogPath $sLogFile -Message "User folder permissions set." #Version 1.x

    #Add user to Security Groups
    $SB.Text = "Add user to Security Groups"
    $UserGroups | Add-ADGroupMember –Member $sAMAccountName #20141114

    #Add user to Distribution Lists
    $SB.Text = "Add user to Distribution Groups"
    $UserLists | Add-DistributionGroupMember -Member $sAMAccountName -BypassSecurityGroupManagerCheck #20141114
    
    #Add user to Combo Groups
    $SB.Text = "Add user to Combo Groups"
    $UserCombo | Add-ADGroupMember –Member $sAMAccountName #20141120
    
    #Mail Enable new user
    $SB.Text = "Mail Enable new user"
    Enable-Mailbox $userPrincipalName -DomainController $DomainController #20141117
    $SB.Text = "Done."
    Write-LogInfo -LogPath $sLogFile -Message "Mailbox enabled." #Version 1.x

    #Send mail to the supervisor #Version 1.x
    $sMailFrom = $XML.Options.Settings.Email.Administrator
    $sEmailTo = $XML.Options.Settings.Email.Supervisor
    $sEmailSubject = $XML.Options.Settings.Email.SubjectToSupervisor
    $sEmailBody = (Get-Content mailToSupervisor.txt)
    Send-Email -EmailFrom $sMailFrom -EmailTo $sEmailTo -EmailSubject $sEmailSubject -EmailBody $sEmailBody -EmailHTML $XML.Options.Settings.Email.EnableHtml

    #Send mail to the user #Version 1.x
    $sMailTo = $sAMAccountName + "@" + $Domain
    $sEmailSubject = $XML.Options.Settings.Email.SubjectToUser
    $sEmailBody = (Get-Content mailToUser.txt)
    Send-Email -EmailFrom $sMailFrom -EmailTo $sMailTo -EmailSubject $sEmailSubject -EmailBody $sEmailBody -EmailHTML $XML.Options.Settings.Email.EnableHtml

}

$cboDomain_SelectedIndexChanged={
    $cboPath.Items.Clear()
    Write-Verbose "Adding OUs to combo box"
    $XML.Options.Domains.Domain | Where-Object{$_.Name -match $cboDomain.Text} | Select-Object -ExpandProperty Path | ForEach-Object{$cboPath.Items.Add($_)}   
    Write-Verbose "Creating required account fields"
    
    if ($XML.Options.Settings.DisplayName.Generate) {$txtDN.Text = Set-DisplayName}
    if ($XML.Options.Settings.sAMAccountName.Generate) {$txtsAM.Text = Set-sAMAccountName}
    if ($XML.Options.Settings.UPN.Generate) {$txtUPN.Text = Set-UPN}
}

$cboSite_SelectedIndexChanged={
    Write-Verbose "Updating site fields with address information"
    $Site = $XML.Options.Locations.Location | Where-Object{$_.Site -match $cboSite.Text}
    $txtStreetAddress.Text = $Site.StreetAddress
    $txtCity.Text = $Site.City
    $txtState.Text = $Site.State
    $txtPostalCode.Text = $Site.PostalCode
    $txtOfficePhone.Text = $Site.Phone #20141120
    $txtFax.Text = $Site.Fax #20141120
}

$cboGroup_SelectedIndexChanged={ #20141120
    Write-Verbose "Updating groups fields with list information"
    $Group = @($XML.Options.Groups.Group | Where-Object {$_.Name -match $cboGroup.Text}) #20141120
    $arrayGroups = @($Group | ForEach-Object { $_.List } | Where-Object { $_.Type -match "SecurityGroup" } | ForEach-Object { $_.'#text' } ) #20141120
    #$arrayGroups = @($GroupLists | ForEach-Object { $_.'#text' } ) #20141120
    for ($i = 0; $i -lt $clbGroups.Items.Count; $i++) { if($arrayGroups -Contains $clbGroups.Items[$i]){ $clbGroups.SetItemChecked( $i, $true ) } else { $clbGroups.SetItemChecked( $i, $false ) } } #20141114
    $arrayLists = @($Group | ForEach-Object { $_.List } | Where-Object { $_.Type -match "DistributionList" } | ForEach-Object { $_.'#text' } ) #20141120
    for ($i = 0; $i -lt $clbLists.Items.Count; $i++) { if($arrayLists -Contains $clbLists.Items[$i]) { $clbLists.SetItemChecked( $i, $true ) } else { $clbLists.SetItemChecked( $i, $false ) } } #20141114
    $arrayCombo = @($Group | ForEach-Object { $_.List } | Where-Object { $_.Type -match "ComboGroup" } | ForEach-Object { $_.'#text' } ) #20141120
    for ($i = 0; $i -lt $clbCombo.Items.Count; $i++) { if($arrayCombo -Contains $clbCombo.Items[$i]) { $clbCombo.SetItemChecked( $i, $true ) } else { $clbCombo.SetItemChecked( $i, $false ) } } #20141120
}

$txtName_TextChanged={
    Write-Verbose "Creating required account fields"
    
    if ($XML.Options.Settings.DisplayName.Generate -eq $True) {$txtDN.Text = Set-DisplayName}
    if ($XML.Options.Settings.sAMAccountName.Generate -eq $True) {$txtsAM.Text = (Set-sAMAccountName)}
    if ($XML.Options.Settings.UPN.Generate -eq $True) {$txtUPN.Text = Set-UPN}
}

$createTemplateToolStripMenuItem_Click={
    $CSVTemplate.ShowDialog()
}

$CSVTemplate_FileOk=[System.ComponentModel.CancelEventHandler]{
    "" |
    Select-Object Domain,Path,FirstName,LastName,Office,Title,Description,Department,Company,Phone,StreetAddress,City,State,PostalCode,Password,sAMAccountName,userPrincipalName,DisplayName |
    Export-CSV $CSVTemplate.FileName -NoTypeInformation 
}

$formMode_Click={
    if($formMode.Text -eq 'CSV Mode'){
        $formMode.Text = "Single-User Mode"
        Get-Variable | Where-Object{$_.Name -match "txt"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left'}catch{}}
        Get-Variable | Where-Object{$_.Name -match "cbo"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left'}catch{}}
        Get-Variable | Where-Object{$_.Name -match "btn"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left'}catch{}}
        $formMain.Size = '1724,670'
        $formMain.FormBorderStyle = 'Fixed3D'
        $formMain.MaximizeBox = $False
        $formMain.MinimizeBox = $False
        $btnFirst.Visible = $True
        $btnPrev.Visible = $True
        $btnNext.Visible = $True
        $btnLast.Visible = $True
        $btnImportCSV.Visible = $True
        $btnSubmitAll.Visible = $True
        $lvCSV.Visible = $True
        $cboDomain.Width = '175'
        $cboPath.Width = '249'
        $txtFirstName.Width = '175'
        $txtLastName.Width = '175'
        $txtOffice.Width = '175'
        $cboTitle.Width = '175'
        $cboDescription.Width = '175'
        $cboDepartment.Width = '175'
        $txtCompany.Width = '175'
        $txtOfficePhone.Width = '175'
        $txtMobilePhone.Width = '175'
        $txtFax.Width = '175'
        $cboSite.Width = '175'
        $cboGroup.Width = '100'
        $txtStreetAddress.Width = '175'
        $txtCity.Width = '175'
        $txtState.Width = '175'
        $txtPostalCode.Width = '175'
        $txtPassword.Width = '175'
        $txtDN.Width = '175'
        $txtsAM.Width = '175'
        $txtUPN.Width = '175'
    }
    else{
        $formMode.Text = "CSV Mode"
        $formMain.Size = '560,670'
        $formMain.FormBorderStyle = 'Fixed3D'
        $formMain.MaximizeBox = $False
        $formMain.MinimizeBox = $False
        Get-Variable | Where-Object{$_.Name -match "txt"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left,Right'}catch{}}
        Get-Variable | Where-Object{$_.Name -match "cbo"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left,Right'}catch{}}
        Get-Variable | Where-Object{$_.Name -match "btn"} | ForEach-Object{Try{$_.Value.Anchor = 'Top,Left,Right'}catch{}}
        $btnFirst.Visible = $False
        $btnPrev.Visible = $False
        $btnNext.Visible = $False
        $btnLast.Visible = $False
        $btnImportCSV.Visible = $False
        $btnSubmitAll.Visible = $False
        $lvCSV.Visible = $False
    }
}

$btnImportCSV_Click={
    $OFDImportCSV.ShowDialog()
    $CSV = Import-Csv $OFDImportCSV.FileName
    $i = 0
    ForEach ($Entry in $CSV){
        $User = New-Object System.Windows.Forms.ListViewItem($i)
        ForEach ($Col in ($lvCSV.Columns | Where-Object{$_.Text -ne "ID"})){
            $Field = $Col.Text
            $SubItem = "$($Entry.$Field)"
            if($Field -eq 'FirstName'){$Script:GivenName = $SubItem}
            if($Field -eq 'LastName'){$Script:Surname = $SubItem}
            if($Field -eq 'Domain'){$Domain = $SubItem}
            if($Field -eq 'sAMAccountName' -AND $SubItem -eq ""){$SubItem = Set-sAMAccountName -Csv}
            if($Field -eq 'userPrincipalName' -AND $SubItem -eq ""){$SubItem = Set-UPN -Csv}
            if($Field -eq 'DisplayName' -AND $SubItem -eq ""){$SubItem = Set-DisplayName -Csv}
            $User.SubItems.Add($SubItem)
        }
        $lvCSV.Items.Add($User)
        $i++
    }
}

$lvCSV_SelectedIndexChanged={
    try{$cboDomain.SelectedItem = $lvCSV.SelectedItems[0].SubItems[1].Text}catch{}
    try{$cboPath.SelectedItem = $lvCSV.SelectedItems[0].SubItems[2].Text}catch{}
    try{$txtFirstName.Text = $lvCSV.SelectedItems[0].SubItems[3].Text}catch{}
    try{$txtLastName.Text = $lvCSV.SelectedItems[0].SubItems[4].Text}catch{}
    try{$txtOffice.Text = $lvCSV.SelectedItems[0].SubItems[5].Text}catch{}
    try{$cboTitle.SelectedItem = $lvCSV.SelectedItems[0].SubItems[6].Text}catch{}
    try{$cboDescription.SelectedItem = $lvCSV.SelectedItems[0].SubItems[7].Text}catch{}
    try{$cboDepartment.SelectedItem = $lvCSV.SelectedItems[0].SubItems[8].Text}catch{}
    try{$txtCompany.Text = $lvCSV.SelectedItems[0].SubItems[9].Text}catch{}
    try{$txtOfficePhone.Text = $lvCSV.SelectedItems[0].SubItems[10].Text}catch{}
    try{$txtFax.Text = $lvCSV.SelectedItems[0].SubItems[21].Text}catch{}
    try{$txtMobilePhone.Text = $lvCSV.SelectedItems[0].SubItems[20].Text}catch{}
    try{$txtStreetAddress.Text = $lvCSV.SelectedItems[0].SubItems[11].Text}catch{}
    try{$txtCity.Text = $lvCSV.SelectedItems[0].SubItems[12].Text}catch{}
    try{$txtState.Text = $lvCSV.SelectedItems[0].SubItems[13].Text}catch{}
    try{$txtPostalCode.Text = $lvCSV.SelectedItems[0].SubItems[14].Text}catch{}
    try{$txtPassword.Text = $lvCSV.SelectedItems[0].SubItems[15].Text}catch{}
    try{$txtsAM.Text = $lvCSV.SelectedItems[0].SubItems[16].Text}catch{}
    try{$txtuPN.Text = $lvCSV.SelectedItems[0].SubItems[17].Text}catch{}
    try{$txtDN.Text = $lvCSV.SelectedItems[0].SubItems[18].Text}catch{}
}

$btnFirst_Click={
    $lvCSV.Items | ForEach-Object{$_.Selected = $False}
    $lvCSV.Items[0].Selected = $True
}

$btnLast_Click={
    $LastRow = ($lvCSV.Items).Count - 1
    $lvCSV.Items | ForEach-Object{$_.Selected = $False}
    $lvCSV.Items[$LastRow].Selected = $True
}

$btnNext_Click={
    $LastRow = ($lvCSV.Items).Count - 1
    [Int]$Index = $lvCSV.SelectedItems[0].Index
    if($LastRow -gt $Index){
        $lvCSV.Items | ForEach-Object{$_.Selected = $False}
        $lvCSV.Items[$Index+1].Selected = $True
    }
}

$btnPrev_Click={
    [Int]$Index = $lvCSV.SelectedItems[0].Index
    if($Index -gt 0){
        $lvCSV.Items | ForEach-Object{$_.Selected = $False}
        $lvCSV.Items[$Index-1].Selected = $True
    }
}

$MenuExit_Click={
    $formMain.Close()
}

$btnSubmitAll_Click={
    $lvCSV.Items | ForEach-Object{
        
        $Domain = $_.Subitems[1].Text
        $Path = $_.Subitems[2].Text
        $GivenName = $_.Subitems[3].Text
        $Surname = $_.Subitems[4].Text
        $Office = $_.Subitems[5].Text
        $Title = $_.Subitems[6].Text
        $Description = $_.Subitems[7].Text
        $Department = $_.Subitems[8].Text
        $Company = $_.Subitems[9].Text
        $OfficePhone = $_.Subitems[10].Text
        $StreetAddress = $_.Subitems[11].Text
        $City = $_.Subitems[12].Text
        $State = $_.Subitems[13].Text
        $PostalCode = $_.Subitems[14].Text
        $MobilePhone = $_.Subitems[20].Text
        $Fax = $_.Subitems[21].Text
        
        $Name = "$GivenName $Surname"
        $nameList += "$Name ($sAMAccount)'n"

        if($XML.Options.Settings.Password.ChangeAtLogon -eq "True"){$ChangePasswordAtLogon = $True}
        else{$ChangePasswordAtLogon = $false}
        
        if($XML.Options.Settings.AccountStatus.Enabled -eq "True"){$Enabled = $True}
        else{$Enabled = $false}
        
        if($_.Subitems[16].Text -eq $null){$sAMAccountName = Set-sAMAccountName}
        else{$sAMAccountName = $_.Subitems[16].Text}

        if($_.Subitems[17].Text -eq $null){$userPrincipalName = Set-UPN}
        else{$userPrincipalName = $_.Subitems[17].Text}
        
        if($_.Subitems[18].Text -eq $null){$DisplayName = Set-DisplayName}
        else{$DisplayName = $_.Subitems[18].Text}

        $AccountPassword = $_.Subitems[15].Text | ConvertTo-SecureString -AsPlainText -Force
        $DomainController = $XML.Options.Settings.DomainController #20141117
        $DomainNS = $XML.Options.Settings.DomainNS
        $HomePage = $XML.Options.Settings.HomePage
        $ScriptPath = $XML.Options.Settings.ScriptPath
        $UserDirectory = $XML.Options.Settings.UserDirectory
        $HomeDirectory = $UserDirectory+$samAccountName
        $HomeDrive = $XML.Options.Settings.HomeDrive
        $Country = $XML.Options.Default.Country
        $Subfolders = $XML.Options.Settings.Subfolders
        $Subfolder = $XML.Options.Settings.Subfolders.Subfolder

        $User = @{
            Name = $Name
            GivenName = $GivenName
            Surname = $Surname
            Path = $Path
            samAccountName = $samAccountName
            userPrincipalName = $userPrincipalName
            DisplayName = $DisplayName
            AccountPassword = $AccountPassword
            ChangePasswordAtLogon = $ChangePasswordAtLogon
            Enabled = $Enabled
            OfficePhone = $OfficePhone
            Fax = $Fax
            Mobile = $MobilePhone
            Description = $Description
            Title = $Title
            Department = $Department
            Company = $Company
            Office = $Office
            StreetAddress = $StreetAddress
            City = $City
            State = $State
            PostalCode = $PostalCode
            Country = $Country
            HomePage = $HomePage
            ScriptPath = $ScriptPath
            HomeDirectory = $HomeDirectory
            HomeDrive = $HomeDrive
        }
        #create new user account
        $SB.Text = "Creating new user $sAMAccountName"
        $ADError = $Null
        New-ADUser @User -ErrorVariable ADError
        if ($ADerror)                {
            $SB.Text = "[$sAMAccountName] $ADError"
            Write-LogError -LogPath $sLogFile -Message $sAMAccountName , $ADError -ExitGracefully $False #Version 1.x
        }
        else{
            $SB.Text = "$sAMAccountName created successfully."
            Write-LogInfo -LogPath $sLogFile -Message "User [$sAMAccountName] created by $env:UserDomainName \ $env:username on $env:ComputerName" #Version 1.x
        }
        
        #create user folder and set permissions
        $SB.Text = "Creating user folder and setting permissions"
        New-Item -path $UserDirectory -name $sAMAccountName -type directory
        Write-LogInfo -LogPath $sLogFile -Message "User folder created." #Version 1.x
        ForEach($Subfolder in $Subfolders.Subfolder){       #Version 1.x
            New-Item -path $HomeDirectory -name $Subfolder -type directory
            Write-LogInfo -LogPath $sLogFile -Message "User subfolder ($Subfolder) created." #Version 1.x
        }
        $DomainUser = $DomainNS + '\' + $sAMAccountName
        $ACL = Get-acl $HomeDirectory
        $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($userPrincipalName, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
        $Acl.SetAccessRule($Ar)
        Set-Acl $HomeDirectory $Acl
        Write-LogInfo -LogPath $sLogFile -Message "User folder permissions set." #Version 1.x
        
        #Mail Enable new user
        $SB.Text = "Mail Enable new user"
        Enable-Mailbox $userPrincipalName -DomainController $DomainController #20141117
        $SB.Text = "Done."
        Write-LogInfo -LogPath $sLogFile -Message "Mailbox enabled." #Version 1.x

        #Send mail to the user
        $sMailBodyForUser = "Hi userPrincipalName, 'nWelcome to the <corporate name>. 'nYou can find help files on your home page."
        $sMailTo = $sAMAccountName + "@" + $Domain
        Send-Email -EmailFrom $sMailFrom -EmailTo $sMailTo -EmailSubject "Welcome $userPrincipalName" -EmailBody $sMailBodyForUser -EmailHTML $False
    }

    #Send mail to the supervisor, one for all users in the CSV
    $sHTMLBodyForAdmin = "Hi Supervisor, 'nNew users created through ANUC today. 'nCreated Accounts: 'n$nameList 'nCreator Admin: $env:username"
    $sMailFrom = $env:username + "@" + $Domain
    Send-Email -EmailFrom $sMailFrom -EmailTo "supervisor@anuc.com" -EmailSubject "New Active Directory user created" -EmailBody $sHTMLBodyForAdmin -EmailHTML $False

    
}
#endregion Generate form action functions
#----------------------------------------------
#region Generated Events
#----------------------------------------------

$Form_StateCorrection_Load=
{
    #Correct the initial state of the form to prevent the .Net maximized form issue
    $formMain.WindowState = $InitialFormWindowState
}

$Form_Cleanup_FormClosed=
{
    #Remove all event handlers from the controls
    try
    {
        $btnSubmitAll.remove_Click($btnSubmitAll_Click)
        $btnLast.remove_Click($btnLast_Click)
        $btnNext.remove_Click($btnNext_Click)
        $btnPrev.remove_Click($btnPrev_Click)
        $btnFirst.remove_Click($btnFirst_Click)
        $btnImportCSV.remove_Click($btnImportCSV_Click)
        $lvCSV.remove_SelectedIndexChanged($lvCSV_SelectedIndexChanged)
        $cboSite.remove_SelectedIndexChanged($cboSite_SelectedIndexChanged)
        $cboGroup.remove_SelectedIndexChanged($cboGroup_SelectedIndexChanged) #20141120
        $cboDomain.remove_SelectedIndexChanged($cboDomain_SelectedIndexChanged)
        $txtLastName.remove_TextChanged($txtName_TextChanged)
        $txtFirstName.remove_TextChanged($txtName_TextChanged)
        $btnSubmit.remove_Click($btnSubmit_Click)
        $formMain.remove_Load($formMain_Load)
        $formMode.remove_Click($formMode_Click)
        $CSVTemplate.remove_FileOk($CSVTemplate_FileOk)
        $CreateCSVTemplate.remove_Click($createTemplateToolStripMenuItem_Click)
        $MenuExit.remove_Click($MenuExit_Click)
        $formMain.remove_Load($Form_StateCorrection_Load)
        $formMain.remove_FormClosed($Form_Cleanup_FormClosed)
    }
    catch [Exception]
    { }
}
#endregion Generated Events
#----------------------------------------------
#region Generated Form Code
#----------------------------------------------

#
# formMain
#
$formMain.Controls.Add($btnSubmitAll)
$formMain.Controls.Add($btnLast)
$formMain.Controls.Add($btnNext)
$formMain.Controls.Add($btnPrev)
$formMain.Controls.Add($btnFirst)
$formMain.Controls.Add($btnImportCSV)
$formMain.Controls.Add($lvCSV)
$formMain.Controls.Add($cboGroup) #20141120
$formMain.Controls.Add($lblGroup) #20141120
$formMain.Controls.Add($lblGroups) #20141114
$formMain.Controls.Add($clbGroups) #20141114
$formMain.Controls.Add($lblLists) #20141114
$formMain.Controls.Add($clbLists) #20141114
$formMain.Controls.Add($lblCombo) #20141120
$formMain.Controls.Add($clbCombo) #20141120
$formMain.Controls.Add($txtUPN)
$formMain.Controls.Add($txtsAM)
$formMain.Controls.Add($txtDN)
$formMain.Controls.Add($cboDepartment)
$formMain.Controls.Add($lblUserPrincipalName)
$formMain.Controls.Add($lblSamAccountName)
$formMain.Controls.Add($lblDisplayName)
$formMain.Controls.Add($SB)
$formMain.Controls.Add($cboSite)
$formMain.Controls.Add($lblSite)
$formMain.Controls.Add($cboDescription)
$formMain.Controls.Add($txtPassword)
$formMain.Controls.Add($lblPassword)
$formMain.Controls.Add($cboDomain)
$formMain.Controls.Add($lblCurrentDomain)
$formMain.Controls.Add($txtPostalCode)
$formMain.Controls.Add($txtState)
$formMain.Controls.Add($txtCity)
$formMain.Controls.Add($txtStreetAddress)
$formMain.Controls.Add($txtOffice)
$formMain.Controls.Add($txtCompany)
$formMain.Controls.Add($cboTitle)
$formMain.Controls.Add($txtOfficePhone)
$formMain.Controls.Add($txtFax)
$formMain.Controls.Add($txtMobilePhone)
$formMain.Controls.Add($txtLastName)
$formMain.Controls.Add($cboPath)
$formMain.Controls.Add($lblOU)
$formMain.Controls.Add($txtFirstName)
$formMain.Controls.Add($lblPostalCode)
$formMain.Controls.Add($lblState)
$formMain.Controls.Add($lblCity)
$formMain.Controls.Add($lblStreetAddress)
$formMain.Controls.Add($lblOffice)
$formMain.Controls.Add($lblCompany)
$formMain.Controls.Add($lblDepartment)
$formMain.Controls.Add($lblTitle)
$formMain.Controls.Add($btnSubmit)
$formMain.Controls.Add($lblDescription)
$formMain.Controls.Add($lblOfficePhone)
$formMain.Controls.Add($lblFax)
$formMain.Controls.Add($lblMobilePhone)
$formMain.Controls.Add($lblLastName)
$formMain.Controls.Add($lblFirstName)
$formMain.Controls.Add($menustrip1)

$formMain.AcceptButton = $btnSubmit
$formMain.FormBorderStyle = 'Fixed3D'
$formMain.ClientSize = '544, 635' #subtract 16,35 pts for borders
$formMain.MaximizeBox = $False
$formMain.MinimizeBox = $False
$formMain.MainMenuStrip = $System_Windows_Forms_MenuStrip_1
$formMain.Name = "formMain"
$formMain.ShowIcon = $False
$formMain.StartPosition = 'CenterScreen'
$formMain.Text = $XML.Options.Product #20141117
$formMain.add_Load($formMain_Load)

$System_Windows_Forms_MenuStrip_1.Location = '0, 0'
$System_Windows_Forms_MenuStrip_1.Name = ""
$System_Windows_Forms_MenuStrip_1.Size = '271, 24'
$System_Windows_Forms_MenuStrip_1.TabIndex = 1
$System_Windows_Forms_MenuStrip_1.Visible = $False
#
# lblCurrentDomain
#
$lblCurrentDomain.Location = '10, 35'
$lblCurrentDomain.Name = "lblCurrentDomain"
$lblCurrentDomain.Size = '100, 23'
$lblCurrentDomain.TabIndex = 39
$lblCurrentDomain.Text = "Current Domain"
$lblCurrentDomain.TextAlign = 'MiddleLeft'
#
# cboDomain
#
$cboDomain.Anchor = 'Top, Left, Right'
$cboDomain.FormattingEnabled = $True
$cboDomain.Location = '118, 35'
$cboDomain.Name = "cboDomain"
$cboDomain.Size = '173, 21'
$cboDomain.TabIndex = 1
$cboDomain.add_SelectedIndexChanged($cboDomain_SelectedIndexChanged)
$cboDomain.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboDomain.Enabled = $True
#
# lblOU
#
$lblOU.Location = '10, 65'
$lblOU.Name = "lblOU"
$lblOU.Size = '36, 23'
$lblOU.TabIndex = 26
$lblOU.Text = "OU"
$lblOU.TextAlign = 'MiddleLeft'
#
# cboPath
#
$cboPath.Anchor = 'Top, Left, Right'
$cboPath.FormattingEnabled = $True
$cboPath.Location = '45, 65'
$cboPath.Name = "cboPath"
$cboPath.Size = '247, 21'
$cboPath.TabIndex = 2
$cboPath.Enabled = $True
#
# lblFirstName
#
$lblFirstName.Location = '10, 110'
$lblFirstName.Name = "lblFirstName"
$lblFirstName.Size = '100, 23'
$lblFirstName.TabIndex = 12
$lblFirstName.Text = "First Name"
$lblFirstName.TextAlign = 'MiddleLeft'
#
# txtFirstName
#
$txtFirstName.Anchor = 'Top, Left, Right'
$txtFirstName.Location = '118, 110'
$txtFirstName.Name = "txtFirstName"
$txtFirstName.Size = '173, 20'
$txtFirstName.TabIndex = 3
$txtFirstName.add_TextChanged($txtName_TextChanged)
$txtFirstName.Enabled = $True
#
# lblLastName
#
$lblLastName.Location = '10, 135'
$lblLastName.Name = "lblLastName"
$lblLastName.Size = '100, 23'
$lblLastName.TabIndex = 13
$lblLastName.Text = "Last Name"
$lblLastName.TextAlign = 'MiddleLeft'
#
# txtLastName
#
$txtLastName.Anchor = 'Top, Left, Right'
$txtLastName.Location = '118, 135'
$txtLastName.Name = "txtLastName"
$txtLastName.Size = '173, 20'
$txtLastName.TabIndex = 4
$txtLastName.add_TextChanged($txtName_TextChanged)
$txtLastName.Enabled = $True
#
# lblOffice
#
$lblOffice.Location = '10, 160'
$lblOffice.Name = "lblOffice"
$lblOffice.Size = '100, 23'
$lblOffice.TabIndex = 20
$lblOffice.Text = "Office"
$lblOffice.TextAlign = 'MiddleLeft'
#
# txtOffice
#
$txtOffice.Anchor = 'Top, Left, Right'
$txtOffice.Location = '118, 160'
$txtOffice.Name = "txtOffice"
$txtOffice.Size = '173, 20'
$txtOffice.TabIndex = 5
$txtOffice.Enabled = $True
#
# lblTitle
#
$lblTitle.Location = '10, 185'
$lblTitle.Name = "lblTitle"
$lblTitle.Size = '100, 23'
$lblTitle.TabIndex = 17
$lblTitle.Text = "Title"
$lblTitle.TextAlign = 'MiddleLeft'
#
# cboTitle
#
$cboTitle.Anchor = 'Top, Left, Right'
$cboTitle.FormattingEnabled = $True
$cboTitle.Location = '118, 185'
$cboTitle.Name = "cboTitle"
$cboTitle.Size = '173, 20'
$cboTitle.TabIndex = 6
$cboTitle.Enabled = $True
#
# lblDescription
#
$lblDescription.Location = '10, 210'
$lblDescription.Name = "lblDescription"
$lblDescription.Size = '100, 23'
$lblDescription.TabIndex = 15
$lblDescription.Text = "Description"
$lblDescription.TextAlign = 'MiddleLeft'
#
# cboDescription
#
$cboDescription.Anchor = 'Top, Left, Right'
$cboDescription.FormattingEnabled = $True
$cboDescription.Location = '118, 210'
$cboDescription.Name = "cboDescription"
$cboDescription.Size = '173, 21'
$cboDescription.TabIndex = 7
$cboDescription.Enabled = $True
#
# lblDepartment
#
$lblDepartment.Location = '10, 235'
$lblDepartment.Name = "lblDepartment"
$lblDepartment.Size = '100, 23'
$lblDepartment.TabIndex = 18
$lblDepartment.Text = "Department"
$lblDepartment.TextAlign = 'MiddleLeft'
#
# cboDepartment
#
$cboDepartment.Anchor = 'Top, Left, Right'
$cboDepartment.FormattingEnabled = $True
$cboDepartment.Location = '118, 235'
$cboDepartment.Name = "cboDepartment"
$cboDepartment.Size = '173, 21'
$cboDepartment.TabIndex = 8
$cboDepartment.Enabled = $True
#
# lblCompany
#
$lblCompany.Location = '10, 260'
$lblCompany.Name = "lblCompany"
$lblCompany.Size = '100, 23'
$lblCompany.TabIndex = 19
$lblCompany.Text = "Company"
$lblCompany.TextAlign = 'MiddleLeft'
#
# txtCompany
#
$txtCompany.Anchor = 'Top, Left, Right'
$txtCompany.Location = '118, 260'
$txtCompany.Name = "txtCompany"
$txtCompany.Size = '173, 20'
$txtCompany.TabIndex = 9
$txtCompany.Enabled = $True
#
# lblMobilePhone
#
$lblMobilePhone.Location = '10, 285'
$lblMobilePhone.Name = "lblMobilePhone"
$lblMobilePhone.Size = '100, 23'
$lblMobilePhone.TabIndex = 14
$lblMobilePhone.Text = "Mobile Phone"
$lblMobilePhone.TextAlign = 'MiddleLeft'
#
# txtMobilePhone
#
$txtMobilePhone.Anchor = 'Top, Left, Right'
$txtMobilePhone.Location = '118, 285'
$txtMobilePhone.Name = "txtMobilePhone"
$txtMobilePhone.Size = '173, 20'
$txtMobilePhone.TabIndex = 10
$txtMobilePhone.Enabled = $True
#
# lblSite
#
$lblSite.Location = '10, 320'
$lblSite.Name = "lblSite"
$lblSite.Size = '100, 23'
$lblSite.TabIndex = 44
$lblSite.Text = "Site"
$lblSite.TextAlign = 'MiddleLeft'
#
# cboSite
#
$cboSite.Anchor = 'Top, Left, Right'
$cboSite.FormattingEnabled = $True
$cboSite.Location = '118, 320'
$cboSite.Name = "cboSite"
$cboSite.Size = '173, 21'
$cboSite.TabIndex = 11
$cboSite.add_SelectedIndexChanged($cboSite_SelectedIndexChanged)
$cboSite.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboSite.Enabled = $True
#
# lblOfficePhone
#
$lblOfficePhone.Location = '10, 345'
$lblOfficePhone.Name = "lblOfficePhone"
$lblOfficePhone.Size = '100, 23'
$lblOfficePhone.TabIndex = 14
$lblOfficePhone.Text = "Office Phone"
$lblOfficePhone.TextAlign = 'MiddleLeft'
#
# txtOfficePhone
#
$txtOfficePhone.Anchor = 'Top, Left, Right'
$txtOfficePhone.Location = '118, 345'
$txtOfficePhone.Name = "txtOfficePhone"
$txtOfficePhone.Size = '173, 20'
$txtOfficePhone.TabIndex = 10
$txtOfficePhone.Enabled = $True
#
# lblFax
#
$lblFax.Location = '10, 370'
$lblFax.Name = "lblFax"
$lblFax.Size = '100, 23'
$lblFax.TabIndex = 14
$lblFax.Text = "Office Fax"
$lblFax.TextAlign = 'MiddleLeft'
#
# txtFax
#
$txtFax.Anchor = 'Top, Left, Right'
$txtFax.Location = '118, 370'
$txtFax.Name = "txtFax"
$txtFax.Size = '173, 20'
$txtFax.TabIndex = 10
$txtFax.Enabled = $True
#
# lblStreetAddress
#
$lblStreetAddress.Location = '10, 395'
$lblStreetAddress.Name = "lblStreetAddress"
$lblStreetAddress.Size = '100, 23'
$lblStreetAddress.TabIndex = 21
$lblStreetAddress.Text = "Street Address"
$lblStreetAddress.TextAlign = 'MiddleLeft'
#
# txtStreetAddress
#
$txtStreetAddress.Anchor = 'Top, Left, Right'
$txtStreetAddress.Location = '118, 395'
$txtStreetAddress.Name = "txtStreetAddress"
$txtStreetAddress.Size = '173, 20'
$txtStreetAddress.TabIndex = 12
$txtStreetAddress.Enabled = $True
#
# lblCity
#
$lblCity.Location = '10, 420'
$lblCity.Name = "lblCity"
$lblCity.Size = '100, 23'
$lblCity.TabIndex = 22
$lblCity.Text = "City"
$lblCity.TextAlign = 'MiddleLeft'
#
# txtCity
#
$txtCity.Anchor = 'Top, Left, Right'
$txtCity.Location = '118, 420'
$txtCity.Name = "txtCity"
$txtCity.Size = '173, 20'
$txtCity.TabIndex = 13
$txtCity.Enabled = $True
#
# lblState
#
$lblState.Location = '10, 445'
$lblState.Name = "lblState"
$lblState.Size = '100, 23'
$lblState.TabIndex = 23
$lblState.Text = "State"
$lblState.TextAlign = 'MiddleLeft'
#
# txtState
#
$txtState.Anchor = 'Top, Left, Right'
$txtState.Location = '118, 445'
$txtState.Name = "txtState"
$txtState.Size = '173, 20'
$txtState.TabIndex = 14
$txtState.Enabled = $True
#
# lblPostalCode
#
$lblPostalCode.Location = '10, 470'
$lblPostalCode.Name = "lblPostalCode"
$lblPostalCode.Size = '100, 23'
$lblPostalCode.TabIndex = 24
$lblPostalCode.Text = "Postal Code"
$lblPostalCode.TextAlign = 'MiddleLeft'
#
# txtPostalCode
#
$txtPostalCode.Anchor = 'Top, Left, Right'
$txtPostalCode.Location = '118, 470'
$txtPostalCode.Name = "txtPostalCode"
$txtPostalCode.Size = '173, 20'
$txtPostalCode.TabIndex = 15
$txtPostalCode.Enabled = $True
#
# lblDisplayName
#
$lblDisplayName.Location = '10, 505'
$lblDisplayName.Name = "lblDisplayName"
$lblDisplayName.Size = '100, 23'
$lblDisplayName.TabIndex = 46
$lblDisplayName.Text = "Display Name"
$lblDisplayName.TextAlign = 'MiddleLeft'
#
# txtDN
#
$txtDN.Anchor = 'Top, Left, Right'
$txtDN.Location = '118, 505'
$txtDN.Name = "txtDN"
$txtDN.Size = '173, 20'
$txtDN.TabIndex = 49
#
# lblSamAccountName
#
$lblSamAccountName.Location = '10, 530'
$lblSamAccountName.Name = "lblSamAccountName"
$lblSamAccountName.Size = '100, 23'
$lblSamAccountName.TabIndex = 47
$lblSamAccountName.Text = "samAccountName"
$lblSamAccountName.TextAlign = 'MiddleLeft'
#
# txtsAM
#
$txtsAM.Anchor = 'Top, Left, Right'
$txtsAM.Location = '118, 530'
$txtsAM.Name = "txtsAM"
$txtsAM.Size = '173, 20'
$txtsAM.TabIndex = 50
#
# lblUserPrincipalName
#
$lblUserPrincipalName.Location = '10, 555'
$lblUserPrincipalName.Name = "lblUserPrincipalName"
$lblUserPrincipalName.Size = '100, 23'
$lblUserPrincipalName.TabIndex = 48
$lblUserPrincipalName.Text = "userPrincipalName"
$lblUserPrincipalName.TextAlign = 'MiddleLeft'
#
# txtUPN
#
$txtUPN.Anchor = 'Top, Left, Right'
$txtUPN.Location = '118, 555'
$txtUPN.Name = "txtUPN"
$txtUPN.Size = '173, 20'
$txtUPN.TabIndex = 51
#
# lblPassword
#
$lblPassword.Location = '10, 580'
$lblPassword.Name = "lblPassword"
$lblPassword.Size = '100, 23'
$lblPassword.TabIndex = 41
$lblPassword.Text = "Password"
$lblPassword.TextAlign = 'MiddleLeft'
#
# txtPassword
#
$txtPassword.Anchor = 'Top, Left, Right'
$txtPassword.Location = '118, 582'
$txtPassword.Name = "txtPassword"
$txtPassword.Size = '173, 20'
$txtPassword.TabIndex = 16
$txtPassword.UseSystemPasswordChar = $True
#
# lblGroup                                  #20141120
#
$lblGroup.Location = '305, 40'
$lblGroup.Name = "lblGroup"
$lblGroup.Size = '100, 23'
$lblGroup.TabIndex = 44
$lblGroup.Text = "Groups Template"
$lblGroup.TextAlign = 'MiddleLeft'
#
# cboGroup                                  #20141120
#
$cboGroup.Anchor = 'Top, Left, Right'
$cboGroup.FormattingEnabled = $True
$cboGroup.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboGroup.Location = '405, 40'
$cboGroup.Name = "cboGroup"
$cboGroup.Size = '100, 21'
$cboGroup.TabIndex = 11
$cboGroup.add_SelectedIndexChanged($cboGroup_SelectedIndexChanged)
$cboGroup.Enabled = $True
#
# lblLists                                  #20141114
#
$lblLists.Location = '305, 65'
$lblLists.Name = "lblLists"
$lblLists.Size = '100, 23'
$lblLists.Width = 210
$lblLists.Text = "Distribution Groups"
$lblLists.TextAlign = 'MiddleLeft'
#
# clbLists                                  #20141114
#
$clbLists.Location = '305, 90'
$clbLists.Name = "clbLists"
$clbLists.Size = '210, 150'
$clbLists.CheckOnClick = $true;
$clbLists.TabIndex = 17
$clbLists.Enabled = $True
#
# lblGroups                                 #20141114
#
$lblGroups.Location = '305, 245'
$lblGroups.Name = "lblGroups"
$lblGroups.Size = '100, 23'
$lblGroups.Width = 210
$lblGroups.Text = "Security Groups"
$lblGroups.TextAlign = 'MiddleLeft'
#
# clbGroups                                 #20141114
#
$clbGroups.Location = '305, 270'
$clbGroups.Name = "clbGroups"
$clbGroups.Size = '210, 150'
$clbGroups.CheckOnClick = $true;
$clbGroups.TabIndex = 18
$clbGroups.Enabled = $True
#
# lblCombo                                  #20141120
#
$lblCombo.Location = '305, 425'
$lblCombo.Name = "lblCombo"
$lblCombo.Size = '100, 23'
$lblCombo.Width = 210
$lblCombo.Text = "Combo Groups"
$lblCombo.TextAlign = 'MiddleLeft'
#
# clbCombo                                  #20141120
#
$clbCombo.Location = '305, 450'
$clbCombo.Name = "clbCombo"
$clbCombo.Size = '210, 150'
$clbCombo.CheckOnClick = $true;
$clbCombo.TabIndex = 18
$clbCombo.Enabled = $True
#
# btnSubmit
#
$btnSubmit.Location = '416, 0'
$btnSubmit.Name = "btnSubmit"
$btnSubmit.Size = '100, 25'
$btnSubmit.TabIndex = 19
$btnSubmit.Text = "Submit"
$btnSubmit.UseVisualStyleBackColor = $True
$btnSubmit.add_Click($btnSubmit_Click)
#
# SB
#
$SB.Location = '0, 610'
$SB.Name = "SB"
$SB.Size = '304, 22'
$SB.TabIndex = 45
$SB.Text = "Ready"
#
# btnSubmitAll
#
$btnSubmitAll.Location = '728, 35'
$btnSubmitAll.Name = "btnSubmitAll"
$btnSubmitAll.Size = '75, 25'
$btnSubmitAll.TabIndex = 59
$btnSubmitAll.Text = "Submit All"
$btnSubmitAll.UseVisualStyleBackColor = $True
$btnSubmitAll.Visible = $False
$btnSubmitAll.add_Click($btnSubmitAll_Click)
#
# btnLast
#
$btnLast.Location = '697, 35'
$btnLast.Name = "btnLast"
$btnLast.Size = '30, 25'
$btnLast.TabIndex = 58
$btnLast.Text = ">>"
$btnLast.UseVisualStyleBackColor = $True
$btnLast.Visible = $False
$btnLast.add_Click($btnLast_Click)
#
# btnNext
#
$btnNext.Location = '666, 35'
$btnNext.Name = "btnNext"
$btnNext.Size = '30, 25'
$btnNext.TabIndex = 57
$btnNext.Text = ">"
$btnNext.UseVisualStyleBackColor = $True
$btnNext.Visible = $False
$btnNext.add_Click($btnNext_Click)
#
# btnPrev
#
$btnPrev.Location = '635, 35'
$btnPrev.Name = "btnPrev"
$btnPrev.Size = '30, 25'
$btnPrev.TabIndex = 56
$btnPrev.Text = "<"
$btnPrev.UseVisualStyleBackColor = $True
$btnPrev.Visible = $False
$btnPrev.add_Click($btnPrev_Click)
#
# btnFirst
#
$btnFirst.Location = '604, 35'
$btnFirst.Name = "btnFirst"
$btnFirst.Size = '30, 25'
$btnFirst.TabIndex = 55
$btnFirst.Text = "<<"
$btnFirst.UseVisualStyleBackColor = $True
$btnFirst.Visible = $False
$btnFirst.add_Click($btnFirst_Click)
#
# btnImportCSV
#
$btnImportCSV.Location = '528, 35'
$btnImportCSV.Name = "btnImportCSV"
$btnImportCSV.Size = '75, 25'
$btnImportCSV.TabIndex = 54
$btnImportCSV.Text = "Import CSV"
$btnImportCSV.UseVisualStyleBackColor = $True
$btnImportCSV.Visible = $False
$btnImportCSV.add_Click($btnImportCSV_Click)
#
# lvCSV
#
$lvCSV.FullRowSelect = $True
$lvCSV.GridLines = $True
$lvCSV.Location = '530, 65'
$lvCSV.Name = "lvCSV"
$lvCSV.Size = '1150, 535'
$lvCSV.TabIndex = 53
$lvCSV.UseCompatibleStateImageBehavior = $False
$lvCSV.View = 'Details'
$lvCSV.Visible = $False
$lvCSV.add_SelectedIndexChanged($lvCSV_SelectedIndexChanged)
#
# menustrip1
#
[void]$menustrip1.Items.Add($fileToolStripMenuItem)
$menustrip1.Location = '0, 0'
$menustrip1.Name = "menustrip1"
$menustrip1.Size = '304, 24'
$menustrip1.TabIndex = 52
$menustrip1.Text = "menustrip1"
#
# fileToolStripMenuItem
#
[void]$fileToolStripMenuItem.DropDownItems.Add($formMode)
[void]$fileToolStripMenuItem.DropDownItems.Add($CreateCSVTemplate)
[void]$fileToolStripMenuItem.DropDownItems.Add($MenuExit)
$fileToolStripMenuItem.Name = "fileToolStripMenuItem"
$fileToolStripMenuItem.Size = '37, 20'
$fileToolStripMenuItem.Text = "File"
#
# formMode
#
$formMode.Name = "formMode"
$formMode.Size = '185, 22'
$formMode.Text = "CSV Mode"
$formMode.add_Click($formMode_Click)
#
# CSVTemplate
#
$CSVTemplate.CheckPathExists = $False
$CSVTemplate.DefaultExt = "csv"
$CSVTemplate.FileName = Join-Path $ParentFolder "ANUCusers.csv"
$CSVTemplate.Filter = "CSV Files|*.csv|All Files|*.*"
$CSVTemplate.ShowHelp = $True
$CSVTemplate.Title = "Create CSV Template For ANUC"
$CSVTemplate.add_FileOk($CSVTemplate_FileOk)
#
# OFDImportCSV
#
$OFDImportCSV.FileName = Join-Path $ParentFolder "ANUCusers.csv"
$OFDImportCSV.ShowHelp = $True
#
# CreateCSVTemplate
#
$CreateCSVTemplate.Name = "CreateCSVTemplate"
$CreateCSVTemplate.Size = '185, 22'
$CreateCSVTemplate.Text = "Create CSV Template"
$CreateCSVTemplate.add_Click($createTemplateToolStripMenuItem_Click)
#
# MenuExit
#
$MenuExit.Name = "MenuExit"
$MenuExit.Size = '185, 22'
$MenuExit.Text = "Exit"
$MenuExit.add_Click($MenuExit_Click)

#endregion Generated Form Code
#----------------------------------------------

#Save the initial state of the form
$InitialFormWindowState = $formMain.WindowState
#Init the OnLoad event to correct the initial state of the form
$formMain.add_Load($Form_StateCorrection_Load)
#Clean up the control events
$formMain.add_FormClosed($Form_Cleanup_FormClosed)
#Show the Form
return $formMain.ShowDialog()
Write-LogInfo -LogPath $sLogFile -Message "Form created." #Version 1.x

} #End Function
#endregion Form Functions
#----------------------------------------------

#Call OnApplicationLoad to initialize
if((OnApplicationLoad) -eq $true) {
    #Call the form
    Call-ANUC_pff | Out-Null
    #Perform cleanup
    OnApplicationExit
}

#For debugging
$VerbosePreference = $oldVerbosePreference