
###########################################################################

Import-Module D:\Powershell_Scripts\CommonModule\CommonModule.ps1

$global:htmlbody = @()
$global:errorbody = @()
$newVMSQLSummary = @()

$CSSStyle = fnGetCSSStyles

$CurrentDateTime = (Get-Date).ToString()
$ScriptPath = fnGetCurrentDirectory

$csvfilename = $ReportName + '_' +  ((($CurrentDateTime).Replace('/','')).Replace(':','')).replace(' ','_') + '.csv'
$csvPath = $scriptPath + '\Logs\' + $csvfilename

$logfilename = $ReportName + '_' +  ((($CurrentDateTime).Replace('/','')).Replace(':','')).replace(' ','_') + '.html'
$LogFile = $scriptPath + '\Logs\' + $logfilename

 ###########################################################################

# Changing Variables 
$vmName = 'SharePoint16'
$intipAddress = '192.168.10.41'
$extipAddress = '192.168.0.14'

# Constant Variables
$vhdxPath = 'D:\HyperVMachines\Templates\WindowsServer2016_New.vhdx'
$vmVhdxPath = "D:\HyperVMachines\$vmName\$vmName.vhdx"
$cred = fnRetrieveCredentials -domain none
$domainCred = fnRetrieveCredentials -domain sakharam
$vmInternetSwitch = 'NATSwitch'
$vmNetworkSwitch = 'InternalSwitch'
$subnetMask = '255.255.255.0'
$domainName = 'sakharam.com'
$domainIPAddress = '192.168.10.30'
$internetGateway = '192.168.0.1'
$vmMemory = 4
$dDrivePath = "D:\HyperVMachines\$vmName\D_Drive.vhdx"
$ReportName = "New VM Creation - $vmName"

###########################################################################

# To check if the Powershell is being executed as Administrator 
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())

if (($currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) -eq $false)  {
    Write-Host "Current user context is not as administrator. Please open the ISE with administrator priviliges." -ForegroundColor Red
    Exit
}


# To create a new virtual machine name on the host
try
{
    If(Get-VM -Name $vmName -ErrorAction SilentlyContinue)
    {
        "The virtual machine $vmName is already created. Proceeding with further configuration: " | fnWriteLog -LogType Warning -string

        <#
        Stop-VM -Name $vmName -Force
        If((Get-VM -Name $vmName).State -eq 'Running')
        {
            While ((Get-VM -Name $vmName).State -eq 'Running')
            {
                Write-Host "Waiting for the virtual machine to shut down..." -ForegroundColor Cyan
                Start-Sleep 2    
            } 
        }
        Remove-VM -Name $vmName -Force
        #>
    }
    else
    {
        New-VM -Name $vmName -MemoryStartupBytes 2GB -Path "D:\HyperVMachines"  -ErrorAction STOP
        "Virtual Machine $vmName is created successfully." | fnwriteLog -LogType Success -string
    }
}
catch
{
    "Error while creating a new VM $vmName. Error Message: $($_.Exception.Message). " | fnWriteLog -LogType Error -string
    If ($($_.Exception.Message) -like "*Logon failure: the user has not been granted the requested logon type at this computer*")
    {
        "Trying to update the group policy on the host:" | fnWriteLog -LogType Warning -string
        gpupdate /force
        "Group Policy updated successfully." | fnWriteLog -LogType Success -string
    }

    try
    {
        New-VM -Name $vmName -MemoryStartupBytes 2GB -Path D:\HyperVMachines -ErrorAction STOP
        "Virtual Machine $vmName is created successfully." | fnwriteLog -LogType Success -string
    }
    catch
    {
        "Error creating a new virtual machine $vmName. Error Message: $($_.Exception.Message)" | fnWriteLog -LogType Error -string
    }
}


# Copy the VHDX template from the source to the new VM folder
If(!(Test-Path -Path $vmVhdxPath))
{
    try
    {
        Copy -Path $vhdxPath -Destination $vmVhdxPath -Verbose -ErrorAction STOP
        "File $vhdxPath copied to $vmVhdxPath successfully" | fnwriteLog -LogType Info -string
    }
    catch
    {
        "Error: Copying the file #vhdxPath to the path $vmVhdxPath failed." | fnwriteLog -LogType Error -string
    }
}
else
{
    "The gold image vhdx copy $vmvhdxPath already exists at the specificed location." | fnwriteLog -LogType Warning -string
}


# Adding the VHDX Gold Image to the virtual machine 
If(!(Test-Path -Path $vmVhdxPath))
{
    try
    {
        Add-VMHardDiskDrive -VMName $vmName -path $vmVhdxPath 
        "Virtual hard disk is attached successfully to the virtual machie $vmName." | fnWriteLog -LogType Success -string
    }
    catch
    {
      "Adding the virtual hard disk $vmVhdxPath to the virtual machine $vmName failed. Error: $_.Exception.Message" | fnwriteLog -LogType Error -string  
    }
}
else
{
    "The virtual hard disk $vmVhdxPath is already added to the virtual machine" | fnwriteLog -LogType Warning -string  
}


# Connect the internal network adaptor to the virtual machine
If(Get-VMSwitch -Name $vmNetworkSwitch -ErrorAction STOP)
{
    if((Get-VMNetworkAdapter -VMName $vmName).SwitchName -eq $vmNetworkSwitch)
    {
        "The network switch $vmNetworkSwitch is already connected to the virtual machine $vmName successfully." | fnWriteLog -LogType Warning -string
    }
    else
    {
        try
        {
            If((Get-VM -Name $vmName).State -eq 'Running') { Stop-VM -Name $vmName -Force }
            Connect-VMNetworkAdapter -VMName $vmName -SwitchName $vmNetworkSwitch
            "The network switch $vmNetworkSwitch is connected to the virtual machine $vmName successfully." | fnWriteLog -LogType Success -string
        }
        catch
        {
            "Connecting the network switch $vmNetworkSwitch to the virtual machine $vmName failed with the error: $($_.Exception.Message)" | fnWriteLog -LogType Error -string
        }
    }
}
else
{
    "Error: Please verify if the network switch $vmNetworkSwitch exists in Hyper-V." | fnwriteLog -LogType Error -string
}

# Add the internet network adaptor to the virtual machine
If(Get-VMSwitch -Name $vmInternetSwitch -ErrorAction STOP)
{
    If((Get-VMNetworkAdapter -VMName $vmName | Where-Object {$_.SwitchName -eq $vmInternetSwitch}).SwitchName)
    {
        "The internet network switch $vmInternetSwitch is already added to the virtual machine $vmName. " | fnWriteLog -LogType Warning -string
    }
    else
    {
        try
        {
            If((Get-VM -Name $vmName).State -eq 'Running') { Stop-VM -Name $vmName -Force }
            Add-VMNetworkAdapter -VMName $vmName -SwitchName $vmInternetSwitch
            "The internet network switch $vmInternetSwitch is added to the virtual machine $vmName successfully. " | fnWriteLog -LogType Success -string
        }
        catch
        {
            "Connecting the network switch $vmInternetSwitch to the virtual machine $vmName failed with the error: $($_.Exception.Message)" | fnWriteLog -LogType Error -string
        }
    }
}
else
{
    "Error: Please verify if the network switch $vmInternetSwitch exists in Hyper-V." | fnwriteLog -LogType Error -string
}
#Set the VM memory
If((Get-VMMemory -VMName $vmName | Select Startup).StartUp/1GB -ne $vmMemory)
{
    $dynamicMemoryEnabledFlag = (Get-VMMemory -VMName $vmName | Select DynamicMemoryEnabled).DynamicMemoryEnabled
    If("$dynamicMemoryEnabledFlag" -eq "False")
    {
        "The memory of the virtual machine $vmName is already set to $vmMemory and dynamic memory is disabled. " | fnWriteLog -LogType Warning -string
    }
    else
    {
        try
        {
            If((Get-VM -Name $vmName).State -eq 'Running') { Stop-VM -Name $vmName -Force }
            Set-VMMemory $vmName -DynamicMemoryEnabled $false -StartupBytes ($vmMemory*1024*1024*1024)
            "The memory for the virtual machine $vmName is set to $vmMemory GB successfully" | fnWriteLog -LogType Success -string
        }
        catch
        {
            "Updating the virtual machine memory to $vmMemory failed with the error: $($_.Exception.Message)" | fnWriteLog -LogType Error -string
        }
    }
}
else
{
   $dynamicMemoryEnabledFlag = (Get-VMMemory -VMName $vmName | Select DynamicMemoryEnabled).DynamicMemoryEnabled
   If("$dynamicMemoryEnabledFlag" -eq "False")
   {
       "The memory of the virtual machine $vmName is already set to $vmMemory and dynamic memory is disabled. " | fnWriteLog -LogType Warning -string
   } 
   else
   {
       If((Get-VM -Name $vmName).State -eq 'Running') { Stop-VM -Name $vmName -Force }
       Set-VMMemory $vmName -DynamicMemoryEnabled $false  
       "The memory of the virtual machine $vmName is already set to $vmMemory. Dynamic memory is now disabled. " | fnWriteLog -LogType Warning -string
   }
}

# Creating a virtual drive as D:\ drive to attach to the virtual machine
If(!(Test-Path -Path $dDrivePath))
{
    try
    {
        New-VHD -Path $dDrivePath -SizeBytes (15 * 1073741824) -Dynamic
        "Virtual hard disk attached to the virtual machine at location $dDrivePath successfully" | fnWriteLog -LogType Success -string
    }
    catch
    {
       "Creating the virtual hard disk at $dDrivePath failed with the error: $($_.Exception.Message) " | fnWriteLog -LogType Error -string
    }
}
else
{
    "Warning: D:\ drive VHDX already exists at path $dDrivePath." | fnWriteLog -LogType Warning -string
}

# Add the drive to the VM
$vmDiskDetails = Get-VMHardDiskDrive -VMName $vmName

If(!($vmDiskDetails | Where {$_.ControllerType -eq 'SCSI'}))
{
    try
    {
        If((Get-VM -Name $vmName).State -eq 'Running') { Stop-VM -Name $vmName -Force }
        Add-VMHardDiskDrive -VMName $vmName -ControllerType SCSI -ControllerNumber 0 -Path "D:\HyperVMachines\$vmName\D_Drive.vhdx"
        "Virtual hard disk added to the virtual machine $vmName successfully." | fnWriteLog -LogType Success -string
    }
    catch
    {
        "Adding virtual hard disk to the virtual machine $vmName failed with the error message: $($_.Exception.Message)" | fnWriteLog -LogType Error -string
    }
}
else
{
    "The D:\ drive is already attached to the virtual machine $vmName." | fnWriteLog -LogType Warning -string
}

# To disable autoamatic checkpoints
If((Get-VM -Name $vmName).AutomaticCheckpointsEnabled -eq 'True')
{
    try
    {
        If((Get-VM -Name $vmName).State -eq 'Running') { Stop-VM -Name $vmName -Force }
        Set-VM -Name $vmName -CheckpointType Standard -AutomaticCheckpointsEnabled:$false
        "Automatic checkpoints disabled successfully." | fnWriteLog -LogType Success -string
    }
    catch
    {
        "Disabling automatic checkpoints failed with the error message: $($_.Exception.Message)" | fnWriteLog -LogType Error -string
    }
}
else
{
    "Automatic Checkpoint is already disabled for the virtual machine $vmName." | fnWriteLog -LogType Warning -string
}

# Start the virtual machine
If((Get-VM -Name $vmName | Select State).State -eq 'Off')
{
    Start-VM $vmName
    "Virutal machine $vmName started successfully." | fnWriteLog -LogType Success -string
}
else
{
    "Virtual machine is already in running state." | fnWriteLog -LogType Warning -string
}

Write-Host “`n Waiting for PowerShell Direct to start on VM $vmName...” -ForegroundColor Yellow

$breakflag = $null
$startTime = Get-Date

While ((Invoke-Command -VMName $vmName -Credential $cred -ScriptBlock {"TestPhrase"} -ea SilentlyContinue) -ne “TestPhrase”) 
{
    Sleep -Seconds 15
    $waitTime = (New-TimeSpan -Start $startTime -End (Get-Date)).minutes
    Write-host " Wait time : $waitTime mins"  -ForegroundColor Yellow
    If($waitTime -gt 2)
    { 
        If((Invoke-Command -VMName $vmName -Credential $cred -ScriptBlock {"TestPhrase"} -ea SilentlyContinue) -ne “TestPhrase”)
        {
            $password = $cred.Password
            $localCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$vmName\administrator",$password
            $cred = $localCredentials
            Write-Host "`n Changing the credentails to local credentails" -ForegroundColor Yellow
            Break
        }
        else
        {
            $breakflag = $true
            Break         
        }
    }
}

# To change the credentials to domain credentails (in case of domain joined)    
If($breakflag -eq $true)
{ 
        Write-Host " Changing to domain credentials as the server seems to be joined to the domain." -ForegroundColor Cyan
        $cred = $domainCred
}

 
Write-Verbose "PowerShell Direct responding on VM $vmName. Moving On...." -Verbose


# To check if the Powershell Direct is working on the virtual machine
try
{
    $hostname = $null
    $hostname = Invoke-Command -VMName $vmName -ScriptBlock { hostname  } -Credential $cred -Verbose
    If($hostname)
    {
        "Powershell Direct is working on the virtual machine $vmName. HOSTNAME output: $hostname" | fnWriteLog -LogType Success -string
    }
    else
    {
        "The virtual machine is not respoding to the Powershell Direct commands. Please check the logs." | fnWriteLog -LogType Error -string
        Exit
    }
}
catch
{
    "Powershell Direct is not working on the virtual machine $vmName. Please check the logs." | fnWriteLog -LogType Error -string
    Break
}

# To rename the virtual machine
$vmRename = Invoke-Command -VMName $vmName -ScriptBlock {
            param($vmName)

            $executionResultObject = @{CustomResult = "None"
                                       CustomMessage = "None"}
            $executionResult = New-Object -TypeName PSObject -Property $executionResultObject

            $hostname = $env:computername
            If($vmName -ne $hostname)
            {
                try
                {
                    Rename-Computer -NewName "$vmName" -Restart -ErrorAction STOP
                    $executionResult.CustomResult = "Success"
                    $executionResult.CustomMessage = "The name of the virtual machine $env:ComputerName is changed to $vmName successfully."
                }
                catch
                {
                    $executionResult.CustomResult = "Failure"
                    $executionResult.CustomMessage = "Renaming the virtual machine from $env:ComputerName failed with the error : $_.Exception "
                }
            }
            else
            {
                $executionResult.CustomResult = "Warning"
                $executionResult.CustomMessage = "The suggested new name and the current virtual machine name is same."
            }
            Return $executionResult
    } -ArgumentList $vmName -Credential $cred -Verbose

$outputMessage = $vmRename.CustomMessage
Switch ($vmRename.CustomResult)
{  
    "Success" { "Success : $outputMessage" | fnWriteLog -LogType Success -string }
    "Failure" { "Error   : $outputMessage" | fnWriteLog -LogType Error -string }
    "Warning" { "Warning : $outputMessage" | fnWriteLog -LogType Warning -string }
}

If($vmRename.CustomResult -eq "Success" -or $vmRename.CustomResult -eq "Failure")
{
    Write-Verbose “Waiting for virutal machine to boot up after hostname change...” -Verbose

    While ((Invoke-Command -VMName $vmName -Credential $cred -ScriptBlock {"TestPhrase"} -ea SilentlyContinue) -ne “TestPhrase”) 
    {
        Sleep -Seconds 15
        $waitTime = (New-TimeSpan -Start $startTime -End (Get-Date)).minutes
        Write-host " Wait time : $waitTime mins"  -ForegroundColor Yellow
        If($waitTime -gt 5)
        { 
            If((Invoke-Command -VMName $vmName -Credential $cred -ScriptBlock {"TestPhrase"} -ea SilentlyContinue) -ne “TestPhrase”)
            {
                $password = $cred.Password
                $localCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$vmName\administrator",$password
                $cred = $localCredentials
                Write-Host "`n Changing the credentails to local credentails" -ForegroundColor Yellow
                Break
            }
            else
            {
                $breakflag = $true
                Break         
            }
        }
    }
 
    Write-Verbose "PowerShell Direct responding on VM $vmName. Moving On...." -Verbose
}

# To set the IP Address of the internal network adaptor on virtual machine
$hostmacaddressinternalswitch = (Get-VMNetworkAdapter -VMName $vmName | Where-Object {$_.SwitchName -eq "$vmNetworkSwitch"}).MacAddress
If($hostmacaddressinternalswitch)
{
    $hostmacaddressinternalswitch = (&{for ($counter = 0;$counter -lt $hostmacaddressinternalswitch.length;$counter += 2)
    {
        $hostmacaddressinternalswitch.substring($counter,2)
    }
    }) -join '-'
}

$updateVmIP = Invoke-Command -VMName $vmName -ScriptBlock {
            param($intipAddress,$hostmacaddressinternalswitch)
            $executionResultObject = @{CustomResult = "None"
                                       CustomMessage = "None"}
            $executionResult = New-Object -TypeName PSObject -Property $executionResultObject

            $networkName = (Get-NetAdapter | Where-Object {$_.MacAddress -eq "$hostmacaddressinternalswitch"}).Name
            $interfaceIndex = (Get-NetAdapter | where { $_.Name -eq $networkName }).InterfaceIndex
            $currentIPAddress = (Get-NetIPAddress -InterfaceIndex $interfaceIndex | Where-Object {$_.AddressFamily -eq 'IPv4'}).IPAddress
             If($currentIPAddress -ne $intipAddress)            {                try                {                    New-NetIPAddress –IPAddress $intipAddress -PrefixLength 24 -InterfaceIndex $interfaceIndex                    $executionResult.CustomResult = "Success"
                    $executionResult.CustomMessage = ""                }                catch                {                    $executionResult.CustomResult = "Failure"
                    $executionResult.CustomMessage = "Exception: $_.Exception"                }            }
            else            {                $executionResult.CustomResult = "Warning"
                $executionResult.CustomMessage = ""            }
        Return $executionResult
    } -ArgumentList $intipAddress,$hostmacaddressinternalswitch -Credential $Cred -Verbose -ErrorAction STOP

$outputMessage = $updateVmIP.CustomMessage
Switch ($updateVmIP.CustomResult)
{  
    "Success" { "Success : The IP address of the virtual switch $vmNetworkSwitch is set to IP: $intipAddress. $outputMessage" | fnWriteLog -LogType Success -string }
    "Failure" { "Error   : Error while setting the IP address $intipAddress to the virtual switch $vmNetworkSwitch. $outputMessage" | fnWriteLog -LogType Error -string }
    "Warning" { "Warning : The IP address of the virtual switch $vmNetworkSwitch is already set to $intipAddress. $outputMessage" | fnWriteLog -LogType Warning -string }
}

# To set the IP Address of the internet network adaptor on virtual machine
$hostmacaddressinternetswitch = (Get-VMNetworkAdapter -VMName $vmName | Where-Object {$_.SwitchName -eq "$vmInternetSwitch"}).MacAddress
If($hostmacaddressinternetswitch)
{
    $hostmacaddressinternetswitch = (&{for ($counter = 0;$counter -lt $hostmacaddressinternetswitch.length;$counter += 2)
    {
        $hostmacaddressinternetswitch.substring($counter,2)
    }
    }) -join '-'
}

$updateVmIP = Invoke-Command -VMName $vmName -ScriptBlock {
            param($extipAddress,$hostmacaddressinternetswitch,$internetGateway)
            $executionResultObject = @{CustomResult = "None"
                                       CustomMessage = "None"}
            $executionResult = New-Object -TypeName PSObject -Property $executionResultObject

            $networkName = (Get-NetAdapter | Where-Object {$_.MacAddress -eq "$hostmacaddressinternetswitch"}).Name
            $interfaceIndex = (Get-NetAdapter | where { $_.Name -eq $networkName }).InterfaceIndex
            $currentIPAddress = (Get-NetIPAddress -InterfaceIndex $interfaceIndex | Where-Object {$_.AddressFamily -eq 'IPv4'}).IPAddress
             If($currentIPAddress -ne $extipAddress)            {                try                {                    New-NetIPAddress –IPAddress $extipAddress -PrefixLength 24 -InterfaceIndex $interfaceIndex -DefaultGateway $internetGateway                    $executionResult.CustomResult = "Success"
                    $executionResult.CustomMessage = ""                }                catch                {                    $executionResult.CustomResult = "Failure"
                    $executionResult.CustomMessage = "Exception: $_.Exception"                }            }
            else            {                $executionResult.CustomResult = "Warning"
                $executionResult.CustomMessage = ""            }
        Return $executionResult
    } -ArgumentList $extipAddress,$hostmacaddressinternetswitch,$internetGateway -Credential $Cred -Verbose -ErrorAction STOP

$outputMessage = $updateVmIP.CustomMessage
Switch ($updateVmIP.CustomResult)
{  
    "Success" { "Success : The IP address of the virtual switch $vmInternetSwitch is set to IP: $extipAddress. $outputMessage" | fnWriteLog -LogType Success -string }
    "Failure" { "Error   : Error while setting the IP address $extipAddress to the virtual switch $vmInternetSwitch. $outputMessage" | fnWriteLog -LogType Error -string }
    "Warning" { "Warning : The IP address of the virtual switch $vmInternetSwitch is already set to $intipAddress. $outputMessage" | fnWriteLog -LogType Warning -string }
}

# To enable Ping and RDP Access to the virtual machine
Invoke-Command -VMName $vmName -ScriptBlock {
    netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol="icmpv4:8,any" dir=in action=allow
    Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server'-name "fDenyTSConnections" -Value 0
    Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
} -Credential $cred -Verbose


# To test the TCP Connection
try
{
    $tcpConnection = Test-NetConnection $intipAddress -CommonTCPPort rdp
}
catch
{
    $tcpConnection = $_.Exception.Message
}

If($tcpConnection.TcpTestSucceeded -eq "True")
{
    "TCP Connection succeeded for the virtual machine $vmName" | fnWriteLog -LogType Success -string
}
else
{
    "TCP Connection failed for the virtual machine $vmName. Error Message : $tcpConnection" | fnWriteLog -LogType Success -string
}

# To initialize and configure additional drive on the virtual machine
$diskAddition = Invoke-Command -VMName $vmName -ScriptBlock {
        $executionResultObject = @{CustomResult = "None"
                                   CustomMessage = "None"}
        $executionResult = New-Object -TypeName PSObject -Property $executionResultObject

        try
        {
            $diskNumber = (Get-Disk | Where-Object { $_.OperationalStatus -eq 'Offline' }).Number
            If($diskNumber)
            {
                Set-Disk -NUmber $diskNumber -IsOffline $false
                Initialize-Disk -Number $diskNumber -PartitionStyle MBR
                New-Partition -DiskNumber $diskNumber -UseMaximumSize -AssignDriveLetter | Format-Volume -Confirm:$false -FileSystem NTFS -Force

                $executionResult.CustomResult = "Success"
                $executionResult.CustomMessage = "The virtual disk is initialized, formatted and attached successfully."
            }
            else
            {
                $executionResult.CustomResult = "Warning"
                $executionResult.CustomMessage = "There are no disks to be initialized and formatted." 
            }
        }
        catch
        {
            $executionResult.CustomResult = "Error"
            $executionResult.CustomMessage = "Addition of the virutal disk failed on the virtual machine. Error: $_.Exception"
        }
        Return $executionResult
    } -Credential $cred -Verbose

$outputMessage = $diskAddition.CustomMessage
Switch ($diskAddition.CustomResult)
{  
    "Success" { "Success : $outputMessage" | fnWriteLog -LogType Success -string }
    "Failure" { "Error   : $outputMessage" | fnWriteLog -LogType Error -string }
    "Warning" { "Warning : $outputMessage" | fnWriteLog -LogType Warning -string }
}

# To change the default user language to English
$vmLanguageSetup = Invoke-Command -VMName $vmName -ScriptBlock {
            try
            {
                $UserLanguageList = New-WinUserLanguageList -Language en-US
                $UserLanguageList[0].Handwriting = $True
                Set-WinUserLanguageList -LanguageList $UserLanguageList -Force
                Return "Success"
            }
            catch
            {
                Return $_.Exception.Message
            }
} -Credential $cred -Verbose

If($vmLanguageSetup -eq "Success")
{
    "The user interface language of virtual machine $vmName is set to English-US successfully." | fnWriteLog -LogType Success -string
}
else
{
    "Updating the user interface language of the virtual machine $vmName failed with the error : $vmLanguageSetup"| fnWriteLog -LogType Error -string
}

# To set the DNS address to the domain controller IP for internal network adaptor
$hostmacaddressinternalswitch = (Get-VMNetworkAdapter -VMName $vmName | Where-Object {$_.SwitchName -eq "$vmNetworkSwitch"}).MacAddress
If($hostmacaddressinternalswitch)
{
    $hostmacaddressinternalswitch = (&{for ($counter = 0;$counter -lt $hostmacaddressinternalswitch.length;$counter += 2)
    {
        $hostmacaddressinternalswitch.substring($counter,2)
    }
    }) -join '-'
}

$internalDnsIPUpdate = Invoke-Command -VMName $vmName -ScriptBlock {
        param($domainIPAddress,$hostmacaddressinternalswitch)

        $executionResultObject = @{CustomResult = "None"
                                   CustomMessage = "None"}
        $executionResult = New-Object -TypeName PSObject -Property $executionResultObject
        $dnsIPAddress = (Get-DNSClientServerAddress -InterfaceIndex (Get-NetAdapter | Where-Object {$_.MacAddress -eq "$hostmacaddressinternalswitch"} | Select InterfaceIndex).InterfaceIndex | Select ServerAddresses).ServerAddresses
        If($dnsIPAddress -contains $domainIPAddress)
        {
            $executionResult.CustomResult = "Warning"
            $executionResult.CustomMessage = "The DNS IP address of the virtual machine $env:ComputerName is already set to $domainIPAddress."
        }
        else
        {
            try
            {
                Set-DNSClientServerAddress -InterfaceIndex (Get-NetAdapter | Where-Object {$_.MacAddress -eq "$hostmacaddressinternalswitch"} | Select InterfaceIndex).InterfaceIndex –ServerAddresses ("$domainIPAddress") -ErrorAction STOP
                $executionResult.CustomResult = "Success"
                $executionResult.CustomMessage = "The DNS IP address of the virtual machine $env:ComputerName is set to $domainIPAddress successfully."
            }
            catch
            {
                $executionResult.CustomResult = "Failure"
                $executionResult.CustomMessage = "Failed to set DNS IP address of the virtual machine $env:ComputerName to $domainIPAddress. Exception : $_.Exception"
            }
        }
        Return $executionResult

    } -ArgumentList $domainIPAddress,$hostmacaddressinternalswitch -Credential $cred -Verbose

$outputMessage = $internalDnsIPUpdate.CustomMessage
Switch ($internalDnsIPUpdate.CustomResult)
{
    "Success" { "Success : $outputMessage" | fnWriteLog -LogType Success -string }
    "Failure" { "Error   : $outputMessage" | fnWriteLog -LogType Error -string }
    "Warning" { "Warning : $outputMessage" | fnWriteLog -LogType Warning -string }
}

# To set the DNS address to the google DNS IP for internet network adaptor
$hostmacaddressinternetswitch = (Get-VMNetworkAdapter -VMName $vmName | Where-Object {$_.SwitchName -eq "$vmInternetSwitch"}).MacAddress
If($hostmacaddressinternetswitch)
{
    $hostmacaddressinternetswitch = (&{for ($counter = 0;$counter -lt $hostmacaddressinternetswitch.length;$counter += 2)
    {
        $hostmacaddressinternetswitch.substring($counter,2)
    }
    }) -join '-'
}

$externalDnsIPUpdate = Invoke-Command -VMName $vmName -ScriptBlock {
        param($hostmacaddressinternetswitch)

        $executionResultObject = @{CustomResult = "None"
                                   CustomMessage = "None"}
        $executionResult = New-Object -TypeName PSObject -Property $executionResultObject
        $dnsIPAddress = (Get-DNSClientServerAddress -InterfaceIndex (Get-NetAdapter | Where-Object {$_.MacAddress -eq "$hostmacaddressinternetswitch"} | Select InterfaceIndex).InterfaceIndex | Select ServerAddresses).ServerAddresses
        If($dnsIPAddress -contains '8.8.8.8')
        {
            $executionResult.CustomResult = "Warning"
            $executionResult.CustomMessage = ""
        }
        else
        {
            try
            {
                Set-DNSClientServerAddress -InterfaceIndex (Get-NetAdapter | Where-Object {$_.MacAddress -eq "$hostmacaddressinternetswitch"} | Select InterfaceIndex).InterfaceIndex –ServerAddresses ('8.8.8.8') -ErrorAction STOP
                $executionResult.CustomResult = "Success"
                $executionResult.CustomMessage = ""
            }
            catch
            {
                $executionResult.CustomResult = "Failure"
                $executionResult.CustomMessage = "Exception : $_.Exception"
            }
        }
        Return $executionResult

    } -ArgumentList $hostmacaddressinternetswitch -Credential $cred -Verbose

$outputMessage = $externalDnsIPUpdate.CustomMessage
Switch ($externalDnsIPUpdate.CustomResult)
{
    "Success" { "Success : The DNS IP address of the virtual adaptor $vmInternetSwitch is set to 8.8.8.8 successfully. $outputMessage" | fnWriteLog -LogType Success -string }
    "Failure" { "Error   : Failed to set DNS IP address of the virtual adatpor $vmInternetSwitch  to 8.8.8.8. $outputMessage" | fnWriteLog -LogType Error -string }
    "Warning" { "Warning : The DNS IP address of the virtual machine $vmInternetSwitch is already set to 8.8.8.8. $outputMessage" | fnWriteLog -LogType Warning -string }
}

#To add a server to the domain
$domainAddition = Invoke-Command -VMName $vmName -ScriptBlock {
                  param($domainName,$domainCred)
                  $executionResultObject = @{CustomResult = "None"
                                             CustomMessage = "None"}
                  $executionResult = New-Object -TypeName PSObject -Property $executionResultObject

                  If((Get-WmiObject -Class Win32_ComputerSystem).Domain -eq 'WORKGROUP')
                  {
                      try
                      {
                          Add-Computer -DomainName $domainName -Credential $domainCred -Restart -Force -ErrorAction STOP
                          $executionResult.CustomResult = "Success"
                          $executionResult.CustomMessage = "The virtual machine is added to the doamin $domainName successfully."
                      }
                      catch
                      {
                          $executionResult.CustomResult = "Failure"
                          $executionResult.CustomMessage = $_.Exception
                      }
                  }
                  else
                  {
                      $executionResult.CustomResult = "Warning"
                      $executionResult.CustomMessage = "The virtual machine is already a part of domain $domainName."
                  }
                  Return $executionResult
    } -ArgumentList $domainName,$domainCred -Credential $cred -Verbose

$outputMessage = $domainAddition.CustomMessage
Switch ($domainAddition.CustomResult)
{
    "Success" { "Virtual machine $vmName added to the domain $domain successfully." | fnWriteLog -LogType Success -string }
    "Failure" { "Adding the virtual machine $vmName to the domain $domainName failed. Error : $outputMessage " | fnWriteLog -LogType Error -string }
    "Warning" { "Warning Message: $outputMessage" | fnWriteLog -LogType Warning -string }
}

<#
# To install the AD role on the server
Invoke-Command -VMName $vmName -ScriptBlock {
        Install-WindowsFeature AD-Domain-Services -IncludeManagementTools
} -Credential $cred -Verbose

# To configure the server as Domain controller
Invoke-Command -VMName $vmName -ScriptBlock {
        Install-ADDSForest -DomainName sakharam.com -SafeModeAdministratorPassword $domainCred.Password -Force
}
#>


#################Common Code Started#####################

#Creating the HTML Report...
ConvertTo-EnhancedHTML -HTMLFragments "$htmlbody $newVMSQLSummary" -CssStyleSheet $CSSStyle -PreContent "<h1> $ReportName for $serverName : $CurrentDateTime </h1>" | Out-File $LogFile
Invoke-Item $LogFile


#To Sakharam - With the errors
#$emailbody = ConvertTo-EnhancedHTML -HTMLFragments "$htmlbody $cpuSummary" -CssStyleSheet $CSSStyle | Out-String
#fnSendEmail -fromEmail "ODDBSupport@Advent.com" -toEmail "vpote@sscinc.com" -emailSubject "CPU Utilization Report for $serverName : $CurrentDateTime" -emailBody $emailBody 

#################Common Code Started#####################
