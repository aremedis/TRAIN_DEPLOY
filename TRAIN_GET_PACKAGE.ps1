<#
This script will grab the latest install script package from HOPS2, and then edit the paramater.csv file with the relevant server information.

# Changelog
20Aug2014 - Script Created
06Oct2014 - Changed source server to CLASS domain controller
13Oct2014 - Added code for renaming preinstalled sql server instance
    http://agilebi.com/ddarden/2009/04/06/powershell-script-to-reset-the-local-instance-of-sql-server/
13Oct2014 - Added comments
14Oct2014 - added code to choose a different source based on subnet
            added code to remove the installation package once install complete
#>



# Load Assemblies we need to access SMO

$asm = [reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo")

$asm = [reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo")

$asm = [reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoEnum")

$asm = [reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.SqlEnum")

$asm = [reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.WmiEnum")

$asm = [reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.SqlWmiManagement")

#find the current Computer's IP address
$IPAddress = (Get-WmiObject -class win32_NetworkAdapterConfiguration -Filter 'ipenabled = "true"').ipaddress[0]
[Environment]::SetEnvironmentVariable("IPAddress", $IPAddress, "User")

$splitIP = $IPAddress.split(".")
$subnet = $splitIP[3]

## Declare IP ranges based on office Location

$ChicagoSubnet = "15"
$ViennaSubnet = "14"
$AustinSubnet = ""

function DetermineSource()
{
    if ($subnet -eq $ChicagoSubnet)
    {
        return "\\rosemontclassdc.class.tmaresources.com\PersonifyInstallation\7.5.2"
    }
    elseif ($subnet -eq $ViennaSubnet)
    {
        return "\\class-dc1.class.tmaresources.com\c$\support\Software\7.5.2"
    }
    elseif ($subnet -eq $AustinSubnet)
    {
        return "\\class-dc1.class.tmaresources.com\c$\support\Software\7.5.2"
    }
    else 
    {
        return "\\class-dc1.class.tmaresources.com\c$\support\Software\7.5.2"
    }
}

# Declaring environment variables
$source = DetermineSource
$destination = "C:\support\7.5.2"
$CurrentComputer = $env:computername
$ParamFile = "C:\support\7.5.2\PersonifyInstallation\PersonifyInstallationParameters.csv"
$retry = 0


## define functions

function CopyPackage()  #Copy's the install package and scripts from $source
    {
        "Starting Copy"
        Copy-Item $source -Destination $destination -recurse -force
        #if the copy was successful edit the paramaters of the install .csv
        if (TestCopy -eq "True")
            {
                "Copy Successful"
                SetParams
            }
        else 
            {
                #attempt to re-copy the folder
                if(!($retry -gt 2))
                    {
                        $retry ++
                        "Retrying to copy"
                        CopyPackage

                    }
                else 
                    {
                        "Unable to confirm the integrity of the source vs destination"
                        exit
                    }
                
                        
            }

    }


function SetParams() #adds the parameter values to the .csv files for the install scripts
    {
    ## Paramater String
    #  AddCustomDataTypes,Drive,EnvironmentType,Location,OrganizationID,OrganizationUnitID,UseHostHeader,ApplicationServer,BusinessObjectsServer,BusinessOBjectsServerIP,DatabaseServer,PBBIserver,SMTPServer,TRSServer,WebServer,WebSiteIP,ApplicationServicePortInt,ApplicationServicePortExt,CacheServicePortInt,InterfaceServicePortInt,ApplicationServiceURL,AutoUpdatesURL,BusinessObjectsURL,DataServicesURL,EbusinessURL,HomePageURL,PADSSURL,SSOURL,ServiceAccount,ServiceAccountPassword,ApplicationPoolAccount,ApplicationPoolAccountPassword,BOEAccount,BOEAccountPassword,PBBIuser,PBBIpassword
    ##
    "Setting the Parameters in the .csv file"
    $NewLine = "N,C:,TRAIN,Local,$CurrentComputer,TRAIN,N,localhost,class-vpboe1,172.22.14.21,$CurrentComputer,localhost,localhost,localhost,localhost,$IPAddress,9000,9000,9300,9200,localhost,localhost,localhost,localhost,localhost,localhost,localhost,localhost,CLASS\Class_user,training,CLASS\Class_user,training,CLASS\Class_user,training,CLASS\Class_user,training"
    $NewLine | add-content -Path $ParamFile


    }

function ConnHops()
    {
            if (!($uncServer))
            {
                $uncServer = DetermineSource
                net use $uncServer 
            }
            else{
                net use $uncserver /delete
            }
    }

#Main Function
function GoNow()
    {
    
    if (!(Test-Path -Path $source))
        {
            ConnHops
            GoNow
        }
    else 
        {
            CopyPackage
        }
    }

#test that the source and destination folders are the same
function TestCopy()
    {
        $a = Get-ChildItem $destination -recurse
        $b = Get-ChildItem $source -recurse
        if ($a.count -eq $b.count)
            {
                return "True"
            }
        else 
            {
                return "False"
            }
    }

function LaunchInstall()
    {
        cd c:\
        .\support\7.5.2\PersonifyInstallation\PS_PersonifyInstallation.ps1
    }

function DesktopIcon()
    {
        Copy-Item "C:\Personify\$CurrentComputer\TRAIN\Web\AutoUpdates" -Destination "C:\Users\Public" -recurse -force
        $shell = New-Object -ComObject WScript.shell
        $desktop = "C:\Users\Public\Desktop"
        $shortcut = $shell.CreateShortcut("$desktop\Personify 360.lnk")
        $shortcut.TargetPath = "C:\Users\Public\AutoUpdates\Personify.exe"
        $shortcut.IconLocation = "C:\Users\Public\AutoUpdates\Personify.exe"
	    $shortcut.WorkingDirectory = "C:\Users\Public\AutoUpdates"
        $shortcut.Save()

    }

# Change the name SQL Server instance name (stored inside SQL Server) to the name of the machine.

function ServerInstanceName{

    Write "Renaming SQL Server Instance"

    $smo = 'Microsoft.SqlServer.Management.Smo.'

    $server = new-object ($smo + 'server') .

    $database = $server.Databases["master"]

    $mc = new-object ($smo + 'WMI.ManagedComputer') .


    $newServerName = $mc.Name


    $database.ExecuteNonQuery("EXEC sp_dropserver @@SERVERNAME")

    $database.ExecuteNonQuery("EXEC sp_addserver '$newServerName', 'local'")


    Write-Host "Renamed server to '$newServerName'`n"

}

Function Restart-Service
{
    PARAM([STRING]$SERVERNAME=$ENV:COMPUTERNAME,[STRING]$SERVICENAME)
#Default instance – MSSQLSERVER  ,
#Named instance is MSSQL$KAT – Try to retain its value by negating "$" meaning using "`"
#hence you need to pass service name like MSSQL`$KAT
    $SERVICE = GET-SERVICE -COMPUTERNAME $SERVERNAME -NAME $SERVICENAME -ERRORACTION SILENTLYCONTINUE
    IF( $SERVICE.STATUS -EQ "RUNNING" )
        {
            $DEPSERVICES = GET-SERVICE -COMPUTERNAME $SERVERNAME -Name $SERVICE.SERVICENAME -DEPENDENTSERVICES | WHERE-OBJECT {$_.STATUS -EQ "RUNNING"}
            IF( $DEPSERVICES -NE $NULL )
                {
                    FOREACH($DEPSERVICE IN $DEPSERVICES)
                        {
                            Stop-Service -InputObject (Get-Service -ComputerName $SERVERNAME -Name $DEPSERVICES.ServiceName)
                        }
                }
            Stop-Service -InputObject (Get-Service -ComputerName $SERVERNAME -Name $SERVICE.SERVICENAME) -Force
            if($?)
                {
                    Start-Service -InputObject (Get-Service -ComputerName $SERVERNAME -Name $SERVICE.SERVICENAME)
                    $DEPSERVICES = GET-SERVICE -COMPUTERNAME $SERVERNAME -NAME $SERVICE.SERVICENAME -DEPENDENTSERVICES | WHERE-OBJECT {$_.STATUS -EQ "STOPPED"}
                    IF( $DEPSERVICES -NE $NULL )
                        {
                            FOREACH($DEPSERVICE IN $DEPSERVICES)
                                {
                                    Start-Service -InputObject (Get-Service -ComputerName $SERVERNAME -Name $DEPSERVICE.SERVICENAME)
                                }
                        }
                }
        }
    ELSEIF ( $SERVICE.STATUS -EQ "STOPPED" )
        {
            Start-Service -InputObject (Get-Service -ComputerName $SERVERNAME -Name $SERVICE.SERVICENAME)
            $DEPSERVICES = GET-SERVICE -COMPUTERNAME $SERVERNAME -NAME $SERVICE.SERVICENAME -DEPENDENTSERVICES | WHERE-OBJECT {$_.STATUS -EQ "STOPPED"}
            IF( $DEPSERVICES -NE $NULL )
                {
                    FOREACH($DEPSERVICE IN $DEPSERVICES)
                        {
                            Start-Service -InputObject (Get-Service -ComputerName $SERVERNAME -Name $DEPSERVICE.SERVICENAME)
                        }
                }
        }
    ELSE
    {
        "THE SPECIFIED SERVICE DOES NOT EXIST"
    }
}



## begin script




ServerInstanceName    # MSSQLSERVER service needs to be restarted after this change

Restart-Service -SERVICENAME MSSQLSERVER   #restartes the SQL server service



GoNow

"Finished Copying, Launching Personify Installation in 10 seconds"

Start-Sleep -s 10

LaunchInstall  #Launches Brant's install scripts from $destination

DesktopIcon  #Create Desktop Shortcut for All Users


# ConnHops

"Removing Installation Package"
#Delete the install files after installation is complete
function Get-Tree($Path,$Include='*') { 
    @(Get-Item $Path -Include $Include) + 
        (Get-ChildItem $Path -Recurse -Include $Include) | 
        sort pspath -Descending -unique
} 

function Remove-Tree($Path,$Include='*') { 
    Get-Tree $Path $Include | Remove-Item -force -recurse
} 

Remove-Tree $destination
