<#
.SYNOPSIS
Connect to a SAP host and validate the password for the account is correct by using RFC.
#>
function Validate-SAPPassword
{
	[CmdletBinding()]
	param (
		[String]$HostName,
		[String]$UserName,
		[String]$CurrentPassword,
        [String]$PrivilegedAccountUserName,
        [String]$GenericField1,
        [String]$GenericField2,
        [String]$GenericField3
       	)

    $Client = $GenericField1
    $Instance = $GenericField2
    $SID = $GenericField3
	
	try
	{
		   $Path = "C:\Program Files\SAP\SAP_DotNetConnector3_Net40_x64\"

          [String]$File = $Path + "sapnco.dll"
          Add-Type -Path $File
          [String]$File = $Path + "sapnco_utils.dll"
          Add-Type -Path $File
 
            $cfgParams = New-Object SAP.Middleware.Connector.RfcConfigParameters
            $cfgParams.Add($cfgParams::Name, $SID)
            $cfgParams.Add($cfgParams::AppServerHost, $HostName)
            $cfgParams.Add($cfgParams::SystemNumber, $Instance)
            $cfgParams.Add($cfgParams::Client, $Client)
            $cfgParams.Add($cfgParams::User, $UserName)

            $cfgParams.Add($cfgParams::Password, $CurrentPassword)
            $destination = [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($cfgParams)
		
		Try {
                      #-Metadata----------------------------------------------------
            $rfcFunction = $destination.Repository.CreateFunction("RFC_PING")
            $destination.Ping()
            $resultsarray = "Ping successful"
            
            }
            Catch {
              $resultsarray = $_.Exception.Message
            }
        
		switch -wildcard ($resultsarray.ToString().ToLower())
		{
			"*Host Offline*" { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName' - it appears the Host is not online, or a firewall is blocking port 135 and 445."; break }
			"*You cannot call a method on a null-valued expression*" { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName' - UserName or Password is incorrect."; break } #TODOD
            "*Access is denied*" { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName' - it appears the account may be locked out."; break }
			#Add other wildcard matches here as required
            default { Write-Output "Success" }
		}
	}
	catch
	{
		switch -wildcard ($error[0].Exception.ToString().ToLower())
		{
            "*The network path was not found*" { Write-Output "Success"; break } #Can return this value if the password is correct for PowerShell 5. If incorrect, it will return a different value
            "*disabled*" { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName' - it appears the account is disabled."; break }
			"*No such host is known*" { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName' - it appears the Host is not online or no longer exists."; break }
            "*locked out*" { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName' - it appears the account may be locked out."; break }
            "*an extended error*" { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName' - UserName or Password is incorrect."; break }
            #Using ValidateCredentials can still cause access denied because of UAC, so we need to allow for it
            "*Access is denied*" 
            { 
                if ($error[0].Exception.ToString() -like '*System.DirectoryServices.AccountManagement.PrincipalOperationException*') 
                { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName' - it appears the account is disabled or locked out, or Username or Password is incorrect."; break }
                else
                { Write-Output "Success"; break }
            } 

			#Add other wildcard matches here as required
			default { Write-Output "Failed to validate the local password for account '$UserName' on Host '$HostName'. Error = " $error[0].Exception }
		}
	}
	finally {  }
}


#Make a call to the Validate-SAPPassword function
#Validate-SAPPassword -HostName '[HostName]' -UserName '[UserName]' -CurrentPassword '[CurrentPassword]' -Client '[GenericField1]' -Instance '[GenericField2]' -SID '[GenericField3]'

Validate-SAPPassword -HostName '[HostName]' -UserName '[UserName]' -CurrentPassword '[CurrentPassword]' -GenericField1 '[GenericField1]' -GenericField2 '[GenericField2]' -GenericField3 '[GenericField3]'