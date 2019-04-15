<#
.SYNOPSIS
Connect to a SAP host using the supplied Privileged Account Credentials, and change the password for a local account.
#>
function Set-SAPPassword
{
	[CmdletBinding()]
	param (
		[String]$HostName,
		[String]$UserName,
        [String]$OldPassword,
		[String]$NewPassword,
		[String]$PrivilegedAccountUserName,
		[String]$PrivilegedAccountPassword,
        [String]$GenericField1,
        [String]$GenericField2,
        [String]$GenericField3
	)

    $Client = $GenericField1
    $Instance = $GenericField2
    $SID = $GenericField3


	
	try
	{
		#Establish the PowerShell Credentials used to execute the script block - based on the Privileged Account Credentials selected for this script	
        if ($PrivilegedAccountUserName -eq '') {
            $CredPassword = $PrivilegedAccountPassword
            $PSUsername = $PrivilegedAccountUserName
        }
        else {
            $CredPassword = $OldPassword
            $PSUsername = $UserName
        }

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
            $cfgParams.Add($cfgParams::User, $PSUserName)

            $cfgParams.Add($cfgParams::Password, $CredPassword)
            $destination = [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($cfgParams)
		
		Try {
                      #-Metadata----------------------------------------------------
              $rfcFunction = $destination.Repository.CreateFunction("ME_USER_CHANGE_PASSWORD")
              $rfcFunction.SetValue("USERNAME", $UserName)
              $rfcFunction.SetValue("PASSWORD", $OldPassword)
              $rfcFunction.SetValue("NEW_PASSWORD", $NewPassword)
              #$rfcFunction.SetValue("USE_BAPI_RETURN", "1")

              $rfcFunction.Invoke($destination)
              
              if ($rfcFunction.GetValue("SUCCESSFUL") -eq "X") {
              $resultsarray = "Success"
              } else {
              $resultsarray = $rfcFunction.GetValue("ERROR_MESSAGE")
              }
              
            }
            Catch {
              $resultsarray = $_.Exception.Message
            }		
		
		if ($resultsarray -eq "Success")
		{
			Write-Output "Success"
		}
		else
		{
			switch -wildcard ($resultsarray.ToString().ToLower())
			{
				"*Access is denied*" { Write-Output "Failed to reset the local SAP account '$UserName' on Host '$HostName' due to the error 'Access is Denied. Please refer to Microsoft's documentation about 'about_Remote_Troubleshooting', in particular the 'LocalAccountTokenFilterPolicy' registry key if remoting in with a Local Administrator's account."; break }
				"*WinRM cannot complete the operation*" { Write-Output "Failed to reset the local Windows password for account '$UserName' on Host '$HostName' as it appears the Host is not online, or PowerShell Remoting is not enabled."; break }
				"*WS-Management service running*" { Write-Output "Failed to reset the local Windows password for account '$UserName' on Host '$HostName' as it appears the Host is not online, or PowerShell Remoting is not enabled."; break }
				"*cannot find the computer*" { Write-Output "Failed to reset the local Windows password for account '$UserName' on Host '$HostName' as it appears the Host is not online, or PowerShell Remoting is not enabled."; break }
				"*no logon servers available*" { Write-Output "Failed to reset the local Windows password for account '$UserName' on Host '$HostName'. There are currently no logon servers available to service the logon request."; break }
				"*currently locked*" { Write-Output "Failed to reset the local password for account '$UserName' on Host '$HostName'. The referenced account is currently locked out and may not be logged on to."; break }
				"*user name or password is incorrect*" { Write-Output "Failed to reset the local password for account '$UserName' on Host '$HostName' as the Privileged Account password appears to be incorrect, or the account is currently locked."; break }
				"*username does not exist*" { Write-Output "Failed to reset the local password for account '$UserName' on Host '$HostName' as the UserName does not exist."; break }
				#Add other wildcard matches here as required
				default { Write-Output "Failed to reset the local password for account '$UserName' on Host '$HostName'.Error = $resultsarray." }
			}
		}
	}
	catch
	{
		switch -wildcard ($error[0].Exception.ToString().ToLower())
		{
			"*The user name or password is incorrect*" { Write-Output "Failed to connect to the Host '$HostName' to reset the password for the account '$UserName'. Please check the Privileged Account Credentials provided are correct."; break }
			"*cannot bind argument to parameter*" { Write-Output "Failed to reset the local password for account '$UserName' on Host '$HostName' as it appears you may not have associated a Privileged Account Credential on the Reset Options tab for the Password record."; break }
			#Add other wildcard matches here as required
			default { Write-Output "Failed to reset the local Windows password for account '$UserName' on Host '$HostName'. Error = " $error[0].Exception }
		}
	}
}

#Make a call to the Set-SAPPassword function
Set-SAPPassword -HostName '[HostName]' -UserName '[UserName]' -OldPassword '[OldPassword]' -NewPassword '[NewPassword]' -PrivilegedAccountUserName '[PrivilegedAccountUserName]' -PrivilegedAccountPassword '[PrivilegedAccountPassword]' -GenericField1 '[GenericField1]' -GenericField2 '[GenericField2]' -GenericField3 '[GenericField3]'