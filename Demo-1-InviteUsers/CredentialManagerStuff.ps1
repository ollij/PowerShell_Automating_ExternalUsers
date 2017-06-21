 Install-Module CredentialManager 

 New-StoredCredential -Target "yourStoredCredentialTarget" -UserName "you@yourtenant.onmicrosoft.com" -Password "yourPassword" -Type Generic -Persist LocalMachine

 Get-StoredCredential -Target "yourStoredCredentialTarget" 
