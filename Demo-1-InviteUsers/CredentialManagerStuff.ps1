 Install-Module CredentialManager 

 New-StoredCredential -Target "opaxAdmin" -UserName "olli@opax.onmicrosoft.com" -Password "sss" -Type Generic -Persist LocalMachine

 Get-StoredCredential -Target "opaxAdmin" 