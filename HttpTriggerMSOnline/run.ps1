using namespace System.Net

# Set a parameter
param($Request)

# Instantiate Credentials
$username = "<username>"
$password = ConvertTo-SecureString "<password>" -AsPlainText -Force
$psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)

# Import the MSOnline Module and establish a session to Exchange Online
# Import-Module MSOnline
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $psCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

# Create a variable to pass a value from the Request to Powershell
$name = $Request.Query.Name

# Create a Shared Mailbox
$cmd = "New-Mailbox –Name """ + $name + """ -Shared"
$cmdScriptBlock = [Scriptblock]::Create($cmd)
Invoke-Command -Session $Session -ScriptBlock $cmdScriptBlock

# Compose an HTTP Response to send when the command is run 
$HttpResponse = [HttpResponseContext]@{ 
    StatusCode = 200
    Body       = "Success" 
} 

# Send the Response back 
Push-OutputBinding -Name Response -Value $HttpResponse

# Terminate the Exchange Session
Remove-PSSession $Session
