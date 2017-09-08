<#
	CSEC 475 - Windows Forensics
	Fall 2017

	Take in hostname from user, log into the box, get artifacts, output to screen and .csv file, send .csv file
	in an email (optional).
	
	This must be run from the command line with the following command:
		Invoke-Command -ComputerName [host's name] -FilePath [local file path]

#>

# Global Variables
$keep_looping = 1; 			# 1 == continue loop, 0 == end loop


# Do-While Loop
Do{

	# Get hostname and login credentials from user
	$remote_host = Read-Host -Prompt "Input the hostname of the remote host"
	$login_username = Read-Host -Prompt "Input a valid username for the remote host"
	
	#Try to establish a PSRemote session on the target host
	try{
	
		# Create new file to add all variables
			$file = New-Item c:\Users\$login_username\Desktop\lab01_output.txt -type file
	
		# Get PC OS info
			$os = Get-WmiObject win32_operatingsystem; Add-Content $file $os
		
		# Get time
			$current_time = (Get-Date).ToString('hh:mm:ss')
			$time_zone = Get-Timezone
			$uptime = (Get-Date) - ($os.ConvertToDateTime($os.lastbootuptime))

		# Get Windows version
			$os_version_name = $os.Caption
			$os_numerical_info = [environment]::OSVersion.Version

		# Get system hardware specs
			$cpu_info = Get-WmiObject –Class Win32_processor
			$ram_info = Get-WmiObject -Class "Win32_PhysicalMemoryArray"
			$hdd_info = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $remote_host

		# Get domain controller information
			$domain_info = Get-ADDomainController

		# Get hostname and domain
			$host_name = (Get-WmiObject Win32_ComputerSystem).Name
			$domain_name = (Get-WmiObject Win32_ComputerSystem).Domain

		# Get user info
			$local_users = Get-WmiObject -Class Win32_UserAccount -Namespace "root\cimv2"
			$domain_users = Get-ADUser -Filter * -Properties * | Get-Member -MemberType property

		# Get information about start-at-boot processes
			$boot_processes = Get-CimInstance Win32_StartupCommand | Select-Object Name, command, Location, User | Format-List

		# Get list of scheduled tasks
			$scheduled_tasks = GetScheduledTask
		
		# Get Network Info
			$arp_info = Get-NetNeighbor
			$mac_addrs = Get-WmiObject win32_networkadapterconfiguration | select description, macaddress
			$routing_table = Get-NetRoute
			$ip_4and6_addrs = Get-NetIPAddress
			$dns_servers = Get-DnsClientServerAddress | Select-Object –ExpandProperty ServerAddresses
			$listening_services_info = Get-NetTCPConnection -State Listen
			$established_connections = Get-NetTCPConnection -State Established
			$dns_cache = Get-DnsClientCache

		# Get share, printer, and wifi info
			$net_shares = Get-WmiObject -Class Win32_Share
			$printer_info = Get-WMIObject Win32_Printer

			$wifi_profiles = @() 
			$wifi_profiles += (netsh wlan show profiles)|Select-String "\:(.+)$" | Foreach{$_.Matches.Groups[1].Value.Trim()} |sort-object

		# Get list of installed software
			$software_list = Get-WmiObject -Class Win32_Product

		# Get process list
			$process_list = Get-Process

		# Get list of drivers installed - requires Admin
			$driver_list = Get-WindowsDriver -Online

		# Get list of files in Downloads and Documents folder
			$user_documents = Get-ChildItem -Path 'C:\Users\*\Documents' -Recurse
			$user_downloads = Get-ChildItem -Path 'C:\Users\*\Downloads' -Recurse

		# Original Artifact 1 - Get list of installed hardware
			$hardware_list = Get-PnpDevice -PresentOnly

		# Original Artifact 2 - BIOS Info
			$bios_info = Get-WmiObject -Class Win32_Bios

		# Original Artifact 3 - Get all email addresses in outlook inbox
			Clear-Host
			$Folder = "InBox"
			Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
			$Outlook = New-Object -ComObject Outlook.Application
			$Namespace = $Outlook.GetNameSpace("MAPI")
			$NameSpace.Folders.Item(1)
			$Email = $NameSpace.Folders.Item(1).Folders.Item($Folder).Items
			$Email | Sort-Object SenderEmailAddress -Unique | FT SenderEmailAddress

	}
	# If no connection can be made, tell user
	catch{
		Write-Output "Something went wrong...."
	}
	
	#Ask user if they want to send file as email
	$send_mail = Read-Host -Prompt "Enter 1 to send data in an email, 0 to skip"
	
	If($send_mail -eq 1){
		
		Add-PSSnapin Microsoft.Exchange.Management.Powershell.Admin -erroraction silentlyContinue
		
		$email_from = Read-Host -Prompt "Enter the email address of the sender"
		$email_to = Read-Host -Prompt "Enter the email address of the receiver"
		$email_subject = Read-Host -Prompt "Enter the email subject line"
		
		$smtpServer = "127.0.0.1"
		$att = new-object Net.Mail.Attachment($file)
		$msg = new-object Net.Mail.MailMessage
		$smtp = new-object Net.Mail.SmtpClient($smtpServer)
		$msg.From = $email_from
		$msg.Subject = $email_subject
		$msg.Body = "Attached is the host data acquired via the script from lab 01"
		$msg.Attachments.Add($att)
		$smtp.Send($msg)
		$att.Dispose()
		
	}
	
	#Ask user if they want to keep going
	$keep_looping = Read-Host -Prompt "Enter 1 to try another host or 0 to quit"
	
} While ($keep_looping -eq 1)
