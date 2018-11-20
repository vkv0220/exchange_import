# Подгружаем набор командлетов
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
Import-Module ActiveDirectory
# Описываем массив для данных
$ListContacts = @();
# выгружаем список почтовых ящиков из конкретной OU или всего домена (можно также указать дополнительные фильтры при необходимости)
$ListEmplMailbox = Get-Mailbox #|-OrganizationalUnit "domain.name/OE/Contacts"
# Формируем массив ящиков с необходимыми полями
Foreach ($mailbox in $ListEmplMailbox ){
	# Запрашиваем атрибуты из учетки в AD
	$User = $mailbox.SamAccountName | Get-ADUser -Properties *
	$Contact = new-object -TypeName PSObject
	#ФИО
	$Contact | Add-Member -MemberType NoteProperty -Name FIO -Value $User.DisplayName
	#Alias как правило совпадает с SamAccountName
	$Contact | Add-Member -MemberType NoteProperty -Name Alias -Value $mailbox.Alias
	#email
	$Contact | Add-Member -MemberType NoteProperty -Name Email -Value $User.EmailAddress
	#должность
	$Contact | Add-Member -MemberType NoteProperty -Name Title -Value $User.Title
	#Телефон
	$Contact | Add-Member -MemberType NoteProperty -Name Phone -Value $User.telephoneNumber
	#Компания
	$Contact | Add-Member -MemberType NoteProperty -Name Company -Value "CompanyName"
	$ListContacts += $Contact
}
# Выгружаем массив в файл
$ListContacts | Export-Csv -Path "C:\Temp\ListContact_domain2.csv" -Encoding UTF8 -NoTypeInformation -Delimiter ";"