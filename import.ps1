# Подгружаем набор командлетов
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
# Указываем файл из которого брать информацию о контактах
Set-Location "C:\Obmen"
$ListContacts = Import-Csv .\ListContact_domain2.csv -Delimiter ";" #| select -First 40
# Загружаем имеющиеся в AD контакты
$ListExistContacts = Get-Contact
# начинаем создание контактов
foreach ($Contact in $ListContacts)
{
	$Contact_Exist = $null
	# Проверяем существует ли такой контакт
	foreach ($ExistContact in $ListExistContacts)
	{
		if ($Contact.Mail -eq $ExistContact.WindowsEmailAddress) {$Contact_Exist = $true }
	}
	if ($Contact_Exist) {"Контакт существует: " + $Contact.Name }
	else
	{
		# Если такого контакта нет, то создаем его
		# парсим строку с ФИО на массив Фамилия Имя Отчетство
		# $fio[0] - Фамилия
		# $fio[1] - Имя
		# $fio[2] - Отчество
		$fio = $Contact.FIO -split " "
		$io = $fio[1] + " " + $fio[2]
		# Дописываем префикс, чтобы не попасть в имеющиеся алиасы
		$MailNickName = $Contact.Alias + "_contact"
		# Создаем контакт
		$newcontact = New-MailContact -Name $Contact.FIO -DisplayName $Contact.FIO -LastName $io -FirstName $fio[0] -Alias $MailNickName -OrganizationalUnit "domain.name/OE/Contacts" -ExternalEmailAddress $Contact.Email
		# Дописываем в созданном контакте компанию должность и телефон
		Set-Contact -Identity $newcontact.Identity -Company $Contact.Company -Title $Contact.Title -Phone $Contact.Phone
		$newcontact.Identity
	}
}