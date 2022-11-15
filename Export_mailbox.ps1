Write-Host $args[0].GetType().Name;
if($args[0].GetType().Name -like "String") {
    Write-Host "Получено имя почтового ящика:" $args[0];
    Write-Host "Подгружаем набор командлетов";
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue;
    Import-Module ActiveDirectory;

    if(Get-Mailbox $args[0]) {
        Write-Host "Почтовый ящик найден";
        Get-Mailbox $args[0];
        Write-Host "Экспорт почтового ящика в расположение по умолчанию: \\backup-server\c$\Backup\$($args[0]).pst";
        $filepath = "\\backup-server\c$\Backup\$($args[0]).pst";
        New-MailboxExportRequest -Mailbox $args[0] -Filepath $filepath;
    }else {
        Write-Host "Почтовый ящик:" $args[0] "не найден, проверьте вводимые данные";
    }
}else {
   Write-Host "Неверный тип данных, ожидается строка:" $args[0];
}

