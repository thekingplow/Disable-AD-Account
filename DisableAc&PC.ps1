#v2.3 (12.07.2025)
#Developed by Danilovich M.D.


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices


# Проверка запуска с правами администратора
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $currentScript = $MyInvocation.MyCommand.Definition
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$currentScript`"" -Verb RunAs
    exit
}








#USER FUNCTIONS

function Get-ADUsersFromOU {
    try {
        $ou = "OU=ExampleOU,DC=example,DC=com"  #Указать путь к вашей OU
        $users = Get-ADUser -Filter * -SearchBase $ou -Property DisplayName | Select-Object -ExpandProperty DisplayName | Sort-Object
        return $users
    } catch {
        Write-Host "Ошибка при получении пользователей из Active Directory: $_"
        return @()
    }
}



function Remove-ADUserProperties {
    param(
        [string]$username
    )
    
    try {
        $user = Get-ADUser -Filter { DisplayName -eq $username } -Properties company, title, department, manager
        if ($user) {
            Set-ADUser -Identity $user -Clear company, title, department, manager
            return "Данные пользователя '$username' были удалены."
        } else {
            return "Пользователь '$username' не найден."
        }
    } catch {
        return "Ошибка при удалении данных: $_"
    }
}



function Add-UserToDomainGuests {
    param(
        [string]$username
    )
    
    try {
        $user = Get-ADUser -Filter { DisplayName -eq $username } -Properties DistinguishedName
        if ($user) {
            $group = "Domain Guests"
            $groupDN = (Get-ADGroup -Filter { Name -eq $group }).DistinguishedName
            Add-ADGroupMember -Identity $groupDN -Members $user.DistinguishedName
            return "Пользователь '$username' добавлен в группу '$group'."
        } else {
            return "Пользователь '$username' не найден."
        }
        } catch {
        return "Ошибка при добавлении пользователя в группу: $_"
    }
}



function Set-UserPrimaryGroup {
    param(
        [string]$username
    )
    
        try {
             $user = Get-ADUser -Filter { DisplayName -eq $username } -Properties DistinguishedName
            if ($user) {
#В Windows у каждого пользователя есть одна основная группа - PrimaryGroupID. По умолчанию это Domain Users (RID 513)
                $primaryGroupID = 514  #514 - RID группы Domain Guests
                    if ($null -ne $primaryGroupID) {
                        Set-ADUser -Identity $user.DistinguishedName -Replace @{PrimaryGroupID = $primaryGroupID}
                        return "Основная группа пользователя '$username' изменена на 'Domain Guests'."
                    } else {
                        return "Не удалось получить PrimaryGroupID для группы 'Domain Guests'."
                            }
                    } else {
                        return "Пользователь '$username' не найден."
                }           
         } catch {
        return "Ошибка при изменении основной группы пользователя: $_"
    }
}



function Remove-AllUserGroups {
    param(
        [string]$username
    )
    
    try {
        $user = Get-ADUser -Filter { DisplayName -eq $username } -Properties DistinguishedName, PrimaryGroupID
        if ($user) {
            $primaryGroupID = $user.PrimaryGroupID
            $groups = Get-ADUser -Identity $user.DistinguishedName -Properties MemberOf | Select-Object -ExpandProperty MemberOf
            foreach ($groupDN in $groups) {
                $group = Get-ADGroup -Identity $groupDN -Properties PrimaryGroupID
                if ($group.PrimaryGroupID -ne $primaryGroupID) {
                    Remove-ADGroupMember -Identity $groupDN -Members $user.DistinguishedName -Confirm:$false
                }
            }
            return "Все группы пользователя '$username', кроме основной, были удалены."
        } else {
            return "Пользователь '$username' не найден."
        }
    } catch {
        return "Ошибка при удалении групп пользователя: $_"
    }
}



function Disable-UserAccount {
    param(
        [string]$username
    )
    
    try {
        $user = Get-ADUser -Filter { DisplayName -eq $username } -Properties Enabled
        if ($user) {
            Disable-ADAccount -Identity $user.DistinguishedName
            return "Учетная запись пользователя '$username' была отключена."
        } else {
            return "Пользователь '$username' не найден."
        }
    } catch {
        return "Ошибка при отключении учетной записи: $_"
    }
}



function Move-UserToOU {
    param(
        [string]$username,
        [string]$ouDN
    )
    
    try {
        $user = Get-ADUser -Filter { DisplayName -eq $username } -Properties DistinguishedName
        if ($user) {
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $ouDN
            return "Пользователь '$username' был перемещен в OU '$ouDN'."
        } else {
            return "Пользователь '$username' не найден."
        }
    } catch {
        return "Ошибка при перемещении пользователя: $_"
    }
}



# Функция для скрытия почты пользователя на Exchange сервере
function Hide-MailboxFromAddressLists {
    param(
        [string]$username
    )
    
    try {
        # Проверка запуска с правами администратора
        if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            $currentScript = $MyInvocation.MyCommand.Definition
            Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$currentScript`"" -Verb RunAs
            exit
        }

        # Подключение к Exchange
        try {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://YourExchangeServer-XX/PowerShell/ -Authentication Kerberos
            Import-PSSession $Session -DisableNameChecking -AllowClobber > $null 2>&1
            Write-Host "Подключение к Exchange успешно установлено." -ForegroundColor Green
        } catch {
            return "Ошибка подключения к Exchange: $($_.Exception.Message)"
        }

        # Получение пользователя
        $user = Get-ADUser -Filter { DisplayName -eq $username } -Properties Mail
            if ($user) {
            $mailbox = Get-Mailbox -Identity $user.Mail
            if ($mailbox) {
                Set-Mailbox -Identity $mailbox.Alias -HiddenFromAddressListsEnabled $true > $null 2>&1
                return "Почтовый ящик пользователя '$username' скрыт из адресных списков."
            } else {
                return "Почтовый ящик пользователя '$username' не найден."
            }
        } else {
            return "Пользователь '$username' не найден."
        }
    } catch {
        return "Ошибка при скрытии почтового ящика: $_"
    }
}





#COMPUTERS FUNCTIONS

function Get-ComputersFromOU {
    param(
        [string]$ouDN = "OU=<Department>,OU=<Unit>,OU=<MainOU>,DC=<DomainName>,DC=local"
    )

    try {
        $computers = Get-ADComputer -Filter * -SearchBase $ouDN -Property Name | Select-Object -ExpandProperty Name | Sort-Object
        return $computers
    } catch {
        Write-Host "Ошибка при получении компьютеров из OU: $_"
        return @()
    }
}



function Remove-ComputerFromAD {
    param (
        [string]$computerName
    )

    try {
        $adComputer = Get-ADComputer -Identity $computerName -ErrorAction Stop
        Remove-ADComputer -Identity $adComputer -Confirm:$false
        [System.Windows.Forms.MessageBox]::Show("Компьютер '$computerName' успешно удален из Active Directory.", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Обновление списка компьютеров
        Update-ComputerList
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при удалении компьютера: $_", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



function Remove-ComputerDNSRecord {
    param (
        [string]$computerName,
        [string]$dnsZone = "domain.local",
        [string]$dnsServer = "nameserver-xx"
    )

    try {
        Remove-DnsServerResourceRecord -ZoneName $dnsZone -Name $computerName -RRType "A" -ComputerName $dnsServer -Force
        [System.Windows.Forms.MessageBox]::Show("DNS-запись компьютера '$computerName' успешно удалена.", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при удалении DNS-записи: $_", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



Import-Module ConfigurationManager
$memSiteLocation = "CompanyName:"
Set-Location $memSiteLocation

function Remove-ComputerFromMEM {
    param(
        [string]$computerName
    )

    try {
        $device = Get-CMDevice | Where-Object { $_.Name -eq $computerName }

        if ($device) {
            Remove-CMResource -ResourceId $device.ResourceID -Force
            [System.Windows.Forms.MessageBox]::Show("Компьютер '$computerName' успешно удален из MEM (SCCM).", "Успех", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Компьютер '$computerName' не найден в MEM (SCCM).", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при удалении компьютера из MEM: $_", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



function OpenSCOM {
    try {
        $scomPath = "C:\Program Files\System Center Operations Manager\Console\Microsoft.EnterpriseManagement.Monitoring.Console.exe"
        Start-Process $scomPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при запуске SCOM: $_", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



# Функция для обновления списка компьютеров в ListBox
function Update-ComputerList {
    $computers = Get-ComputersFromOU
    $listBoxComputers.Items.Clear()
    $listBoxComputers.Items.AddRange($computers)
}















# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Disable account user's and pc"
$form.Size = New-Object Drawing.Size(920, 650) 
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle 
$form.MaximizeBox = $false 
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Arial", 13, [System.Drawing.FontStyle]::Bold)

$scriptPath = $PSScriptRoot

# Установка иконки
$iconPath = Join-Path -Path $scriptPath -ChildPath "images\ac_off.ico" # Укажите путь к вашей иконке
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)

$imagePath = Join-Path -Path $scriptPath -ChildPath "images\bg.jpg"
$image = [System.Drawing.Image]::FromFile($imagePath)

# Устанавливаем изображение как фон формы
$form.BackgroundImage = $image
$form.BackgroundImageLayout = "Stretch"  # Растягиваем изображение на всю форму



$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "v2.3 (12.07.2025)"
$labelVersion.Location = New-Object System.Drawing.Point(0, 0)
$labelVersion.Font = New-Object System.Drawing.Font("Arial", 7.5, [System.Drawing.FontStyle]::Bold)  # Увеличение размера шрифта и жирный шрифт
$labelVersion.AutoSize = $true  # Автоматический размер под текст
$labelVersion.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelVersion.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelVersion)



# Создание заголовка
$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Text = "DISABLE ACCOUNT AND PC"
$labelTitle.Location = New-Object System.Drawing.Point(170, 40)
$labelTitle.Font = New-Object System.Drawing.Font("Arial", 30, [System.Drawing.FontStyle]::Bold)  # Увеличение размера шрифта и жирный шрифт
$labelTitle.AutoSize = $true  # Автоматический размер под текст
$labelTitle.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelTitle.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelTitle)



$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "ПОЛЬЗОВАТЕЛЬ:"
$labelUser.Top = 150
$labelUser.Left = 120
$labelUser.BackColor = [System.Drawing.Color]::Transparent
$labelUser.AutoSize = $true
$labelUser.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelUser)



$textBoxUser = New-Object System.Windows.Forms.TextBox
$textBoxUser.Top = 190
$textBoxUser.Left = 120
$textBoxUser.Width = 265
$form.Controls.Add($textBoxUser)

$autoComplete = New-Object System.Windows.Forms.AutoCompleteStringCollection
$users = Get-ADUsersFromOU
$autoComplete.AddRange($users)

$textBoxUser.AutoCompleteCustomSource = $autoComplete
$textBoxUser.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$textBoxUser.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource


$textBoxUser.Add_KeyPress({
    param($sender, $e)
    # Разрешаем только буквы (латиница и кириллица) и управляющие символы  Backspace
    if (-not ([char]::IsLetter($e.KeyChar) -or [char]::IsControl($e.KeyChar) -or $e.KeyChar -eq ' ')) {
        $e.Handled = $true
    }
})








# Создание единой кнопки для выполнения всех действий
$buttonDisableUser = New-Object System.Windows.Forms.Button
$buttonDisableUser.Text = "ОТКЛЮЧИТЬ АККАУНТ"
$buttonDisableUser.Top = 240
$buttonDisableUser.Left = 120
$buttonDisableUser.Width = 265
$buttonDisableUser.Height = 38
#$buttonDisableUser.BackColor = [System.Drawing.Color]::Silver  # Устанавливаем цвет фона кнопки
$buttonDisableUser.ForeColor = [System.Drawing.Color]::Red			      # Устанавливаем цвет текста кнопки
$buttonDisableUser.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю
$buttonDisableUser.Padding = New-Object System.Windows.Forms.Padding(5, 0, 7.5, 0)  # Сдвигаем текст влево на 5 пикселей

$buttonDisableUser.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonDisableUser.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonDisableUser.FlatAppearance.BorderColor = [System.Drawing.Color]::Brown
$buttonDisableUser.FlatAppearance.BorderSize = 2

$buttonDisableUser.Font = New-Object System.Drawing.Font("Arial", 13.5, [System.Drawing.FontStyle]::Bold)

# Определение относительного пути к иконке
$dsbuserPath = Join-Path -Path $scriptPath -ChildPath "images/disb_ac.png"

# Загрузка и установка иконки для кнопки
$disb = [System.Drawing.Image]::FromFile($dsbuserPath)
$buttonDisableUser.Image = $disb

# Устанавливаем выравнивание иконки
$buttonDisableUser.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
#$buttonDisableUser.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)

$form.Controls.Add($buttonDisableUser)





# Создание новой кнопки на форме
$buttonHideMailbox = New-Object System.Windows.Forms.Button
$buttonHideMailbox.Text = "СКРЫТЬ ПОЧТУ ИЗ АДРЕСНОЙ КНИГИ"
$buttonHideMailbox.Top = 300
$buttonHideMailbox.Left = 120
$buttonHideMailbox.Width = 265
$buttonHideMailbox.Height = 38
$buttonHideMailbox.ForeColor = [System.Drawing.Color]::Blue      # Устанавливаем цвет текста кнопки
$buttonHideMailbox.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

$buttonHideMailbox.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonHideMailbox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonHideMailbox.FlatAppearance.BorderColor = [System.Drawing.Color]::DodgerBlue	
$buttonHideMailbox.FlatAppearance.BorderSize = 2

$buttonHideMailbox.Font = New-Object System.Drawing.Font("Arial", 8.5, [System.Drawing.FontStyle]::Bold)


# Определение относительного пути к иконке
$hidemailPath = Join-Path -Path $scriptPath -ChildPath "images\disb_m.png"

# Загрузка и установка иконки для кнопки
$mail = [System.Drawing.Image]::FromFile($hidemailPath)
$buttonHideMailbox.Image = $mail

# Устанавливаем выравнивание иконки
$buttonHideMailbox.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonHideMailbox.Padding = New-Object System.Windows.Forms.Padding(5, 0, 2, 0)



$form.Controls.Add($buttonHideMailbox)





$labelComputer = New-Object System.Windows.Forms.Label
$labelComputer.Text = "КОМПЬЮТЕР:"
$labelComputer.Top = 150
$labelComputer.Left = 550
$labelComputer.BackColor = [System.Drawing.Color]::Transparent
$labelComputer.AutoSize = $true
$labelComputer.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelComputer)

$textBoxComputer = New-Object System.Windows.Forms.TextBox
$textBoxComputer.Top = 445
$textBoxComputer.Left = 550
$textBoxComputer.Width = 220
$form.Controls.Add($textBoxComputer)

$listBoxComputers = New-Object System.Windows.Forms.ListBox
$listBoxComputers.Top = 190
$listBoxComputers.Left = 550
$listBoxComputers.Width = 220
$listBoxComputers.Height = 250
$form.Controls.Add($listBoxComputers)

# Изначальная загрузка списка компьютеров
$computers = Get-ComputersFromOU
$listBoxComputers.Items.AddRange($computers)

$listBoxComputers.Add_SelectedIndexChanged({
    $selectedComputer = $listBoxComputers.SelectedItem
    $textBoxComputer.Text = $selectedComputer
})





# Кнопка для выполнения всех действий
$buttonDeletePC = New-Object System.Windows.Forms.Button
$buttonDeletePC.Text = "УДАЛИТЬ КОМПЬЮТЕР"
$buttonDeletePC.Top = 490
$buttonDeletePC.Left = 550
$buttonDeletePC.Width = 220
$buttonDeletePC.Height = 40
$buttonDeletePC.ForeColor = [System.Drawing.Color]::red      # Устанавливаем цвет текста кнопки
$buttonDeletePC.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

$buttonDeletePC.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonDeletePC.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonDeletePC.FlatAppearance.BorderColor = [System.Drawing.Color]::Brown
$buttonDeletePC.FlatAppearance.BorderSize = 2

$buttonDeletePC.Font = New-Object System.Drawing.Font("Arial", 10.8, [System.Drawing.FontStyle]::Bold)


# Определение относительного пути к иконке
$deletepcPath = Join-Path -Path $scriptPath -ChildPath "images\del_pc.png"

# Загрузка и установка иконки для кнопки
$deletepc = [System.Drawing.Image]::FromFile($deletepcPath)
$buttonDeletePC.Image = $deletepc

# Устанавливаем выравнивание иконки
$buttonDeletePC.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonDeletePC.Padding = New-Object System.Windows.Forms.Padding(5, 0, 6, 0)

$form.Controls.Add($buttonDeletePC)






# Создание элемента LinkLabel
$linkLabel = New-Object System.Windows.Forms.LinkLabel
$linkLabel.Text = "УЧЕТ ОБОРУДОВАНИЯ"
$linkLabel.Top = 450
$linkLabel.Left = 155
$linkLabel.AutoSize = $true
$linkLabel.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$form.Controls.Add($linkLabel)


# Убираем нижнее подчеркивание и устанавливаем цвет ссылки
$linkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::NeverUnderline
$linkLabel.LinkColor = [System.Drawing.Color]::Yellow
$linkLabel.VisitedLinkColor = [System.Drawing.Color]::White
$linkLabel.ActiveLinkColor = [System.Drawing.Color]::White

# Установка ссылки
$link = "https://wiki.company.by/pages/viewpage.action?pageId=17760294"
$linkLabel.Links.Add(0, $linkLabel.Text.Length, $link)

# Путь к Google Chrome (обновите путь, если у вас он другой)
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

# Добавление обработчика события для открытия ссылки в Google Chrome
$linkLabel.add_LinkClicked({
    param ($sender, $e)
    & $chromePath $e.Link.LinkData.ToString()
})


# Сохраняем цвет по умолчанию
$defaultColor = [System.Drawing.Color]::Yellow
$hoverColor = [System.Drawing.Color]::LightGreen

# Наведение — меняем цвет
$linkLabel.Add_MouseEnter({
    $linkLabel.LinkColor = $hoverColor
})

# Уход мыши — возвращаем цвет
$linkLabel.Add_MouseLeave({
    $linkLabel.LinkColor = $defaultColor
})










 

$buttonDisableUser.Add_Click({
    $username = $textBoxUser.Text
    if ([string]::IsNullOrWhiteSpace($username)) {
         [System.Windows.Forms.MessageBox]::Show(
            "Введите имя пользователя.", 
            "Ошибка", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # Подтверждение перед деактивацией
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "Вы уверены, что хотите деактивировать пользователя '$username'?", 
        "Подтверждение", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    # Если пользователь нажал "Нет" — отмена операции
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    # Выполнение всех функций и отображение сообщений поочередно
    $messages = @()
    $messages += Remove-ADUserProperties -username $username
    $messages += Add-UserToDomainGuests -username $username
    $messages += Set-UserPrimaryGroup -username $username
    $messages += Remove-AllUserGroups -username $username
    $messages += Disable-UserAccount -username $username
    $ouDN = "OU=ExampleOU,DC=example,DC=com"  #Указать путь к OU для перемещения
    $messages += Move-UserToOU -username $username -ouDN $ouDN

    foreach ($message in $messages) {
            [System.Windows.Forms.MessageBox]::Show(
        $message, 
        "Результат", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    }
})








# Привязка функции к кнопке
$buttonHideMailbox.Add_Click({
    $username = $textBoxUser.Text
    if ([string]::IsNullOrWhiteSpace($username)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Введите имя пользователя.", 
            "Ошибка", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    
    # Выполнение функции скрытия почтового ящика из адресных списков
    $message = Hide-MailboxFromAddressLists -username $username
    
    # Показ сообщения с результатом и иконкой информации
    [System.Windows.Forms.MessageBox]::Show(
        $message, 
        "Результат", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})








$buttonDeletePC.Add_Click({
    $computerName = $textBoxComputer.Text

    if (-not [string]::IsNullOrEmpty($computerName)) {
        # Подтверждение для удаления из Active Directory
        $adResult = [System.Windows.Forms.MessageBox]::Show("Вы уверены, что хотите удалить компьютер '$computerName' из Active Directory?", "Подтверждение", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($adResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-ComputerFromAD -computerName $computerName
        }

        # Подтверждение для удаления DNS-записи
        $dnsResult = [System.Windows.Forms.MessageBox]::Show("Вы уверены, что хотите удалить DNS-запись компьютера '$computerName'?", "Подтверждение", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($dnsResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-ComputerDNSRecord -computerName $computerName
        }

        # Подтверждение для удаления из MEM (SCCM)
        $memResult = [System.Windows.Forms.MessageBox]::Show("Вы уверены, что хотите удалить компьютер '$computerName'?' из MEM (SCCM)?", "Подтверждение", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($memResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-ComputerFromMEM -computerName $computerName
        }

        # Подтверждение для запуска SCOM
        $scomResult = [System.Windows.Forms.MessageBox]::Show("Вы уверены, что хотите удалить '$computerName'?' из SCOM?", "Подтверждение", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($scomResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            OpenSCOM
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Выберите компьютер.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})





# Show the form
[void]$form.ShowDialog()
