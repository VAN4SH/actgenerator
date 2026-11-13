#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ScriptDir 'Common.ps1')

$script:Settings = Load-Settings
$script:Products = New-Object System.Collections.Generic.List[object]
$script:Sync = $false

function Msg([string]$text, [string]$title = 'Акты', [string]$icon = 'Information') {
  [void][System.Windows.Forms.MessageBox]::Show($text, $title, 'OK', $icon)
}

# =====================================================================
#  Настройки путей (первый запуск и кнопка «Настройки»)
# =====================================================================
function Show-SettingsDialog([bool]$first) {
  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Настройки: папки'
  $dlg.StartPosition = 'CenterScreen'
  $dlg.Size = New-Object System.Drawing.Size(720, 330)
  $dlg.FormBorderStyle = 'FixedDialog'
  $dlg.MaximizeBox = $false
  $dlg.MinimizeBox = $false
  $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

  $info = New-Object System.Windows.Forms.Label
  $info.Location = New-Object System.Drawing.Point(12, 10)
  $info.Size = New-Object System.Drawing.Size(680, 52)
  if ($first) {
    $info.Text = "Первый запуск. Укажите папку с шаблонами и папку, куда сохранять готовые акты.`nПрограмма уже попробовала найти их рядом с собой."
  } else {
    $info.Text = 'Папка с шаблонами и папка для готовых актов.'
  }
  $dlg.Controls.Add($info)

  $l1 = New-Object System.Windows.Forms.Label
  $l1.Text = 'Папка с шаблонами (Шаблоны):'
  $l1.Location = New-Object System.Drawing.Point(12, 68)
  $l1.Size = New-Object System.Drawing.Size(400, 20)
  $dlg.Controls.Add($l1)
  $tbT = New-Object System.Windows.Forms.TextBox
  $tbT.Location = New-Object System.Drawing.Point(12, 90)
  $tbT.Size = New-Object System.Drawing.Size(580, 24)
  $dlg.Controls.Add($tbT)
  $bT = New-Object System.Windows.Forms.Button
  $bT.Text = 'Обзор...'
  $bT.Location = New-Object System.Drawing.Point(600, 88)
  $bT.Size = New-Object System.Drawing.Size(90, 28)
  $dlg.Controls.Add($bT)

  $l2 = New-Object System.Windows.Forms.Label
  $l2.Text = 'Папка для готовых актов (Готовые):'
  $l2.Location = New-Object System.Drawing.Point(12, 128)
  $l2.Size = New-Object System.Drawing.Size(400, 20)
  $dlg.Controls.Add($l2)
  $tbO = New-Object System.Windows.Forms.TextBox
  $tbO.Location = New-Object System.Drawing.Point(12, 150)
  $tbO.Size = New-Object System.Drawing.Size(580, 24)
  $dlg.Controls.Add($tbO)
  $bO = New-Object System.Windows.Forms.Button
  $bO.Text = 'Обзор...'
  $bO.Location = New-Object System.Drawing.Point(600, 148)
  $bO.Size = New-Object System.Drawing.Size(90, 28)
  $dlg.Controls.Add($bO)

  $lblState = New-Object System.Windows.Forms.Label
  $lblState.Location = New-Object System.Drawing.Point(12, 186)
  $lblState.Size = New-Object System.Drawing.Size(680, 50)
  $lblState.ForeColor = [System.Drawing.Color]::DarkRed
  $dlg.Controls.Add($lblState)

  $bAuto = New-Object System.Windows.Forms.Button
  $bAuto.Text = 'Найти автоматически'
  $bAuto.Location = New-Object System.Drawing.Point(12, 245)
  $bAuto.Size = New-Object System.Drawing.Size(170, 32)
  $dlg.Controls.Add($bAuto)
  $bOk = New-Object System.Windows.Forms.Button
  $bOk.Text = 'Сохранить'
  $bOk.Location = New-Object System.Drawing.Point(500, 245)
  $bOk.Size = New-Object System.Drawing.Size(100, 32)
  $dlg.Controls.Add($bOk)
  $bCancel = New-Object System.Windows.Forms.Button
  $bCancel.Text = 'Отмена'
  $bCancel.Location = New-Object System.Drawing.Point(606, 245)
  $bCancel.Size = New-Object System.Drawing.Size(84, 32)
  $dlg.Controls.Add($bCancel)
  $dlg.CancelButton = $bCancel

  $tbT.Text = [string]$script:Settings.TemplatesDir
  $tbO.Text = [string]$script:Settings.OutDir
  if ([string]::IsNullOrWhiteSpace($tbT.Text)) {
    $f = Find-DefaultFolders
    $tbT.Text = $f.Templates
    if ([string]::IsNullOrWhiteSpace($tbO.Text)) { $tbO.Text = $f.Out }
  }

  $checkState = {
    $miss = @(Get-MissingTemplates $tbT.Text.Trim())
    if (-not (Test-Path -LiteralPath $tbT.Text.Trim())) {
      $lblState.Text = 'Папка шаблонов не найдена.'
    } elseif ($miss.Count -gt 0) {
      $lblState.Text = "В папке шаблонов не хватает файлов:`n" + ($miss -join ', ')
    } else {
      $lblState.Text = ''
    }
  }

  $bT.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = 'Папка с шаблонами'
    if ($tbT.Text -and (Test-Path -LiteralPath $tbT.Text)) { $fb.SelectedPath = $tbT.Text }
    if ($fb.ShowDialog() -eq 'OK') {
      $tbT.Text = $fb.SelectedPath
      if ([string]::IsNullOrWhiteSpace($tbO.Text)) { $tbO.Text = Join-Path (Split-Path -Parent $fb.SelectedPath) 'Готовые' }
      & $checkState
    }
  })
  $bO.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = 'Папка для готовых актов'
    if ($tbO.Text -and (Test-Path -LiteralPath $tbO.Text)) { $fb.SelectedPath = $tbO.Text }
    if ($fb.ShowDialog() -eq 'OK') { $tbO.Text = $fb.SelectedPath }
  })
  $bAuto.Add_Click({
    $f = Find-DefaultFolders
    if ($f.Templates) {
      $tbT.Text = $f.Templates
      $tbO.Text = $f.Out
    } else {
      Msg 'Папку «Шаблоны» автоматически найти не удалось. Укажите вручную.' 'Поиск' 'Warning'
    }
    & $checkState
  })
  $tbT.Add_TextChanged({ & $checkState })

  $script:SettingsSaved = $false
  $bOk.Add_Click({
    $t = $tbT.Text.Trim()
    $o = $tbO.Text.Trim()
    if (-not (Test-Path -LiteralPath $t)) { Msg 'Укажите существующую папку с шаблонами.' 'Настройки' 'Warning'; return }
    if ($o -eq '') { Msg 'Укажите папку для готовых актов.' 'Настройки' 'Warning'; return }
    try {
      if (-not (Test-Path -LiteralPath $o)) { New-Item -ItemType Directory -Force -Path $o | Out-Null }
    } catch { Msg ("Не удалось создать папку:`n" + $_.Exception.Message) 'Настройки' 'Error'; return }
    $miss = @(Get-MissingTemplates $t)
    if ($miss.Count -gt 0) {
      $a = [System.Windows.Forms.MessageBox]::Show(("В папке шаблонов не хватает файлов:`n{0}`n`nСохранить всё равно?" -f ($miss -join "`n")), 'Настройки', 'YesNo', 'Warning')
      if ($a -ne 'Yes') { return }
    }
    $script:Settings.TemplatesDir = $t
    $script:Settings.OutDir = $o
    Save-Settings $script:Settings
    $script:SettingsSaved = $true
    $dlg.Close()
  })
  $bCancel.Add_Click({ $dlg.Close() })

  & $checkState
  [void]$dlg.ShowDialog()
  return $script:SettingsSaved
}

# =====================================================================
#  Справочник сотрудников
# =====================================================================
function Show-EngineersDialog {
  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Сотрудники (исполнители)'
  $dlg.StartPosition = 'CenterParent'
  $dlg.Size = New-Object System.Drawing.Size(780, 420)
  $dlg.FormBorderStyle = 'FixedDialog'
  $dlg.MaximizeBox = $false
  $dlg.MinimizeBox = $false
  $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

  $lst = New-Object System.Windows.Forms.ListBox
  $lst.Location = New-Object System.Drawing.Point(12, 12)
  $lst.Size = New-Object System.Drawing.Size(210, 300)
  $dlg.Controls.Add($lst)

  $y = 14
  $boxes = @{}
  foreach ($def in @(
    @('Name', 'Как показывать в списке (напр. Скворцов П.А.)'),
    @('Init', 'Подпись в акте (напр. П.А. Скворцов)'),
    @('Dolzh', 'Должность (как в акте)'),
    @('Fio', 'Полное ФИО (нужно для акта испытаний)')
  )) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $def[1]
    $l.Location = New-Object System.Drawing.Point(240, $y)
    $l.Size = New-Object System.Drawing.Size(500, 18)
    $dlg.Controls.Add($l)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point(240, ($y + 20))
    $t.Size = New-Object System.Drawing.Size(510, 24)
    $dlg.Controls.Add($t)
    $boxes[$def[0]] = $t
    $y += 56
  }

  $bAdd = New-Object System.Windows.Forms.Button
  $bAdd.Text = 'Добавить как нового'
  $bAdd.Location = New-Object System.Drawing.Point(240, 244)
  $bAdd.Size = New-Object System.Drawing.Size(150, 30)
  $dlg.Controls.Add($bAdd)
  $bSave = New-Object System.Windows.Forms.Button
  $bSave.Text = 'Сохранить изменения'
  $bSave.Location = New-Object System.Drawing.Point(396, 244)
  $bSave.Size = New-Object System.Drawing.Size(160, 30)
  $dlg.Controls.Add($bSave)
  $bDel = New-Object System.Windows.Forms.Button
  $bDel.Text = 'Удалить'
  $bDel.Location = New-Object System.Drawing.Point(562, 244)
  $bDel.Size = New-Object System.Drawing.Size(100, 30)
  $dlg.Controls.Add($bDel)
  $bClose = New-Object System.Windows.Forms.Button
  $bClose.Text = 'Готово'
  $bClose.Location = New-Object System.Drawing.Point(650, 340)
  $bClose.Size = New-Object System.Drawing.Size(100, 32)
  $dlg.Controls.Add($bClose)

  $script:EngList = New-Object System.Collections.Generic.List[object]
  foreach ($e in @($script:Settings.Engineers)) {
    $script:EngList.Add([pscustomobject]@{ Name = [string]$e.Name; Fio = [string]$e.Fio; Dolzh = [string]$e.Dolzh; Init = [string]$e.Init })
  }
  $refresh = {
    param([int]$sel)
    $lst.Items.Clear()
    foreach ($e in $script:EngList) { [void]$lst.Items.Add($e.Name) }
    if ($sel -ge 0 -and $sel -lt $lst.Items.Count) { $lst.SelectedIndex = $sel }
  }
  $lst.Add_SelectedIndexChanged({
    if ($lst.SelectedIndex -ge 0) {
      $e = $script:EngList[$lst.SelectedIndex]
      $boxes['Name'].Text = $e.Name; $boxes['Init'].Text = $e.Init
      $boxes['Dolzh'].Text = $e.Dolzh; $boxes['Fio'].Text = $e.Fio
    }
  })
  $bAdd.Add_Click({
    if ($boxes['Name'].Text.Trim() -eq '' -or $boxes['Init'].Text.Trim() -eq '' -or $boxes['Dolzh'].Text.Trim() -eq '') {
      Msg 'Заполните имя в списке, подпись и должность.' 'Сотрудники' 'Warning'; return
    }
    $script:EngList.Add([pscustomobject]@{ Name = $boxes['Name'].Text.Trim(); Fio = $boxes['Fio'].Text.Trim(); Dolzh = $boxes['Dolzh'].Text.Trim(); Init = $boxes['Init'].Text.Trim() })
    & $refresh ($script:EngList.Count - 1)
  })
  $bSave.Add_Click({
    $i = $lst.SelectedIndex
    if ($i -lt 0) { return }
    $script:EngList[$i] = [pscustomobject]@{ Name = $boxes['Name'].Text.Trim(); Fio = $boxes['Fio'].Text.Trim(); Dolzh = $boxes['Dolzh'].Text.Trim(); Init = $boxes['Init'].Text.Trim() }
    & $refresh $i
  })
  $bDel.Add_Click({
    $i = $lst.SelectedIndex
    if ($i -lt 0) { return }
    if ($script:EngList.Count -le 1) { Msg 'Должен остаться хотя бы один сотрудник.' 'Сотрудники' 'Warning'; return }
    $script:EngList.RemoveAt($i)
    & $refresh ([Math]::Min($i, $script:EngList.Count - 1))
  })
  $bClose.Add_Click({ $dlg.Close() })

  & $refresh 0
  [void]$dlg.ShowDialog($form)
  $script:Settings.Engineers = @($script:EngList.ToArray())
  Save-Settings $script:Settings
}

# =====================================================================
#  Первый запуск: пути
# =====================================================================
if (-not (Test-TemplatesDir $script:Settings.TemplatesDir) -or [string]::IsNullOrWhiteSpace($script:Settings.OutDir)) {
  $ok = Show-SettingsDialog $true
  if (-not $ok) { return }
}

# =====================================================================
#  Главная форма
# =====================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Формирование актов КСБ-СОФТ'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(930, 900)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$pnlBottom = New-Object System.Windows.Forms.Panel
$pnlBottom.Dock = 'Bottom'
$pnlBottom.Height = 78
$form.Controls.Add($pnlBottom)

$pnlMain = New-Object System.Windows.Forms.Panel
$pnlMain.Dock = 'Fill'
$pnlMain.AutoScroll = $true
$form.Controls.Add($pnlMain)
$pnlMain.BringToFront()

$GW = 880  # ширина групп
$LX = 12   # левая колонка (подписи)
$BX = 235  # левая граница полей

function New-Group([string]$title) {
  $g = New-Object System.Windows.Forms.GroupBox
  $g.Text = $title
  $g.Width = $GW
  $g.Left = 8
  $pnlMain.Controls.Add($g)
  return $g
}
function New-Lbl($parent, [string]$text, [int]$x, [int]$y, [int]$w = 215) {
  $l = New-Object System.Windows.Forms.Label
  $l.Text = $text
  $l.Location = New-Object System.Drawing.Point($x, ($y + 3))
  $l.Size = New-Object System.Drawing.Size($w, 20)
  $parent.Controls.Add($l)
  return $l
}
function New-Txt($parent, [int]$x, [int]$y, [int]$w = 620) {
  $t = New-Object System.Windows.Forms.TextBox
  $t.Location = New-Object System.Drawing.Point($x, $y)
  $t.Size = New-Object System.Drawing.Size($w, 24)
  $parent.Controls.Add($t)
  return $t
}
function Add-Row($parent, [string]$label, [int]$y, [int]$w = 620) {
  $null = New-Lbl $parent $label $LX $y
  return (New-Txt $parent $BX $y $w)
}
function New-Date($parent, [int]$x, [int]$y) {
  $d = New-Object System.Windows.Forms.DateTimePicker
  $d.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
  $d.CustomFormat = 'dd.MM.yyyy'
  $d.Location = New-Object System.Drawing.Point($x, $y)
  $d.Size = New-Object System.Drawing.Size(110, 24)
  $d.Value = [datetime]::Today
  $parent.Controls.Add($d)
  return $d
}
function New-Chk($parent, [string]$text, [int]$x, [int]$y, [int]$w = 200) {
  $c = New-Object System.Windows.Forms.CheckBox
  $c.Text = $text
  $c.Location = New-Object System.Drawing.Point($x, $y)
  $c.Size = New-Object System.Drawing.Size($w, 22)
  $parent.Controls.Add($c)
  return $c
}
function New-Cmb($parent, [int]$x, [int]$y, [int]$w, [string]$style = 'DropDown') {
  $c = New-Object System.Windows.Forms.ComboBox
  $c.DropDownStyle = $style
  $c.Location = New-Object System.Drawing.Point($x, $y)
  $c.Size = New-Object System.Drawing.Size($w, 24)
  $parent.Controls.Add($c)
  return $c
}
function New-Btn($parent, [string]$text, [int]$x, [int]$y, [int]$w, [int]$h = 30) {
  $b = New-Object System.Windows.Forms.Button
  $b.Text = $text
  $b.Location = New-Object System.Drawing.Point($x, $y)
  $b.Size = New-Object System.Drawing.Size($w, $h)
  $parent.Controls.Add($b)
  return $b
}
# поле «только цифры» с префиксом и предпросмотром полного номера
function Add-IdRow($parent, [string]$label, [int]$y, [string]$kind, [string]$hint) {
  $null = New-Lbl $parent $label $LX $y
  $t = New-Txt $parent $BX $y 150
  $p = New-Object System.Windows.Forms.Label
  $p.Location = New-Object System.Drawing.Point(($BX + 160), ($y + 3))
  $p.Size = New-Object System.Drawing.Size(460, 20)
  $p.ForeColor = [System.Drawing.Color]::DimGray
  $p.Text = $hint
  $parent.Controls.Add($p)
  $t.Tag = @{ Kind = $kind; Preview = $p; Hint = $hint }
  $t.Add_TextChanged({
    $info = $this.Tag
    $full = Compose-Id $info.Kind $this.Text
    if ($full -eq '') { $info.Preview.Text = $info.Hint } else { $info.Preview.Text = 'в акте: ' + $full }
  })
  return $t
}

# ------------------------------------------------------------ Что оформляем
$grpType = New-Group 'Что оформляем'
$null = New-Lbl $grpType 'Тип акта' $LX 24
$cmbMode = New-Cmb $grpType $BX 24 620 'DropDownList'
[void]$cmbMode.Items.Add('По договору: SNS / Dallas Lock / ViPNet / КриптоПро / Континент TLS Клиент')
[void]$cmbMode.Items.Add('По контракту: Prime / Coordinator / Континент (список позиций)')
$cmbMode.SelectedIndex = 0

$chkAct = New-Chk $grpType 'Акт установки' $LX 58 215
$chkAct.Checked = $true
$chkAct.Enabled = $false
$null = New-Lbl $grpType 'дата акта:' 235 56 80
$dtAkta = New-Date $grpType 315 56
$null = New-Lbl $grpType 'дата установки на ПК:' 450 56 150
$dtUst = New-Date $grpType 600 56

$chkTp = New-Chk $grpType 'Сертификат ТП' $LX 90 215
$null = New-Lbl $grpType 'дата:' 235 88 80
$dtTp = New-Date $grpType 315 88

$chkIsp = New-Chk $grpType 'Акт испытаний' $LX 122 215
$null = New-Lbl $grpType 'дата:' 235 120 80
$dtIsp = New-Date $grpType 315 120
$grpType.Height = 156

# ------------------------------------------------------------ Заказчик
$grpCommon = New-Group 'Заказчик и договор'
$cy = 24
$null = New-Lbl $grpCommon 'Заказчик (полное название)' $LX $cy
$cmbZak = New-Cmb $grpCommon $BX $cy 620 'DropDown'
$cmbZak.AutoCompleteMode = 'SuggestAppend'
$cmbZak.AutoCompleteSource = 'ListItems'
$cy += 30
$cmbOsn = $null
$null = New-Lbl $grpCommon 'Основание' $LX $cy
$cmbOsn = New-Cmb $grpCommon $BX $cy 200 'DropDownList'
[void]$cmbOsn.Items.Add('Договор')
[void]$cmbOsn.Items.Add('Государственный контракт')
[void]$cmbOsn.Items.Add('Контракт')
$cmbOsn.SelectedIndex = 0
$null = New-Lbl $grpCommon '№' 450 $cy 22
$tbNomer = New-Txt $grpCommon 475 $cy 150
$tbNomer.Text = 'СТКБ-'
$null = New-Lbl $grpCommon 'от' 640 $cy 24
$dtDog = New-Date $grpCommon 665 $cy
$cy += 30
$null = New-Lbl $grpCommon 'Город в акте установки' $LX $cy
$cmbGorod = New-Cmb $grpCommon $BX $cy 300 'DropDown'
[void]$cmbGorod.Items.AddRange(@('г. Чебоксары', 'г. Новочебоксарск', 'г. Тверь', 'г. Брянск', 'г. Калуга'))
$cmbGorod.Text = 'г. Чебоксары'
$cy += 30
$tbAdres = Add-Row $grpCommon 'Адрес заказчика' $cy
$cy += 30
$lblGorodIsp = New-Lbl $grpCommon 'Город в акте испытаний' $LX $cy
$cmbGorodIsp = New-Cmb $grpCommon $BX $cy 300 'DropDown'
[void]$cmbGorodIsp.Items.AddRange(@('г. Малоярославец', 'с. Детчино', 'с. Недельное', 'г. Чебоксары'))
$cmbGorodIsp.Text = 'г. Малоярославец'
$cy += 30
$lblZakRod = New-Lbl $grpCommon 'Заказчик в дат. падеже (сертификат)' $LX $cy
$tbZakRod = New-Txt $grpCommon $BX $cy 620
$cy += 32
$grpCommon.Height = $cy + 4

# ------------------------------------------------------------ СЗИ по договору
$grpDog = New-Group 'Что установлено (по договору)'
$chkSns = New-Chk $grpDog 'Secret Net Studio' 12 22 150
$chkDl = New-Chk $grpDog 'Dallas Lock' 170 22 120
$chkTls = New-Chk $grpDog 'Континент TLS Клиент' 296 22 170
$chkVip = New-Chk $grpDog 'ViPNet Client' 472 22 120
$chkKp = New-Chk $grpDog 'КриптоПро CSP' 598 22 140

function New-SubPanel([int]$h) {
  $p = New-Object System.Windows.Forms.Panel
  $p.Size = New-Object System.Drawing.Size(($GW - 20), $h)
  $p.Left = 6
  $grpDog.Controls.Add($p)
  return $p
}
# SNS
$pnlSns = New-SubPanel 160
$chkSns30 = New-Chk $pnlSns 'Версия «30» в названии (Secret Net Studio ... 001 30)' 6 0 520
$chkSns30.Checked = $true
$tbSnsId = Add-IdRow $pnlSns 'Идентификатор SNS' 28 'SNS' '6 последних цифр: РОСС RU.01.03745.XXXXXX'
$null = New-Lbl $pnlSns 'Модули (снимите галочку у тех, которых нет):' 6 58 400
$chkMods = @{}
$modDefs = @(
  @('nsd', 'защита от НСД', 6, 82), @('dev', 'контроль устройств', 6, 106), @('disk', 'защита диска и шифрование контейнеров', 6, 130),
  @('fw', 'персональный межсетевой экран', 440, 82), @('ids', 'обнаружение вторжений', 440, 106), @('av', 'антивирус (технология Касперского)', 440, 130)
)
foreach ($m in $modDefs) {
  $c = New-Chk $pnlSns $m[1] $m[2] $m[3] 400
  $c.Checked = $true
  $chkMods[$m[0]] = $c
}
# Dallas
$pnlDl = New-SubPanel 32
$tbDlId = Add-IdRow $pnlDl 'Идентификатор Dallas Lock' 2 'DL' '6 последних цифр: РОСС RU.0001.02720.XXXXXX'
# TLS
$pnlTls = New-SubPanel 62
$tbTlsSerial = Add-Row $pnlTls 'Заводской номер TLS Клиент' 2 150
$tbTlsReg = Add-IdRow $pnlTls 'Рег. номер (цифры)' 32 'TLSREG' 'например 1109: 1073Б1-001109'
# ViPNet
$pnlVip = New-SubPanel 92
$null = New-Lbl $pnlVip 'Название в акте' $LX 2
$cmbVipTitle = New-Cmb $pnlVip $BX 2 560 'DropDown'
[void]$cmbVipTitle.Items.Add('Программный комплекс ViPNet Client for Windows 4.х (КС2)')
[void]$cmbVipTitle.Items.Add('Программный комплекс ViPNet Client 4 (версия 4.5)')
$cmbVipTitle.SelectedIndex = 0
$tbVipSerial = Add-IdRow $pnlVip 'Серийный № (6 цифр)' 32 'VIPSN' 'КлКС2-4-XXXXXX'
$null = New-Lbl $pnlVip 'Рег. № СКЗИ' $LX 62
$tbVipReg = New-Txt $pnlVip $BX 62 150
$lblVipReg = New-Object System.Windows.Forms.Label
$lblVipReg.Location = New-Object System.Drawing.Point(($BX + 160), 65)
$lblVipReg.Size = New-Object System.Drawing.Size(460, 20)
$lblVipReg.ForeColor = [System.Drawing.Color]::DimGray
$lblVipReg.Text = 'считается сам: 782-(серийный + 971); можно исправить вручную'
$pnlVip.Controls.Add($lblVipReg)
# КриптоПро
$pnlKp = New-SubPanel 32
$tbKpId = Add-IdRow $pnlKp 'Идентификатор КриптоПро' 2 'KP' 'последние цифры: 385Б1А1-00XXXX'

# ------------------------------------------------------------ Рабочие места
$grpPlaces = New-Group 'Рабочие места (где установлено)'
$lblPl = New-Object System.Windows.Forms.Label
$lblPl.Location = New-Object System.Drawing.Point(12, 20)
$lblPl.Size = New-Object System.Drawing.Size(850, 20)
$lblPl.ForeColor = [System.Drawing.Color]::DimGray
$lblPl.Text = 'Одна строка - пункты 6.3-6.5. Несколько строк - таблица. Дату можно не заполнять (возьмётся «дата установки на ПК»).'
$grpPlaces.Controls.Add($lblPl)
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(12, 44)
$dgv.Size = New-Object System.Drawing.Size(850, 150)
$dgv.AllowUserToAddRows = $true
$dgv.AllowUserToDeleteRows = $true
$dgv.RowHeadersVisible = $false
$dgv.AutoSizeColumnsMode = 'Fill'
$dgv.BackgroundColor = [System.Drawing.Color]::White
[void]$dgv.Columns.Add('Room', 'Помещение (номер / наименование)')
[void]$dgv.Columns.Add('Pc', 'Компьютер (сетевое имя / инвентарный номер)')
[void]$dgv.Columns.Add('Date', 'Дата установки (необязательно)')
$dgv.Columns[2].FillWeight = 40
$grpPlaces.Controls.Add($dgv)
$grpPlaces.Height = 206

# ------------------------------------------------------------ СЗИ по контракту
$grpKon = New-Group 'Состав СЗИ (по контракту) - добавляйте позиции'
$tbPomDef = Add-Row $grpKon 'Помещение (по умолчанию)' 22 300
$lstProd = New-Object System.Windows.Forms.ListBox
$lstProd.Location = New-Object System.Drawing.Point(12, 58)
$lstProd.Size = New-Object System.Drawing.Size(330, 230)
$grpKon.Controls.Add($lstProd)
$cmbAddType = New-Cmb $grpKon 12 296 200 'DropDownList'
[void]$cmbAddType.Items.AddRange(@('ViPNet Prime', 'ViPNet Coordinator', 'Континент 3.9', 'Континент TLS-сервер'))
$cmbAddType.SelectedIndex = 1
$btnAdd = New-Btn $grpKon 'Добавить' 220 294 55 28
$btnAdd.Width = 62
$btnDel = New-Btn $grpKon 'Удалить' 288 294 54 28
$btnDel.Width = 54
$pnlEdit = New-Object System.Windows.Forms.Panel
$pnlEdit.Location = New-Object System.Drawing.Point(356, 52)
$pnlEdit.Size = New-Object System.Drawing.Size(510, 290)
$grpKon.Controls.Add($pnlEdit)
$edit = @{}
$ey = 0
foreach ($def in @(
  @('Name', 'Наименование'), @('Serial', 'Серийный / заводской №'), @('Platform', 'Аппаратная платформа'),
  @('Reg', 'Рег. № СКЗИ'), @('Ident', 'Идентификатор СЗИ'), @('IdentShort', 'Идентификатор (Континент)'),
  @('Room', 'Помещение'), @('Addr', 'Адрес (если другой)'), @('Computer', 'Компьютер (Prime)'), @('InstDate', 'Дата установки')
)) {
  $null = New-Lbl $pnlEdit $def[1] 0 $ey 170
  if ($def[0] -eq 'InstDate') {
    $edit[$def[0]] = New-Date $pnlEdit 175 $ey
  } else {
    $edit[$def[0]] = New-Txt $pnlEdit 175 $ey 320
  }
  $ey += 26
}
$chkFsb = New-Chk $pnlEdit 'Есть сертификат ФСБ (Континент КШ)' 0 $ey 330
$grpKon.Height = 352

# ------------------------------------------------------------ Подписант
$grpSign = New-Group 'Исполнитель, подписывающий акт установки'
$null = New-Lbl $grpSign 'Сотрудник' $LX 24
$cmbEng = New-Cmb $grpSign $BX 24 250 'DropDownList'
$btnEng = New-Btn $grpSign 'Сотрудники... (добавить / изменить)' 495 22 250 28
$tbIzpDolzh = Add-Row $grpSign 'Должность в акте' 56 620
$tbIzpInit = Add-Row $grpSign 'Расшифровка подписи' 86 250
$grpSign.Height = 122

# ------------------------------------------------------------ Сертификат ТП
$grpTp = New-Group 'Сертификат ТП (срок ТП - в шаблоне Сертификат_ТП.docx)'
$null = New-Lbl $grpTp 'Продукты (по строке на продукт)' $LX 24
$tbTpProducts = New-Object System.Windows.Forms.TextBox
$tbTpProducts.Multiline = $true
$tbTpProducts.Location = New-Object System.Drawing.Point($BX, 24)
$tbTpProducts.Size = New-Object System.Drawing.Size(620, 56)
$grpTp.Controls.Add($tbTpProducts)
$tbTpDob = Add-Row $grpTp 'Добавочный в телефонах ТП' 90 80
$tbTpDob.Text = [string]$script:Settings.TpDob
$grpTp.Height = 124

# ------------------------------------------------------------ Акт испытаний
$grpIsp = New-Group 'Акт испытаний: комиссии'
$null = New-Lbl $grpIsp 'От Исполнителя - член 1' $LX 24
$cmbIsp1 = New-Cmb $grpIsp $BX 24 250 'DropDownList'
$null = New-Lbl $grpIsp 'От Исполнителя - член 2' $LX 54
$cmbIsp2 = New-Cmb $grpIsp $BX 54 250 'DropDownList'
$tbZakFio1 = Add-Row $grpIsp 'От Заказчика 1: ФИО' 90 300
$tbZakDolzh1 = Add-Row $grpIsp 'должность' 120 300
$tbZakFio2 = Add-Row $grpIsp 'От Заказчика 2: ФИО' 150 300
$tbZakDolzh2 = Add-Row $grpIsp 'должность' 180 300
$grpIsp.Height = 216

# =====================================================================
#  Нижняя панель
# =====================================================================
$btnGen = New-Btn $pnlBottom 'Сформировать' 12 8 170 36
$btnGen.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnOpen = New-Btn $pnlBottom 'Открыть «Готовые»' 192 8 150 36
$btnClear = New-Btn $pnlBottom 'Очистить форму' 352 8 130 36
$btnSettings = New-Btn $pnlBottom 'Настройки (папки)' 492 8 140 36
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12, 50)
$lblStatus.Size = New-Object System.Drawing.Size(880, 22)
$lblStatus.Anchor = 'Left,Right,Bottom'
$pnlBottom.Controls.Add($lblStatus)
$chkAutoClear = New-Object System.Windows.Forms.CheckBox
$chkAutoClear.Text = 'Очищать форму после создания'
$chkAutoClear.Location = New-Object System.Drawing.Point(650, 16)
$chkAutoClear.Size = New-Object System.Drawing.Size(250, 22)
$pnlBottom.Controls.Add($chkAutoClear)

# =====================================================================
#  Логика
# =====================================================================
function Refresh-EngineerCombos {
  $names = @($script:Settings.Engineers | ForEach-Object { [string]$_.Name })
  foreach ($c in @($cmbEng, $cmbIsp1, $cmbIsp2)) {
    $prev = [string]$c.Text
    $c.Items.Clear()
    foreach ($n in $names) { [void]$c.Items.Add($n) }
    if ($prev -and ($names -contains $prev)) { $c.Text = $prev }
    elseif ($c.Items.Count -gt 0) { $c.SelectedIndex = 0 }
  }
}

function Apply-Engineer {
  $e = Get-Engineer $script:Settings ([string]$cmbEng.Text)
  if ($null -ne $e) {
    $tbIzpDolzh.Text = [string]$e.Dolzh
    $tbIzpInit.Text = [string]$e.Init
  }
}

function Update-TpProducts {
  if ($tbTpProducts.Tag -eq 'manual') { return }
  $lines = @()
  if ($cmbMode.SelectedIndex -eq 1) {
    $lines = @('Средства защиты информации (комплекс)')
  } elseif ($chkSns.Checked) {
    $lines = @('Средство защиты информации Secret Net Studio')
  } else {
    if ($chkDl.Checked) { $lines += 'Система защиты информации от несанкционированного доступа Dallas Lock 8.0-К' }
    if ($chkTls.Checked) { $lines += 'Средство криптографической защиты информации Континент TLS Клиент' }
    if ($chkVip.Checked) { $lines += 'Программный комплекс ViPNet Client' }
    if ($chkKp.Checked) { $lines += 'Средство криптографической защиты информации КриптоПро CSP' }
  }
  $script:Sync = $true
  $tbTpProducts.Text = ($lines -join "`r`n")
  $script:Sync = $false
}

function Update-Layout {
  $kon = ($cmbMode.SelectedIndex -eq 1)
  $grpDog.Visible = -not $kon
  $grpPlaces.Visible = -not $kon
  $grpKon.Visible = $kon
  $chkIsp.Enabled = -not $kon
  if ($kon) { $chkIsp.Checked = $false }
  $grpTp.Visible = $chkTp.Checked
  $grpIsp.Visible = ($chkIsp.Checked -and -not $kon)
  $dtTp.Enabled = $chkTp.Checked
  $dtIsp.Enabled = ($chkIsp.Checked -and -not $kon)
  $lblZakRod.Visible = $chkTp.Checked
  $tbZakRod.Visible = $chkTp.Checked
  $showIsp = ($chkIsp.Checked -and -not $kon)
  $lblGorodIsp.Visible = $showIsp
  $cmbGorodIsp.Visible = $showIsp

  # внутри «Заказчик»: сдвигаем строки, чтобы не оставалось дыр
  $yy = 24 + 30 + 30 + 30 + 30
  $lblGorodIsp.Top = $yy + 3; $cmbGorodIsp.Top = $yy
  if ($showIsp) { $yy += 30 }
  $lblZakRod.Top = $yy + 3; $tbZakRod.Top = $yy
  if ($chkTp.Checked) { $yy += 32 }
  $grpCommon.Height = $yy + 8

  # внутри «Что установлено»: подпанели друг под другом
  $py = 52
  foreach ($pair in @(@($chkSns, $pnlSns), @($chkDl, $pnlDl), @($chkTls, $pnlTls), @($chkVip, $pnlVip), @($chkKp, $pnlKp))) {
    $pair[1].Visible = $pair[0].Checked
    if ($pair[0].Checked) { $pair[1].Top = $py; $py += $pair[1].Height + 4 }
  }
  $grpDog.Height = $py + 6

  $y = 6
  foreach ($g in @($grpType, $grpCommon, $grpDog, $grpKon, $grpPlaces, $grpSign, $grpTp, $grpIsp)) {
    if ($g.Visible) { $g.Top = $y; $y = $g.Bottom + 8 }
  }
  Update-TpProducts
}

# --- даты: пока вы не трогали дату вручную, она следует за «датой акта» ---
foreach ($d in @($dtUst, $dtTp, $dtIsp)) { $d.Tag = 'auto' }
$dtAkta.Add_ValueChanged({
  if ($script:Sync) { return }
  $script:Sync = $true
  foreach ($d in @($dtUst, $dtTp, $dtIsp)) { if ($d.Tag -eq 'auto') { $d.Value = $dtAkta.Value } }
  $script:Sync = $false
})
foreach ($d in @($dtUst, $dtTp, $dtIsp)) {
  $d.Add_ValueChanged({ if (-not $script:Sync) { $this.Tag = 'manual' } })
}

# --- ручная правка падежа / продуктов ТП ---
$tbZakRod.Tag = 'auto'
$tbZakRod.Add_TextChanged({ if (-not $script:Sync) { $this.Tag = 'manual' } })
$tbTpProducts.Tag = 'auto'
$tbTpProducts.Add_TextChanged({ if (-not $script:Sync) { $this.Tag = 'manual' } })

$cmbZak.Add_TextChanged({
  if ($tbZakRod.Tag -eq 'manual') { return }
  $script:Sync = $true
  $tbZakRod.Text = Get-Rod (Convert-Quotes $cmbZak.Text)
  $script:Sync = $false
})

# --- история заказчиков ---
function Refresh-Customers {
  $cmbZak.Items.Clear()
  foreach ($c in @($script:Settings.Customers)) { [void]$cmbZak.Items.Add([string]$c.Name) }
}
function Find-Customer([string]$name) {
  foreach ($c in @($script:Settings.Customers)) { if ([string]$c.Name -eq $name) { return $c } }
  return $null
}
function CustVal($c, [string]$n) {
  $p = $c.PSObject.Properties[$n]
  if ($null -ne $p -and $null -ne $p.Value) { return [string]$p.Value }
  return ''
}
$cmbZak.Add_SelectionChangeCommitted({
  $c = Find-Customer ([string]$cmbZak.SelectedItem)
  if ($null -eq $c) { return }
  $script:Sync = $true
  if (CustVal $c 'Gorod') { $cmbGorod.Text = CustVal $c 'Gorod' }
  if (CustVal $c 'GorodIsp') { $cmbGorodIsp.Text = CustVal $c 'GorodIsp' }
  $tbAdres.Text = CustVal $c 'Adres'
  if (CustVal $c 'Rod') { $tbZakRod.Text = CustVal $c 'Rod'; $tbZakRod.Tag = 'manual' }
  $tbZakFio1.Text = CustVal $c 'ZakFio1'; $tbZakDolzh1.Text = CustVal $c 'ZakDolzh1'
  $tbZakFio2.Text = CustVal $c 'ZakFio2'; $tbZakDolzh2.Text = CustVal $c 'ZakDolzh2'
  $script:Sync = $false
})
function Save-Customer {
  $name = $cmbZak.Text.Trim()
  if ($name -eq '') { return }
  $rec = [pscustomobject]@{
    Name = $name; Gorod = $cmbGorod.Text.Trim(); GorodIsp = $cmbGorodIsp.Text.Trim(); Adres = $tbAdres.Text.Trim()
    Rod = $tbZakRod.Text.Trim(); ZakFio1 = $tbZakFio1.Text.Trim(); ZakDolzh1 = $tbZakDolzh1.Text.Trim()
    ZakFio2 = $tbZakFio2.Text.Trim(); ZakDolzh2 = $tbZakDolzh2.Text.Trim()
  }
  $list = New-Object System.Collections.Generic.List[object]
  foreach ($c in @($script:Settings.Customers)) { if ([string]$c.Name -ne $name) { $list.Add($c) } }
  $list.Insert(0, $rec)
  $script:Settings.Customers = @($list.ToArray() | Select-Object -First 300)
}

# --- ViPNet: рег. номер считается сам ---
$script:VipRegAuto = $true
$tbVipReg.Add_TextChanged({ if (-not $script:Sync) { $script:VipRegAuto = ($tbVipReg.Text.Trim() -eq '') } })
$tbVipSerial.Add_TextChanged({
  if ($script:VipRegAuto) {
    $script:Sync = $true
    $tbVipReg.Text = Get-VipReg (Compose-Id 'VIPSN' $tbVipSerial.Text)
    $script:Sync = $false
  }
})

# --- режим / основание ---
$cmbMode.Add_SelectedIndexChanged({
  if ($cmbMode.SelectedIndex -eq 1) {
    $cmbOsn.SelectedIndex = 1
    if ($tbNomer.Text -eq 'СТКБ-') { $tbNomer.Text = '' }
    $cmbGorod.Text = 'г. Тверь'
  } else {
    $cmbOsn.SelectedIndex = 0
    if ($tbNomer.Text.Trim() -eq '') { $tbNomer.Text = 'СТКБ-' }
    $cmbGorod.Text = 'г. Чебоксары'
  }
  Update-Layout
})
foreach ($c in @($chkTp, $chkIsp, $chkSns, $chkDl, $chkTls, $chkVip, $chkKp)) { $c.Add_CheckedChanged({ Update-Layout }) }
$cmbEng.Add_SelectedIndexChanged({ Apply-Engineer })

# --- позиции контракта ---
function Product-Caption($p) {
  $name = switch ($p.Type) {
    'Prime' { 'Prime' } 'Coordinator' { 'Coordinator' } 'Continent' { 'Континент' } 'ContinentTLS' { 'Континент TLS' } default { $p.Type }
  }
  $extra = ''
  if ($p.Type -eq 'Prime' -or $p.Type -eq 'Coordinator') { $extra = $p.Serial } else { $extra = $p.Factory }
  if ([string]::IsNullOrWhiteSpace($extra)) { return $name }
  return "$name - $extra"
}
function Refresh-ProductList([int]$sel = -1) {
  $script:Sync = $true
  $lstProd.Items.Clear()
  for ($i = 0; $i -lt $script:Products.Count; $i++) {
    [void]$lstProd.Items.Add(('{0}. {1}' -f ($i + 1), (Product-Caption $script:Products[$i])))
  }
  if ($sel -ge 0 -and $sel -lt $lstProd.Items.Count) { $lstProd.SelectedIndex = $sel }
  elseif ($lstProd.Items.Count -gt 0) { $lstProd.SelectedIndex = $lstProd.Items.Count - 1 }
  $script:Sync = $false
  Load-ProductEditor
}
function New-Product([string]$kind) {
  $p = [pscustomobject]@{
    Type = $kind; Name = ''; Serial = ''; Factory = ''; Platform = ''; RegSkzi = ''; Identifikator = ''; Ident = ''
    Room = $tbPomDef.Text.Trim(); Address = ''; Computer = ''; HasFsbCert = $false
    InstallDate = $dtUst.Value.ToString('dd.MM.yyyy')
  }
  switch ($kind) {
    'Prime' { $p.Name = 'ViPNet Prime' }
    'Coordinator' { $p.Name = 'ViPNet Coordinator HW 1000 C'; $p.Platform = 'HW1000 Q8' }
    'Continent' { $p.Platform = 'IPC-R300. Детектор атак' }
    'ContinentTLS' { $p.Platform = 'IPCR300' }
  }
  return $p
}
function Save-ProductEditor {
  if ($script:Sync -or $lstProd.SelectedIndex -lt 0) { return }
  $p = $script:Products[$lstProd.SelectedIndex]
  $p.Name = $edit['Name'].Text.Trim()
  if ($p.Type -eq 'Prime' -or $p.Type -eq 'Coordinator') { $p.Serial = $edit['Serial'].Text.Trim() } else { $p.Factory = $edit['Serial'].Text.Trim() }
  $p.Platform = $edit['Platform'].Text.Trim()
  $p.RegSkzi = $edit['Reg'].Text.Trim()
  $p.Identifikator = $edit['Ident'].Text.Trim()
  $p.Ident = $edit['IdentShort'].Text.Trim()
  $p.Room = $edit['Room'].Text.Trim()
  $p.Address = $edit['Addr'].Text.Trim()
  $p.Computer = $edit['Computer'].Text.Trim()
  $p.InstallDate = $edit['InstDate'].Value.ToString('dd.MM.yyyy')
  $p.HasFsbCert = $chkFsb.Checked
}
function Load-ProductEditor {
  $has = ($lstProd.SelectedIndex -ge 0)
  foreach ($k in $edit.Keys) { $edit[$k].Enabled = $has }
  $chkFsb.Enabled = $false
  if (-not $has) {
    foreach ($k in $edit.Keys) { if ($edit[$k] -is [System.Windows.Forms.TextBox]) { $edit[$k].Text = '' } }
    return
  }
  $p = $script:Products[$lstProd.SelectedIndex]
  $script:Sync = $true
  $edit['Name'].Text = [string]$p.Name
  if ($p.Type -eq 'Prime' -or $p.Type -eq 'Coordinator') { $edit['Serial'].Text = [string]$p.Serial } else { $edit['Serial'].Text = [string]$p.Factory }
  $edit['Platform'].Text = [string]$p.Platform
  $edit['Reg'].Text = [string]$p.RegSkzi
  $edit['Ident'].Text = [string]$p.Identifikator
  $edit['IdentShort'].Text = [string]$p.Ident
  $edit['Room'].Text = [string]$p.Room
  $edit['Addr'].Text = [string]$p.Address
  $edit['Computer'].Text = [string]$p.Computer
  try { $edit['InstDate'].Value = [datetime]::ParseExact([string]$p.InstallDate, 'dd.MM.yyyy', $null) } catch { $edit['InstDate'].Value = $dtUst.Value }
  $chkFsb.Checked = [bool]$p.HasFsbCert
  $chkFsb.Enabled = ($p.Type -eq 'Continent')
  $edit['Computer'].Enabled = ($p.Type -eq 'Prime')
  $edit['IdentShort'].Enabled = ($p.Type -eq 'Continent')
  $edit['Reg'].Enabled = ($p.Type -eq 'Coordinator' -or $p.Type -eq 'Continent')
  $edit['Platform'].Enabled = ($p.Type -ne 'Prime')
  $edit['Name'].Enabled = ($p.Type -eq 'Prime' -or $p.Type -eq 'Coordinator')
  $script:Sync = $false
}
$btnAdd.Add_Click({
  Save-ProductEditor
  $kind = @('Prime', 'Coordinator', 'Continent', 'ContinentTLS')[$cmbAddType.SelectedIndex]
  $script:Products.Add((New-Product $kind))
  Refresh-ProductList ($script:Products.Count - 1)
})
$btnDel.Add_Click({
  if ($lstProd.SelectedIndex -lt 0) { return }
  $i = $lstProd.SelectedIndex
  $script:Products.RemoveAt($i)
  Refresh-ProductList ([Math]::Min($i, $script:Products.Count - 1))
})
$lstProd.Add_SelectedIndexChanged({ if (-not $script:Sync) { Load-ProductEditor } })
foreach ($k in $edit.Keys) {
  if ($edit[$k] -is [System.Windows.Forms.TextBox]) {
    $edit[$k].Add_TextChanged({
      if ($script:Sync) { return }
      Save-ProductEditor
      $i = $lstProd.SelectedIndex
      if ($i -ge 0) {
        $script:Sync = $true
        $lstProd.Items[$i] = ('{0}. {1}' -f ($i + 1), (Product-Caption $script:Products[$i]))
        $script:Sync = $false
      }
    })
  } else {
    $edit[$k].Add_ValueChanged({ Save-ProductEditor })
  }
}
$chkFsb.Add_CheckedChanged({ Save-ProductEditor })

# --- кнопки ---
$btnEng.Add_Click({
  $cur = [string]$cmbEng.Text
  Show-EngineersDialog
  Refresh-EngineerCombos
  if ($cur -and (Get-Engineer $script:Settings $cur)) { $cmbEng.Text = $cur }
  Apply-Engineer
})
$btnSettings.Add_Click({ [void](Show-SettingsDialog $false) })
$btnOpen.Add_Click({
  $o = [string]$script:Settings.OutDir
  if (-not (Test-Path -LiteralPath $o)) { New-Item -ItemType Directory -Force -Path $o | Out-Null }
  Start-Process explorer.exe -ArgumentList ('"{0}"' -f $o)
})
$btnClear.Add_Click({
  $script:Sync = $true
  $cmbZak.Text = ''; $tbAdres.Text = ''; $tbZakRod.Text = ''; $tbZakRod.Tag = 'auto'
  $tbNomer.Text = 'СТКБ-'
  foreach ($t in @($tbSnsId, $tbDlId, $tbTlsSerial, $tbTlsReg, $tbVipSerial, $tbVipReg, $tbKpId, $tbZakFio1, $tbZakDolzh1, $tbZakFio2, $tbZakDolzh2)) { $t.Text = '' }
  $script:VipRegAuto = $true
  $dgv.Rows.Clear()
  $script:Products.Clear()
  $dtAkta.Value = [datetime]::Today
  foreach ($d in @($dtUst, $dtTp, $dtIsp)) { $d.Value = [datetime]::Today; $d.Tag = 'auto' }
  $dtDog.Value = [datetime]::Today
  $tbTpProducts.Tag = 'auto'
  foreach ($c in $chkMods.Values) { $c.Checked = $true }
  $script:Sync = $false
  Refresh-ProductList
  Update-Layout
  $lblStatus.Text = 'Форма очищена.'
})

# Ошибка ввода: запоминаем поле, чтобы после сообщения поставить в него курсор
$script:FailCtrl = $null
function Fail($ctrl, [string]$msg) {
  $script:FailCtrl = $ctrl
  throw $msg
}

function Collect-Warnings {
  $w = @()
  # даты: мягкие подсказки, на генерацию не влияют
  $kon0 = ($cmbMode.SelectedIndex -eq 1)
  if ($dtDog.Value.Date -gt $dtAkta.Value.Date) { $w += 'Дата договора позже даты акта установки.' }
  elseif (-not $kon0 -and $dtDog.Value.Date -eq $dtAkta.Value.Date) { $w += 'Дата договора совпадает с датой акта - не забыли её изменить?' }
  if (-not $kon0 -and $dtUst.Value.Date -gt $dtAkta.Value.Date) { $w += 'Дата установки на ПК позже даты акта установки.' }
  if ($chkIsp.Checked -and -not $kon0 -and $dtIsp.Value.Date -lt $dtAkta.Value.Date) { $w += 'Дата акта испытаний раньше даты акта установки.' }
  $chk = { param($txt, $len, $name)
    $v = ([string]$txt).Trim()
    if ($v -match '^\d+$' -and $v.Length -ne $len) { return ("{0}: введено {1} цифр, обычно {2}." -f $name, $v.Length, $len) }
    return $null
  }
  if ($chkSns.Checked) { $m = & $chk $tbSnsId.Text 6 'Идентификатор SNS'; if ($m) { $w += $m } }
  if ($chkDl.Checked) { $m = & $chk $tbDlId.Text 6 'Идентификатор Dallas Lock'; if ($m) { $w += $m } }
  if ($chkVip.Checked) { $m = & $chk $tbVipSerial.Text 6 'Серийный № ViPNet'; if ($m) { $w += $m } }
  if ($chkKp.Checked) { $m = & $chk $tbKpId.Text 6 'Идентификатор КриптоПро'; if ($m -and $tbKpId.Text.Trim().Length -ne 4) { $w += $m } }
  return $w
}

$btnGen.Add_Click({
  $script:FailCtrl = $null
  try {
    Save-ProductEditor
    $kon = ($cmbMode.SelectedIndex -eq 1)
    $nomerTxt = $tbNomer.Text.Trim()
    if ($nomerTxt -eq '' -or $nomerTxt -eq 'СТКБ-') { Fail $tbNomer 'Укажите номер договора / контракта.' }
    if ($cmbZak.Text.Trim() -eq '') { Fail $cmbZak 'Укажите заказчика.' }
    if ($cmbGorod.Text.Trim() -eq '') { Fail $cmbGorod 'Укажите город.' }
    if ($tbIzpDolzh.Text.Trim() -eq '') { Fail $tbIzpDolzh 'Укажите должность исполнителя.' }
    if ($tbIzpInit.Text.Trim() -eq '') { Fail $tbIzpInit 'Укажите расшифровку подписи исполнителя.' }

    $rows = @()
    if (-not $kon) {
      if (-not ($chkSns.Checked -or $chkDl.Checked -or $chkTls.Checked -or $chkVip.Checked -or $chkKp.Checked)) { Fail $chkSns 'Отметьте, что установлено (SNS, Dallas Lock, ViPNet ...).' }
      if ($chkSns.Checked -and $tbSnsId.Text.Trim() -eq '') { Fail $tbSnsId 'Укажите идентификатор SNS.' }
      if ($chkDl.Checked -and $tbDlId.Text.Trim() -eq '') { Fail $tbDlId 'Укажите идентификатор Dallas Lock.' }
      if ($chkTls.Checked -and $tbTlsSerial.Text.Trim() -eq '') { Fail $tbTlsSerial 'Укажите заводской номер Континент TLS Клиент.' }
      if ($chkTls.Checked -and $tbTlsReg.Text.Trim() -eq '') { Fail $tbTlsReg 'Укажите регистрационный номер Континент TLS Клиент.' }
      if ($chkVip.Checked -and $tbVipSerial.Text.Trim() -eq '') { Fail $tbVipSerial 'Укажите серийный номер ViPNet.' }
      if ($chkVip.Checked -and $tbVipReg.Text.Trim() -eq '') { Fail $tbVipReg 'Укажите регистрационный номер ViPNet.' }
      if ($chkKp.Checked -and $tbKpId.Text.Trim() -eq '') { Fail $tbKpId 'Укажите идентификатор КриптоПро.' }
      foreach ($r in $dgv.Rows) {
        if ($r.IsNewRow) { continue }
        $room = [string]$r.Cells[0].Value; $pc = [string]$r.Cells[1].Value; $dt = [string]$r.Cells[2].Value
        if ($room.Trim() -eq '' -and $pc.Trim() -eq '') { continue }
        $rows += @{ Room = $room.Trim(); Pc = $pc.Trim(); Date = $dt.Trim() }
      }
      if ($rows.Count -eq 0) { Fail $dgv 'Заполните хотя бы одно рабочее место (помещение / компьютер).' }
    } else {
      if ($script:Products.Count -lt 1) { throw 'Добавьте хотя бы одну позицию СЗИ.' }
    }
    if ($chkTp.Checked -and $tbTpProducts.Text.Trim() -eq '') { Fail $tbTpProducts 'Для сертификата ТП укажите продукты.' }
    if ($chkIsp.Checked) {
      if ($cmbGorodIsp.Text.Trim() -eq '') { Fail $cmbGorodIsp 'Для акта испытаний укажите город.' }
      if ($tbZakFio1.Text.Trim() -eq '') { Fail $tbZakFio1 'Для акта испытаний укажите ФИО представителя заказчика.' }
    }

    $warn = @(Collect-Warnings)
    if ($warn.Count -gt 0) {
      $a = [System.Windows.Forms.MessageBox]::Show(("Проверьте данные:`n`n{0}`n`nВсё равно сформировать?" -f ($warn -join "`n")), 'Проверка', 'YesNo', 'Question')
      if ($a -ne 'Yes') { return }
    }

    $mods = @(); foreach ($k in 'nsd', 'dev', 'disk', 'fw', 'ids', 'av') { if ($chkMods[$k].Checked) { $mods += $k } }
    $e1 = Get-Engineer $script:Settings ([string]$cmbIsp1.Text)
    $e2 = Get-Engineer $script:Settings ([string]$cmbIsp2.Text)
    function Eng-Fio($e) { if ($null -eq $e) { return '' }; if ([string]$e.Fio -ne '') { return [string]$e.Fio }; return (Get-FioFromInit ([string]$e.Init)) }

    $prodList = @($script:Products | ForEach-Object {
      @{
        Type = $_.Type; Name = $_.Name; Serial = $_.Serial; Factory = $_.Factory; Platform = $_.Platform
        RegSkzi = $_.RegSkzi; Identifikator = $_.Identifikator; Ident = $_.Ident; Room = $_.Room
        Address = $_.Address; Computer = $_.Computer; HasFsbCert = [bool]$_.HasFsbCert; InstallDate = $_.InstallDate
      }
    })

    $data = @{
      Mode          = $(if ($kon) { 'Kontrakt' } else { 'Dogovor' })
      Nomer         = $nomerTxt
      DataDogovora  = $dtDog.Value.ToString('dd.MM.yyyy')
      OsnovanieTip  = [string]$cmbOsn.Text
      Gorod         = $cmbGorod.Text.Trim()
      Zakazchik     = $cmbZak.Text.Trim()
      ZakazchikRod  = $tbZakRod.Text.Trim()
      Adres         = $tbAdres.Text.Trim()
      DataAkta      = $dtAkta.Value.ToString('dd.MM.yyyy')
      DataUstanovki = $dtUst.Value.ToString('dd.MM.yyyy')
      NeedTp        = $chkTp.Checked
      DataTp        = $dtTp.Value.ToString('dd.MM.yyyy')
      NeedIsp       = $chkIsp.Checked
      DataIsp       = $dtIsp.Value.ToString('dd.MM.yyyy')
      GorodIsp      = $cmbGorodIsp.Text.Trim()
      IzpDolzh      = $tbIzpDolzh.Text.Trim()
      IzpInit       = $tbIzpInit.Text.Trim()
      HasSns = $chkSns.Checked; Sns30 = $chkSns30.Checked; SnsIdent = $tbSnsId.Text.Trim(); SnsModules = $mods
      HasDl = $chkDl.Checked; DlIdent = $tbDlId.Text.Trim()
      HasTls = $chkTls.Checked; TlsSerial = $tbTlsSerial.Text.Trim(); TlsReg = $tbTlsReg.Text.Trim()
      HasVip = $chkVip.Checked; VipTitle = $cmbVipTitle.Text.Trim(); VipSerial = $tbVipSerial.Text.Trim(); VipReg = $tbVipReg.Text.Trim()
      HasKp = $chkKp.Checked; KpIdent = $tbKpId.Text.Trim()
      Workplaces    = $rows
      Products      = $prodList
      TpProducts    = @($tbTpProducts.Text -split "`r?`n" | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() })
      TpDob         = $tbTpDob.Text.Trim()
      Isp1Fio = (Eng-Fio $e1); Isp1Dolzh = $(if ($e1) { [string]$e1.Dolzh } else { '' }); Isp1Init = $(if ($e1) { [string]$e1.Init } else { '' })
      Isp2Fio = (Eng-Fio $e2); Isp2Dolzh = $(if ($e2) { [string]$e2.Dolzh } else { '' }); Isp2Init = $(if ($e2) { [string]$e2.Init } else { '' })
      ZakFio1 = $tbZakFio1.Text.Trim(); ZakDolzh1 = $tbZakDolzh1.Text.Trim()
      ZakFio2 = $tbZakFio2.Text.Trim(); ZakDolzh2 = $tbZakDolzh2.Text.Trim()
    }

    $lblStatus.Text = 'Формирование... (Word работает в фоне)'
    $form.Enabled = $false
    [System.Windows.Forms.Application]::DoEvents()

    $res = @(& (Join-Path $ScriptDir 'Generate.ps1') -Data $data)[-1]

    $script:Settings.TpDob = $tbTpDob.Text.Trim()
    $script:Settings.LastEngineer = [string]$cmbEng.Text
    Save-Customer
    Save-Settings $script:Settings
    Refresh-Customers

    $files = @($res.Files)
    $lblStatus.Text = ('Готово: {0} файл(ов) в папке {1}' -f $files.Count, (Split-Path -Leaf $res.Folder))
    $text = ("Создано документов: {0}`n`n{1}`n`nПапка: {2}" -f $files.Count, (($files | ForEach-Object { [IO.Path]::GetFileName($_) }) -join "`n"), $res.Folder)
    if (@($res.Warnings).Count -gt 0) { $text += "`n`nВнимание:`n" + (@($res.Warnings) -join "`n") }
    $a = [System.Windows.Forms.MessageBox]::Show($text + "`n`nОткрыть папку?", 'Готово', 'YesNo', 'Information')
    if ($a -eq 'Yes') { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $res.Folder) }
    if ($chkAutoClear.Checked) { $form.Enabled = $true; $btnClear.PerformClick(); $lblStatus.Text = 'Готово. Форма очищена - можно вводить следующий акт.' }
  } catch {
    $lblStatus.Text = 'Ошибка'
    Msg $_.Exception.Message 'Ошибка' 'Error'
  } finally {
    $form.Enabled = $true
    if ($script:FailCtrl) {
      $fc = $script:FailCtrl; $script:FailCtrl = $null
      try { [void]$fc.Focus() } catch { }
    }
  }
})

# =====================================================================
#  Старт
# =====================================================================
Refresh-EngineerCombos
if ($script:Settings.LastEngineer -and (Get-Engineer $script:Settings $script:Settings.LastEngineer)) { $cmbEng.Text = $script:Settings.LastEngineer }
Apply-Engineer
$v1 = Get-Engineer $script:Settings 'Викторова Е.Г.'
if ($v1) { $cmbIsp1.Text = 'Викторова Е.Г.' }
$v2 = Get-Engineer $script:Settings 'Смирнов И.В.'
if ($v2) { $cmbIsp2.Text = 'Смирнов И.В.' }
Refresh-Customers
Refresh-ProductList
Update-Layout
$chkSns.Checked = $true

# --- мелкие удобства ---
# номер договора: курсор всегда в конец (чтобы не затереть «СТКБ-» и вводить только цифры)
$tbNomer.Add_Enter({
  # переменная скрипта $tbNomer доступна и позже, когда сработает отложенный вызов
  # (локальная $tb к тому моменту уже исчезла - отсюда была ошибка "variable '$tb' cannot be retrieved")
  [void]$tbNomer.BeginInvoke([Action]{ $tbNomer.SelectionStart = $tbNomer.Text.Length; $tbNomer.SelectionLength = 0 })
})
# при выходе из поля убираем двойные пробелы, «г.Москва» -> «г. Москва» и т.п.
foreach ($c in @($cmbZak, $cmbGorod, $cmbGorodIsp, $tbAdres, $tbIzpDolzh, $tbZakFio1, $tbZakDolzh1, $tbZakFio2, $tbZakDolzh2)) {
  $c.Add_Leave({
    $n = Normalize-Text $this.Text
    if ($n -ne $this.Text) { $this.Text = $n }
  })
}
# Ctrl+Enter = «Сформировать»
$form.KeyPreview = $true
$form.Add_KeyDown({
  if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
    $_.SuppressKeyPress = $true
    if ($btnGen.Enabled) { $btnGen.PerformClick() }
  }
})
# подсказки
$tip = New-Object System.Windows.Forms.ToolTip
$tip.AutoPopDelay = 8000
$tip.SetToolTip($btnGen, 'Ctrl+Enter')
$tip.SetToolTip($cmbZak, 'Выберите из списка: адрес, город и представители подставятся сами')
$tip.SetToolTip($tbNomer, 'Достаточно дописать цифры после «СТКБ-»')
$tip.SetToolTip($tbSnsId, 'Только последние 6 цифр - префикс РОСС RU.01.03745. добавится сам')
$tip.SetToolTip($tbDlId, 'Только последние 6 цифр - префикс добавится сам')
$tip.SetToolTip($tbVipSerial, 'Только цифры после КлКС2-4-; регистрационный номер посчитается сам')
$tip.SetToolTip($chkAutoClear, 'После создания документов форма очистится (заказчик останется в списке)')
$tip.SetToolTip($dgv, 'Одна строка - пункты 6.3-6.5. Несколько строк - таблица')

[void]$form.ShowDialog()
