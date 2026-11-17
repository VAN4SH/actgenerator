#Requires -Version 5.1
<#
.SYNOPSIS
  Формирует акты из шаблонов Word (папки берутся из settings.json).
.PARAMETER Data
  Hashtable с полями формы (см. Form.ps1).
.PARAMETER OpenFolder
  Открыть папку с готовыми документами.
.OUTPUTS
  Hashtable: Files (массив путей), Folder, Warnings.
#>
param(
  [Parameter(Mandatory = $true)]
  [hashtable]$Data,
  [switch]$OpenFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ScriptDir 'Common.ps1')

$Data = Normalize-Value $Data

$settings = Load-Settings
$TplDir = $settings.TemplatesDir
$OutRoot = $settings.OutDir
$missing = @(Get-MissingTemplates $TplDir)
if ($missing.Count -gt 0) {
  throw ("В папке шаблонов нет файлов: {0}`nПапка: {1}" -f ($missing -join ', '), $TplDir)
}
if ([string]::IsNullOrWhiteSpace($OutRoot)) { throw 'Не задана папка для готовых актов (кнопка «Настройки»).' }

$script:Warnings = New-Object System.Collections.Generic.List[string]

# ---------------- доступ к данным ----------------
function D([string]$k, $def = '') {
  if ($Data.ContainsKey($k) -and $null -ne $Data[$k]) { return $Data[$k] }
  return $def
}
function DS([string]$k) { return ([string](D $k '')).Trim() }
function DB([string]$k) { $v = D $k $false; return ($v -eq $true -or ([string]$v) -eq 'True') }

# ---------------- работа с Word ----------------
function PT($p) { return ($p.Range.Text -replace "[\r\a\f]", '') }

function Set-Ph($rng, [string]$key, [string]$val) {
  for ($g = 0; $g -lt 300; $g++) {
    $r = $rng.Duplicate
    $f = $r.Find
    $f.ClearFormatting()
    $f.Text = $key
    $f.Forward = $true
    $f.Wrap = 0
    $f.MatchCase = $true
    $f.MatchWildcards = $false
    if (-not $f.Execute()) { break }
    $r.Text = $val
  }
}

function Replace-Map($rng, [hashtable]$map) {
  $keys = $map.Keys | Sort-Object { $_.Length } -Descending
  foreach ($key in $keys) {
    $v = $map[$key]
    if ($null -eq $v) { $v = '' }
    Set-Ph $rng $key ([string]$v)
  }
}

function Remove-Bm($doc, [string]$name) {
  if (-not $doc.Bookmarks.Exists($name)) { return }
  $r = $doc.Bookmarks.Item($name).Range
  $tc = $r.Tables.Count
  for ($k = $tc; $k -ge 1; $k--) { $r.Tables.Item($k).Delete() }
  if ($doc.Bookmarks.Exists($name)) {
    $null = $doc.Bookmarks.Item($name).Range.Delete()
  }
}

function Set-RangeTextNoMark($rng, [string]$text) {
  $r = $rng.Duplicate
  $r.End = $r.End - 1
  $r.Text = $text
}

function Resolve-OutFile([string]$folder, [string]$name) {
  $path = Join-Path $folder $name
  if (Test-Path -LiteralPath $path) {
    try { Remove-Item -LiteralPath $path -Force }
    catch {
      $stamp = Get-Date -Format 'HHmmss'
      $base = [IO.Path]::GetFileNameWithoutExtension($name)
      $path = Join-Path $folder ("{0} ({1}).docx" -f $base, $stamp)
      $script:Warnings.Add(("Файл «{0}» занят (открыт в Word?), создан новый: {1}" -f $name, [IO.Path]::GetFileName($path)))
    }
  }
  return $path
}

function New-FromTemplate([string]$templateName, [string]$destPath, [scriptblock]$edit) {
  $src = Join-Path $TplDir $templateName
  Copy-Item -LiteralPath $src -Destination $destPath -Force
  $doc = $script:WordApp.Documents.Open($destPath)
  try {
    & $edit $doc
    $left = @([regex]::Matches($doc.Content.Text, '\{\{[A-Z0-9_]+\}\}') | ForEach-Object { $_.Value } | Sort-Object -Unique)
    if ($left.Count -gt 0) {
      $script:Warnings.Add(("В файле {0} остались незаполненные поля: {1}" -f [IO.Path]::GetFileName($destPath), ($left -join ', ')))
    }
    $null = $doc.Save()
  } finally {
    $null = $doc.Close($false)
  }
}

# ---------------- общие значения ----------------
$nomer = DS 'Nomer'
$zak = (Convert-Quotes (DS 'Zakazchik')).TrimEnd('.').Trim()
$dataAkta = Parse-DateFlexible (DS 'DataAkta')
if ($null -eq $dataAkta) { throw 'Не указана дата акта установки.' }
$dataDog = Parse-DateFlexible (DS 'DataDogovora')
$mode = DS 'Mode'
if ($mode -eq '') { $mode = 'Dogovor' }
$izpDolzh = DS 'IzpDolzh'
$izpInit = DS 'IzpInit'

function Base-Map {
  $dd = ''
  if ($null -ne $dataDog) { $dd = $dataDog.ToString('dd.MM.yyyy') }
  $ot = DS 'OsnovanieTip'
  if ($ot -eq '') { $ot = 'Договор' }
  return @{
    '{{NOMER}}'          = $nomer
    '{{GOROD}}'          = (DS 'Gorod')
    '{{ZAKAZCHIK}}'      = $zak
    '{{ADRES}}'          = (DS 'Adres')
    '{{DATA_DOGOVORA}}'  = $dd
    '{{OSNOVANIE_TIP}}'  = $ot
    '{{DATA_AKTA_FULL}}' = (Format-RusDate $dataAkta)
    '{{IZP_DOLZH}}'      = $izpDolzh
    '{{IZP_INIT}}'       = $izpInit
  }
}

# ---------------- акт по договору (SNS / Dallas / ViPNet / КриптоПро / TLS-клиент) ----------------
function Fill-ActDogovor($doc) {
  $map = Base-Map

  if (-not (DB 'HasDl')) { Remove-Bm $doc 'blk_dl' }
  if (-not (DB 'HasSns')) { Remove-Bm $doc 'blk_sns' }
  if (-not (DB 'HasTls')) { Remove-Bm $doc 'blk_tls' }
  if (-not (DB 'HasVip')) { Remove-Bm $doc 'blk_vip' }
  if (-not (DB 'HasKp')) { Remove-Bm $doc 'blk_kp' }

  # модули SNS
  if (DB 'HasSns') {
    $selected = @(D 'SnsModules' @())
    $keys = @('nsd', 'dev', 'disk', 'fw', 'ids', 'av')
    foreach ($k in $keys) {
      if ($selected -notcontains $k) { Remove-Bm $doc ('mod_' + $k) }
    }
    $remain = @()
    foreach ($k in $keys) { if ($doc.Bookmarks.Exists('mod_' + $k)) { $remain += ('mod_' + $k) } }
    for ($i = 0; $i -lt $remain.Count; $i++) {
      $rng = $doc.Bookmarks.Item($remain[$i]).Range
      $t = ($rng.Text -replace "[\r\a\f]", '').TrimEnd(';', '.', ' ')
      $end = ';'
      if ($i -eq $remain.Count - 1) { $end = '.' }
      Set-RangeTextNoMark $rng ($t + $end)
    }
  }

  # рабочие места
  $wps = @(D 'Workplaces' @())
  $dtUstDefault = $dataAkta
  $dUst = Parse-DateFlexible (DS 'DataUstanovki')
  if ($null -ne $dUst) { $dtUstDefault = $dUst }
  $room1 = ''; $pc1 = ''; $dateUst1 = $dtUstDefault
  if ($wps.Count -gt 0) {
    $room1 = [string]$wps[0].Room
    $pc1 = [string]$wps[0].Pc
    $d1 = $null
    if ($wps[0].ContainsKey('Date') -and -not [string]::IsNullOrWhiteSpace([string]$wps[0].Date)) { $d1 = Parse-DateFlexible ([string]$wps[0].Date) }
    if ($null -ne $d1) { $dateUst1 = $d1 }
  }
  if ($wps.Count -le 1) {
    Remove-Bm $doc 'loc_table'
  } else {
    Remove-Bm $doc 'loc_lines'
    $tbl = $doc.Bookmarks.Item('loc_table').Range.Tables.Item(1)
    for ($i = 0; $i -lt $wps.Count; $i++) {
      $rowIdx = $i + 2
      if ($rowIdx -gt $tbl.Rows.Count) { $null = $tbl.Rows.Add() }
      $dRow = $dtUstDefault
      if ($wps[$i].ContainsKey('Date') -and -not [string]::IsNullOrWhiteSpace([string]$wps[$i].Date)) { $dRow = Parse-DateFlexible ([string]$wps[$i].Date) }
      $tbl.Cell($rowIdx, 1).Range.Text = [string]($i + 1)
      $tbl.Cell($rowIdx, 2).Range.Text = [string]$wps[$i].Room
      $tbl.Cell($rowIdx, 3).Range.Text = [string]$wps[$i].Pc
      $tbl.Cell($rowIdx, 4).Range.Text = $dRow.ToString('dd.MM.yyyy')
    }
  }
  $map['{{POMESHCHENIE}}'] = $room1
  $map['{{KOMPYUTER}}'] = $pc1
  $map['{{DATA_UST_FULL}}'] = (Format-RusDate $dateUst1)

  # продукты
  $snsSuffix = ''
  if (DB 'Sns30') { $snsSuffix = ' 30' }
  $map['{{SNS_SUFFIX}}'] = $snsSuffix
  $map['{{SNS_IDENT}}'] = (Compose-Id 'SNS' (DS 'SnsIdent'))
  $map['{{DL_IDENT}}'] = (Compose-Id 'DL' (DS 'DlIdent'))
  $map['{{KP_IDENT}}'] = (Compose-Id 'KP' (DS 'KpIdent'))
  $map['{{TLS_SERIAL}}'] = (DS 'TlsSerial')
  $map['{{TLS_REG}}'] = (Compose-Id 'TLSREG' (DS 'TlsReg'))
  $vipKs = Set-VipSnPrefix (DS 'VipKs')
  $vipSerial = Compose-Id 'VIPSN' (DS 'VipSerial')
  $vipReg = DS 'VipReg'
  if ($vipReg -eq '') { $vipReg = Get-VipReg $vipSerial }
  elseif ($vipReg -match '^\d+$') { $vipReg = '782-' + $vipReg }
  $map['{{VIPNET_SERIAL}}'] = $vipSerial
  $map['{{VIPNET_REG}}'] = $vipReg
  $vt = DS 'VipTitle'
  if ($vt -eq '') { $vt = ('Программный комплекс ViPNet Client for Windows 4.х ({0})' -f $vipKs) }
  $map['{{VIPNET_TITLE}}'] = $vt.TrimEnd(':')

  Replace-Map $doc.Content $map
}

# ---------------- акт по контракту (Prime / Coordinator / Континент) ----------------
function Fill-ActKontrakt($doc) {
  $map = Base-Map
  $split = Split-Dolzh $izpDolzh
  $map['{{IZP_L1}}'] = $split[0]
  $map['{{IZP_L2}}'] = $split[1]
  $map['{{IZP_INIT_SP}}'] = (Get-InitSp $izpInit)

  $protoKeys = @{ 'Prime' = 'pr_prime'; 'Coordinator' = 'pr_coord'; 'Continent' = 'pr_cont'; 'ContinentTLS' = 'pr_tls' }
  $bm = @{}
  foreach ($k in $protoKeys.Values) {
    $r = $doc.Bookmarks.Item($k).Range
    $bm[$k] = @($r.Start, $r.End)
  }
  $firstStart = $bm['pr_prime'][0]
  $lastEnd = $bm['pr_tls'][1]
  $pos = $lastEnd

  $products = @(D 'Products' @())
  if ($products.Count -eq 0) { throw 'В акте по контракту нет ни одной позиции СЗИ.' }
  $adresDefault = DS 'Adres'

  foreach ($p in $products) {
    $type = [string]$p.Type
    if (-not $protoKeys.ContainsKey($type)) { throw "Неизвестный тип позиции: $type" }
    $pr = $bm[$protoKeys[$type]]
    $len = $pr[1] - $pr[0]
    $ins = $doc.Range($pos, $pos)
    $ins.FormattedText = $doc.Range($pr[0], $pr[1]).FormattedText
    $copy = $doc.Range($pos, $pos + $len)

    # необязательные строки
    if ($type -eq 'Continent') {
      $kill = @()
      foreach ($pp in $copy.Paragraphs) {
        $t = $pp.Range.Text
        if ([string]::IsNullOrWhiteSpace([string]$p.Ident) -and $t -match '\{\{P_IDENT\}\}') { $kill += $pp.Range }
        elseif ([string]::IsNullOrWhiteSpace([string]$p.RegSkzi) -and $t -match '\{\{P_REG\}\}') { $kill += $pp.Range }
        elseif (-not [bool]$p.HasFsbCert -and $t -match 'СФ/124-4804') { $kill += $pp.Range }
      }
      for ($i = $kill.Count - 1; $i -ge 0; $i--) { $null = $kill[$i].Delete() }
    }

    $addr = [string]$p.Address
    if ([string]::IsNullOrWhiteSpace($addr)) { $addr = $adresDefault }
    $name = [string]$p.Name
    if ([string]::IsNullOrWhiteSpace($name)) {
      if ($type -eq 'Prime') { $name = 'ViPNet Prime' } elseif ($type -eq 'Coordinator') { $name = 'ViPNet Coordinator HW 1000 C' }
    }
    $dInst = $dataAkta
    if (-not [string]::IsNullOrWhiteSpace([string]$p.InstallDate)) { $dInst = Parse-DateFlexible ([string]$p.InstallDate) }
    $pm = @{
      '{{P_NAME}}'     = $name
      '{{P_SERIAL}}'   = [string]$p.Serial
      '{{P_PLATFORM}}' = [string]$p.Platform
      '{{P_REG}}'      = [string]$p.RegSkzi
      '{{P_IDENTSZI}}' = [string]$p.Identifikator
      '{{P_IDENT}}'    = [string]$p.Ident
      '{{P_FACTORY}}'  = [string]$p.Factory
      '{{P_ORG}}'      = $zak
      '{{P_ADDR}}'     = $addr
      '{{P_ROOM}}'     = [string]$p.Room
      '{{P_PC}}'       = [string]$p.Computer
      '{{P_DATE}}'     = (Format-RusDate $dInst)
    }
    Replace-Map $copy $pm
    $pos = $copy.End
  }

  # удалить прототипы
  $null = $doc.Range($firstStart, $lastEnd).Delete()
  Replace-Map $doc.Content $map
}

# ---------------- сертификат ТП ----------------
function Fill-Tp($doc) {
  $dTp = Parse-DateFlexible (DS 'DataTp')
  if ($null -eq $dTp) { $dTp = $dataAkta }
  $nomerSert = DS 'NomerSert'
  if ($nomerSert -eq '') { $nomerSert = '{0}/{1}' -f $nomer, $dTp.ToString('ddMMyyyy') }
  $rod = Convert-Quotes (DS 'ZakazchikRod')
  if ($rod -eq '') { $rod = Get-Rod $zak }
  $prods = @(D 'TpProducts' @() | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
  if ($prods.Count -eq 0) { $prods = @('Средство защиты информации Secret Net Studio') }
  $dob = DS 'TpDob'
  if ($dob -eq '') { $dob = '836' }

  $map = @{
    '{{DATA_TP_FULL}}'  = (Format-RusDate $dTp)
    '{{NOMER_SERT}}'    = $nomerSert
    '{{ZAKAZCHIK_ROD}}' = $rod
    '{{TP_DOB}}'        = $dob
    '{{PRODUKT}}'       = [string]$prods[0]
  }
  $tbl = $null
  foreach ($t in $doc.Tables) { if ($t.Columns.Count -eq 3 -and $t.Range.Text -match '\{\{PRODUKT\}\}') { $tbl = $t } }
  Replace-Map $doc.Content $map
  if ($null -ne $tbl -and $prods.Count -gt 1) {
    for ($i = 1; $i -lt $prods.Count; $i++) {
      $row = $tbl.Rows.Add()
      $ri = $row.Index
      $tbl.Cell($ri, 2).Range.Text = [string]$prods[$i]
      $tbl.Cell($ri, 3).Range.Text = '1'
    }
  }
}

# ---------------- акт испытаний ----------------
function Fill-Isp($doc) {
  $dIsp = Parse-DateFlexible (DS 'DataIsp')
  if ($null -eq $dIsp) { $dIsp = $dataAkta }
  $hasSns = DB 'HasSns'; $hasDl = DB 'HasDl'; $hasVip = DB 'HasVip'; $hasTls = DB 'HasTls'
  $selected = @(D 'SnsModules' @())
  $hasAv = ($hasSns -and ($selected -contains 'av'))

  if (-not ($hasSns -or $hasDl)) { Remove-Bm $doc 'isp_nsd_act'; Remove-Bm $doc 'isp_nsd_find' }
  if (-not $hasAv) { Remove-Bm $doc 'isp_av_act'; Remove-Bm $doc 'isp_av_find' }
  if (-not ($hasVip -or $hasTls)) { Remove-Bm $doc 'isp_net_find' }

  $lq = [string][char]0xAB; $rq = [string][char]0xBB
  $nsd = @()
  if ($hasSns) { $nsd += 'Средство защиты информации Secret Net Studio' }
  if ($hasDl) { $nsd += ('Система защиты информации от несанкционированного доступа ' + $lq + 'Dallas Lock 8.0-К' + $rq) }
  $net = @()
  if ($hasVip) { $net += 'программный комплекс ViPNet Client 4' }
  if ($hasTls) { $net += ('программный комплекс ' + $lq + 'Континент TLS Клиент' + $rq) }

  $wps = @(D 'Workplaces' @())
  $rooms = @(); $pcs = @()
  foreach ($w in $wps) {
    if (-not [string]::IsNullOrWhiteSpace([string]$w.Room)) { $rooms += [string]$w.Room }
    if (-not [string]::IsNullOrWhiteSpace([string]$w.Pc)) { $pcs += [string]$w.Pc }
  }

  $dd = ''
  if ($null -ne $dataDog) { $dd = $dataDog.ToString('dd.MM.yyyy') }
  $map = @{
    '{{GOROD_ISP}}'     = (DS 'GorodIsp')
    '{{DATA_ISP_FULL}}' = (Format-RusDate $dIsp)
    '{{NOMER}}'         = $nomer
    '{{DATA_DOGOVORA}}' = $dd
    '{{ZAKAZCHIK}}'     = $zak
    '{{ADRES}}'         = (DS 'Adres')
    '{{POMESHCHENIE}}'  = ($rooms -join '; ')
    '{{KOMPYUTER}}'     = ($pcs -join '; ')
    '{{ISP_NSD_PROD}}'  = ($nsd -join ' и ')
    '{{ISP_NET_PROD}}'  = ($net -join ' и ')
    '{{ISP_FIO_1}}'     = (DS 'Isp1Fio')
    '{{ISP_DOLZH_1}}'   = (DS 'Isp1Dolzh')
    '{{ISP_INIT_1}}'    = (Get-InitSp (DS 'Isp1Init'))
    '{{ISP_FIO_2}}'     = (DS 'Isp2Fio')
    '{{ISP_DOLZH_2}}'   = (DS 'Isp2Dolzh')
    '{{ISP_INIT_2}}'    = (Get-InitSp (DS 'Isp2Init'))
    '{{ZAK_FIO_1}}'     = (DS 'ZakFio1')
    '{{ZAK_DOLZH_1}}'   = (DS 'ZakDolzh1')
    '{{ZAK_INIT_1}}'    = (Get-InitFromFio (DS 'ZakFio1') $true)
    '{{ZAK_FIO_2}}'     = (DS 'ZakFio2')
    '{{ZAK_DOLZH_2}}'   = (DS 'ZakDolzh2')
    '{{ZAK_INIT_2}}'    = (Get-InitFromFio (DS 'ZakFio2') $true)
  }
  Replace-Map $doc.Content $map
}

# ---------------- main ----------------
$folderName = Sanitize-FileName ('{0} {1}' -f $nomer, (Short-Org $zak))
$outDir = Join-Path $OutRoot $folderName
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$script:WordApp = New-Object -ComObject Word.Application
$script:WordApp.Visible = $false
$script:WordApp.DisplayAlerts = 0
$created = New-Object System.Collections.Generic.List[string]
$nomerFile = Sanitize-FileName $nomer

try {
  $actPath = Resolve-OutFile $outDir ("Акт установки {0}.docx" -f $nomerFile)
  if ($mode -eq 'Kontrakt') {
    New-FromTemplate 'Акт_установки_Составной.docx' $actPath { param($doc) Fill-ActKontrakt $doc }
  } else {
    New-FromTemplate 'Акт_установки_Договор.docx' $actPath { param($doc) Fill-ActDogovor $doc }
  }
  $created.Add($actPath)

  if (DB 'NeedTp') {
    $tpPath = Resolve-OutFile $outDir ("Сертификат ТП {0}.docx" -f $nomerFile)
    New-FromTemplate 'Сертификат_ТП.docx' $tpPath { param($doc) Fill-Tp $doc }
    $created.Add($tpPath)
  }

  if ((DB 'NeedIsp') -and $mode -ne 'Kontrakt') {
    $ispPath = Resolve-OutFile $outDir ("Акт испытаний {0}.docx" -f $nomerFile)
    New-FromTemplate 'Акт_испытаний.docx' $ispPath { param($doc) Fill-Isp $doc }
    $created.Add($ispPath)
  }
} finally {
  $script:WordApp.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:WordApp) | Out-Null
  $script:WordApp = $null
}

if ($OpenFolder) { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $outDir) }

return @{ Files = @($created); Folder = $outDir; Warnings = @($script:Warnings) }
