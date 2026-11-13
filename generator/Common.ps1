# Общие функции: настройки, пути, автоподстановки. Подключается из Form.ps1 и Generate.ps1 (dot-source).

$script:CommonDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsPath = Join-Path $script:CommonDir 'settings.json'

$script:RequiredTemplates = @(
  'Акт_установки_Договор.docx',
  'Акт_установки_Составной.docx',
  'Сертификат_ТП.docx',
  'Акт_испытаний.docx'
)

# Префиксы номеров: [префикс, дополнить нулями до N цифр (0 = не дополнять)]
$script:IdRules = @{
  SNS    = @('РОСС RU.01.03745.', 0)
  DL     = @('РОСС RU.0001.02720.', 0)
  KP     = @('385Б1А1-', 6)
  TLSREG = @('1073Б1-', 6)
  VIPSN  = @('КлКС2-4-', 0)
}
$script:VipRegOffset = 971

$script:Months = @{
  1 = 'января'; 2 = 'февраля'; 3 = 'марта'; 4 = 'апреля'; 5 = 'мая'; 6 = 'июня'
  7 = 'июля'; 8 = 'августа'; 9 = 'сентября'; 10 = 'октября'; 11 = 'ноября'; 12 = 'декабря'
}

function Format-RusDate([datetime]$dt) {
  $lq = [string][char]0xAB
  $rq = [string][char]0xBB
  return ('{0}{1}{2} {3} {4}' -f $lq, $dt.Day, $rq, $script:Months[[int]$dt.Month], $dt.Year)
}

function Parse-DateFlexible([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $dto = [datetime]::MinValue
  foreach ($c in @([cultureinfo]::GetCultureInfo('ru-RU'), [cultureinfo]::InvariantCulture)) {
    if ([datetime]::TryParse($s, $c, [System.Globalization.DateTimeStyles]::None, [ref]$dto)) { return $dto }
  }
  throw "Не удалось разобрать дату: $s"
}

# Составляет полный номер: если введены только цифры - добавляет префикс; иначе берёт как есть.
function Compose-Id([string]$kind, [string]$value) {
  $v = ([string]$value).Trim()
  if ($v -eq '') { return '' }
  $rule = $script:IdRules[$kind]
  if ($v -match '^\d+$') {
    if ($rule[1] -gt 0) { $v = $v.PadLeft([int]$rule[1], '0') }
    return ($rule[0] + $v)
  }
  return $v
}

# Рег. номер СКЗИ ViPNet Client = 782-(серийный номер + 971). Пустая строка, если не получилось.
function Get-VipReg([string]$serialFull) {
  $m = [regex]::Match(([string]$serialFull).Trim(), '(\d+)$')
  if (-not $m.Success) { return '' }
  $n = [int64]$m.Groups[1].Value + $script:VipRegOffset
  return ('782-' + $n.ToString())
}

# 'П.А. Скворцов' -> 'П. А. Скворцов'
function Get-InitSp([string]$init) {
  return ([regex]::Replace(([string]$init).Trim(), '(?<=\.)(?=[А-ЯЁ])', ' '))
}

# 'Теребенцева Елена Леонидовна' -> 'Е. Л. Теребенцева' (spaced) / 'Е.Л. Теребенцева'
function Get-InitFromFio([string]$fio, [bool]$spaced = $true) {
  $parts = @(([string]$fio).Trim() -split '\s+' | Where-Object { $_ })
  if ($parts.Count -eq 0) { return '' }
  if ($parts.Count -eq 1) { return $parts[0] }
  $ini = @()
  for ($i = 1; $i -lt $parts.Count; $i++) { $ini += ($parts[$i].Substring(0, 1).ToUpper() + '.') }
  $sep = ''
  if ($spaced) { $sep = ' ' }
  return (($ini -join $sep) + ' ' + $parts[0])
}

# 'П.А. Скворцов' -> 'Скворцов П.А.' (для списков комиссии, если нет полного ФИО)
function Get-FioFromInit([string]$init) {
  $m = [regex]::Match(([string]$init).Trim(), '^((?:[А-ЯЁ]\.\s?)+)\s*(.+)$')
  if (-not $m.Success) { return $init }
  return ($m.Groups[2].Value.Trim() + ' ' + $m.Groups[1].Value.Trim())
}

# Должность в две строки (как в актах по контракту): 'Инженер' / 'по защите информации 1 категории'
function Split-Dolzh([string]$d) {
  $t = ([string]$d).Trim()
  $m = [regex]::Match($t, '^(.*?)\s+(по защите информации.*)$')
  if ($m.Success) { return @($m.Groups[1].Value, $m.Groups[2].Value) }
  $i = $t.IndexOf(' ')
  if ($i -gt 0) { return @($t.Substring(0, $i), $t.Substring($i + 1)) }
  return @($t, '')
}

# Заказчик в дательном падеже (для сертификата ТП) - простые правила, можно поправить вручную.
function Get-Rod([string]$name) {
  $n = ([string]$name).Trim()
  if ($n -eq '') { return '' }
  $map = @{
    'муниципальное' = 'муниципальному'; 'муниципальный' = 'муниципальному'
    'казенное' = 'казенному'; 'казённое' = 'казённому'; 'бюджетное' = 'бюджетному'
    'автономное' = 'автономному'; 'государственное' = 'государственному'
    'дошкольное' = 'дошкольному'; 'образовательное' = 'образовательному'
    'общеобразовательное' = 'общеобразовательному'; 'учреждение' = 'учреждению'
    'общество' = 'обществу'; 'министерство' = 'министерству'; 'администрация' = 'администрации'
    'управление' = 'управлению'; 'предприятие' = 'предприятию'; 'центр' = 'центру'
    'комитет' = 'комитету'; 'отдел' = 'отделу'; 'школа' = 'школе'
  }
  $tokens = $n -split ' '
  $out = New-Object System.Collections.Generic.List[string]
  $stop = $false
  foreach ($tk in $tokens) {
    if (-not $stop -and ($tk -match '[«"“]' -or $tk -match '\d')) { $stop = $true }
    if (-not $stop) {
      $key = $tk.ToLower()
      if ($map.ContainsKey($key)) { $out.Add($map[$key]); continue }
    }
    $out.Add($tk)
  }
  $res = ($out -join ' ')
  # аббревиатуры (ООО, АО, ГБУ ...) не склоняем и с маленькой буквы не пишем
  $first = [string]$tokens[0]
  $isAbbr = ($first.Length -ge 2 -and $first -cmatch '^[А-ЯЁA-Z]+$')
  if ($res.Length -gt 0 -and -not $isAbbr) { $res = $res.Substring(0, 1).ToLower() + $res.Substring(1) }
  return $res
}

# Прямые кавычки "текст" -> «текст» (как в оригинальных актах)
function Convert-Quotes([string]$s) {
  if ($null -eq $s) { return '' }
  $lq = [string][char]0xAB
  $rq = [string][char]0xBB
  $t = [regex]::Replace($s, '"([^"]*)"', ($lq + '$1' + $rq))
  $t = [regex]::Replace($t, '[“]([^”]*)[”]', ($lq + '$1' + $rq))
  return $t
}

# Аккуратная чистка введённого текста: двойные пробелы, пробел перед запятой,
# "г.Москва" -> "г. Москва", "ул.Ленина" -> "ул. Ленина", нет пробела после запятой.
function Normalize-Text([string]$s) {
  if ($null -eq $s) { return '' }
  $t = $s -replace '[ \t]{2,}', ' '
  $t = [regex]::Replace($t, '\s+([,;])', '$1')
  $t = [regex]::Replace($t, ',(?=[^\s\d,])', ', ')
  $t = [regex]::Replace($t, '(?<![А-Яа-яЁёA-Za-z])(г|ул|д|с|пер)\.(?=[А-ЯЁA-Z0-9])', '$1. ')
  return $t.Trim()
}

# Применяет Normalize-Text ко всем строкам данных формы (в т.ч. внутри списков и таблиц).
function Normalize-Value($v) {
  if ($null -eq $v) { return $null }
  if ($v -is [string]) { return (Normalize-Text $v) }
  if ($v -is [hashtable]) {
    $h = @{}
    foreach ($k in @($v.Keys)) { $h[$k] = Normalize-Value $v[$k] }
    return $h
  }
  if ($v -is [System.Array]) {
    $list = @()
    foreach ($x in $v) { $list += , (Normalize-Value $x) }
    return $list
  }
  if ($v -is [pscustomobject]) {
    $o = [ordered]@{}
    foreach ($p in $v.PSObject.Properties) { $o[$p.Name] = Normalize-Value $p.Value }
    return [pscustomobject]$o
  }
  return $v
}

function Sanitize-FileName([string]$name) {
  $invalid = [IO.Path]::GetInvalidFileNameChars()
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $name.ToCharArray()) {
    if ($invalid -contains $ch) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) }
  }
  $result = $sb.ToString().Trim().TrimEnd('.')
  if ([string]::IsNullOrWhiteSpace($result)) { return 'document' }
  return $result
}

function Short-Org([string]$org, [int]$max = 40) {
  if ([string]::IsNullOrWhiteSpace($org)) { return '' }
  $s = ($org -replace '[«»"]', '').Trim()
  if ($s.Length -gt $max) {
    $cut = $s.Substring(0, $max)
    $sp = $cut.LastIndexOf(' ')
    if ($sp -gt 15) { $cut = $cut.Substring(0, $sp) }
    $s = $cut.Trim()
  }
  return $s
}

# ---------------- Настройки ----------------

function Get-DefaultEngineers {
  return @(
    [pscustomobject]@{ Name = 'Скворцов П.А.'; Fio = ''; Dolzh = 'Инженер по защите информации'; Init = 'П.А. Скворцов' },
    [pscustomobject]@{ Name = 'Смирнов И.В.'; Fio = 'Смирнов Иван Вадимович'; Dolzh = 'Старший инженер по защите информации'; Init = 'И.В. Смирнов' },
    [pscustomobject]@{ Name = 'Кочетков А.А.'; Fio = ''; Dolzh = 'Начальник отдела внедрения'; Init = 'А.А. Кочетков' },
    [pscustomobject]@{ Name = 'Седов С.Ю.'; Fio = ''; Dolzh = 'Заместитель начальника отдела аудита информационных систем'; Init = 'С.Ю. Седов' },
    [pscustomobject]@{ Name = 'Викторова Е.Г.'; Fio = 'Викторова Екатерина Германовна'; Dolzh = 'Заместитель начальника отдела аудита информационных систем'; Init = 'Е.Г. Викторова' }
  )
}

function New-Settings {
  return [pscustomobject]@{
    TemplatesDir = ''
    OutDir       = ''
    TpDob        = '836'
    LastEngineer = 'Скворцов П.А.'
    Engineers    = @(Get-DefaultEngineers)
    Customers    = @()
  }
}

function Load-Settings {
  $s = New-Settings
  if (Test-Path -LiteralPath $script:SettingsPath) {
    try {
      $raw = [IO.File]::ReadAllText($script:SettingsPath, [Text.Encoding]::UTF8)
      $j = $raw | ConvertFrom-Json
      foreach ($name in 'TemplatesDir', 'OutDir', 'TpDob', 'LastEngineer') {
        $prop = $j.PSObject.Properties[$name]
        if ($null -ne $prop -and $null -ne $prop.Value) { $s.$name = [string]$prop.Value }
      }
      $eng = @()
      if ($null -ne $j.PSObject.Properties['Engineers'] -and $null -ne $j.Engineers) { $eng = @($j.Engineers) }
      if ($eng.Count -gt 0) { $s.Engineers = $eng }
      if ($null -ne $j.PSObject.Properties['Customers'] -and $null -ne $j.Customers) { $s.Customers = @($j.Customers) }
    } catch {
      Write-Warning "Не удалось прочитать settings.json: $($_.Exception.Message)"
    }
  }
  return $s
}

function Save-Settings($s) {
  $json = $s | ConvertTo-Json -Depth 6
  [IO.File]::WriteAllText($script:SettingsPath, $json, (New-Object System.Text.UTF8Encoding($true)))
}

function Test-TemplatesDir([string]$dir) {
  if ([string]::IsNullOrWhiteSpace($dir)) { return $false }
  if (-not (Test-Path -LiteralPath $dir)) { return $false }
  foreach ($n in $script:RequiredTemplates) {
    if (-not (Test-Path -LiteralPath (Join-Path $dir $n))) { return $false }
  }
  return $true
}

function Get-MissingTemplates([string]$dir) {
  $missing = @()
  foreach ($n in $script:RequiredTemplates) {
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath (Join-Path $dir $n))) { $missing += $n }
  }
  return $missing
}

# Ищет папку «Шаблоны» рядом со скриптами (до 3 уровней вверх, и вглубь до 2 уровней)
function Find-DefaultFolders {
  $res = @{ Templates = ''; Out = '' }
  $cur = $script:CommonDir
  $starts = @()
  for ($i = 0; $i -lt 4 -and $cur; $i++) {
    $starts += $cur
    $cur = Split-Path -Parent $cur
  }
  foreach ($d in $starts) {
    $direct = Join-Path $d 'Шаблоны'
    if (Test-Path -LiteralPath $direct) { $res.Templates = $direct; break }
  }
  if (-not $res.Templates) {
    foreach ($d in $starts) {
      $hit = Get-ChildItem -LiteralPath $d -Directory -Recurse -Depth 2 -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ieq 'Шаблоны' } | Select-Object -First 1
      if ($hit) { $res.Templates = $hit.FullName; break }
    }
  }
  if ($res.Templates) {
    $res.Out = Join-Path (Split-Path -Parent $res.Templates) 'Готовые'
  }
  return $res
}

function Get-Engineer($settings, [string]$name) {
  foreach ($e in @($settings.Engineers)) { if ($e.Name -eq $name) { return $e } }
  return $null
}
