# Make a Hebrew document right-to-left, or check that it is. The same rules as hebrew_rtl.py, for a computer
# without Python: Windows PowerShell 5.1 is on every Windows machine, and the Python version stays the reference.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File hebrew_rtl.ps1 fix   <file> [<file> ...]
#   powershell -NoProfile -ExecutionPolicy Bypass -File hebrew_rtl.ps1 check <file> [<file> ...]
#
# Word: every paragraph w:bidi, Hebrew runs w:rtl with complex-script font/size/bold, tables w:bidiVisual,
# sections and styles RTL, and a "right" alignment on a Hebrew paragraph removed (in an RTL paragraph it puts
# the text on the left). PowerPoint: Hebrew paragraphs rtl="1" and right-aligned, runs he-IL, tables rtl="1".
# Excel: every sheet right-to-left. HTML: <html lang="he" dir="rtl">. Only XML parts change; the rest is copied.
# Exit code: 0 all good, 1 something is still left-to-right (check) or a file is missing, 2 wrong usage.

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$HebrewRe = [regex]'[֐-׿יִ-ﭏ]'

# Child order inside Word elements (ECMA-376); Word rejects a file whose children are out of order.
$P_PR = @('pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers',
  'pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE',
  'autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents',
  'suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr',
  'sectPr','pPrChange')
$R_PR = @('rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss',
  'imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs',
  'highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath')
$TBL_PR = @('tblStyle','tblpPr','tblOverlap','bidiVisual','tblStyleRowBandSize','tblStyleColBandSize','tblW','jc',
  'tblCellSpacing','tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription','tblPrChange')
$SECT_PR = @('headerReference','footerReference','footnotePr','endnotePr','type','pgSz','pgMar','paperSrc','pgBorders',
  'lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg','textDirection','bidi','rtlGutter','docGrid',
  'printerSettings','sectPrChange')
$A_RPR = @('ln','noFill','solidFill','gradFill','blipFill','pattFill','grpFill','effectLst','effectDag','highlight',
  'uLnTx','uLn','uFillTx','uFill','latin','ea','cs','sym','hlinkClick','hlinkMouseOver','rtl','extLst')
$WS = @('sheetPr','dimension','sheetViews')

$Labels = [ordered]@{
  paragraphs = 'פסקאות שלא מימין לשמאל'; runs = 'קטעי טקסט בעברית שלא מסומנים RTL'
  alignment = 'יישור שמציב עברית בצד שמאל'; tables = 'טבלאות משמאל לימין'
  sections = 'עמודים (sections) משמאל לימין'; styles = 'סגנונות משמאל לימין'
  fonts = 'גופן לעברית חסר'; sizes = 'גודל או הדגשה שלא חלים על עברית'; sheets = 'גיליונות משמאל לימין'
  html = 'דף בלי dir="rtl"'
}

# ---------- XML helpers ----------

function Kids($el, $ns, $name) {
  $out = @()
  if ($null -eq $el) { return ,$out }
  foreach ($c in $el.ChildNodes) {
    if ($c.NodeType -eq [Xml.XmlNodeType]::Element -and $c.NamespaceURI -eq $ns -and (-not $name -or $c.LocalName -eq $name)) { $out += $c }
  }
  return ,$out
}
function Kid($el, $ns, $name) { $k = Kids $el $ns $name; if ($k.Count) { return $k[0] } return $null }
function Desc($el, $ns, $name) { return @($el.GetElementsByTagName($name, $ns)) }
function TextOf($el, $ns) {
  $sb = New-Object Text.StringBuilder
  foreach ($t in $el.GetElementsByTagName('t', $ns)) { [void]$sb.Append($t.InnerText) }
  return $sb.ToString()
}
function NewChild($parent, $name, $ns) { return $parent.OwnerDocument.CreateElement($parent.Prefix, $name, $ns) }
function Ensure($parent, $name, $order, $ns) {
  if (-not $ns) { $ns = $parent.NamespaceURI }
  $existing = Kid $parent $ns $name
  if ($null -ne $existing) { return $existing }
  $new = NewChild $parent $name $ns
  $pos = [Array]::IndexOf($order, $name); if ($pos -lt 0) { $pos = $order.Count }
  $before = $null
  foreach ($c in (Kids $parent $ns $null)) {
    $i = [Array]::IndexOf($order, $c.LocalName)
    if ($i -ge 0 -and $i -gt $pos) { $before = $c; break }
  }
  if ($null -ne $before) { [void]$parent.InsertBefore($new, $before) } else { [void]$parent.AppendChild($new) }
  return $new
}
function EnsureFirst($parent, $name) {
  $existing = Kid $parent $parent.NamespaceURI $name
  if ($null -ne $existing) { return $existing }
  $new = NewChild $parent $name $parent.NamespaceURI
  $first = $null
  foreach ($c in $parent.ChildNodes) { if ($c.NodeType -eq [Xml.XmlNodeType]::Element) { $first = $c; break } }
  if ($null -ne $first) { [void]$parent.InsertBefore($new, $first) } else { [void]$parent.AppendChild($new) }
  return $new
}
function WAttr($el, $name) { if ($el.HasAttribute($name, $W)) { return $el.GetAttribute($name, $W) } return $null }
function SetWAttr($el, $name, $value) {
  $a = $el.OwnerDocument.CreateAttribute($el.Prefix, $name, $W); $a.Value = $value; [void]$el.Attributes.SetNamedItem($a)
}
function IsOff($el) { return ($null -ne $el) -and (@('0','false','off') -contains (WAttr $el 'val')) }
function ClearVal($el) { if ($el.HasAttribute('val', $W)) { $el.RemoveAttribute('val', $W) } }

# ---------- Word ----------

function WComplexScript($rPr, $found, $fix) {
  $fonts = Kid $rPr $W 'rFonts'
  if ($null -ne $fonts -and $null -eq (WAttr $fonts 'cs') -and $null -eq (WAttr $fonts 'cstheme')) {
    $latin = WAttr $fonts 'ascii'; if (-not $latin) { $latin = WAttr $fonts 'hAnsi' }
    if ($latin) { $found.fonts++; if ($fix) { SetWAttr $fonts 'cs' $latin } }
  }
  foreach ($pair in @(@('sz','szCs'), @('b','bCs'), @('i','iCs'))) {
    $src = Kid $rPr $W $pair[0]
    if ($null -ne $src -and $null -eq (Kid $rPr $W $pair[1])) {
      $found.sizes++
      if ($fix) { $el = Ensure $rPr $pair[1] $R_PR $W; $v = WAttr $src 'val'; if ($null -ne $v) { SetWAttr $el 'val' $v } }
    }
  }
}

function WParagraph($p, $found, $fix) {
  $pPr = Kid $p $W 'pPr'
  $isHebrew = $HebrewRe.IsMatch((TextOf $p $W))
  $bidi = $null; if ($null -ne $pPr) { $bidi = Kid $pPr $W 'bidi' }
  if ($null -eq $bidi -or (IsOff $bidi)) {
    $found.paragraphs++
    if ($fix) { if ($null -eq $pPr) { $pPr = EnsureFirst $p 'pPr' }; $bidi = Ensure $pPr 'bidi' $P_PR $W; ClearVal $bidi }
  }
  $jc = $null; if ($null -ne $pPr) { $jc = Kid $pPr $W 'jc' }
  if ($isHebrew -and $null -ne $jc -and (@('right','end') -contains (WAttr $jc 'val'))) {
    $found.alignment++
    if ($fix) { [void]$pPr.RemoveChild($jc) }
  }
  # Kids hands back its list as one object; collect the runs one by one so nothing ends up nested.
  $runs = New-Object Collections.ArrayList
  foreach ($r in (Kids $p $W 'r')) { [void]$runs.Add($r) }
  foreach ($h in (Kids $p $W 'hyperlink')) { foreach ($r in (Kids $h $W 'r')) { [void]$runs.Add($r) } }
  foreach ($r in $runs) {
    if (-not $HebrewRe.IsMatch((TextOf $r $W))) { continue }
    $rPr = Kid $r $W 'rPr'
    $rtl = $null; if ($null -ne $rPr) { $rtl = Kid $rPr $W 'rtl' }
    if ($null -eq $rtl -or (IsOff $rtl)) {
      $found.runs++
      if ($fix) { if ($null -eq $rPr) { $rPr = EnsureFirst $r 'rPr' }; $rtl = Ensure $rPr 'rtl' $R_PR $W; ClearVal $rtl }
    }
    if ($null -ne $rPr) { WComplexScript $rPr $found $fix }
  }
}

function WPart($doc, $found, $fix) {
  $root = $doc.DocumentElement
  foreach ($p in (Desc $root $W 'p')) { WParagraph $p $found $fix }
  foreach ($tbl in (Desc $root $W 'tbl')) {
    $tblPr = Kid $tbl $W 'tblPr'
    $flag = $null; if ($null -ne $tblPr) { $flag = Kid $tblPr $W 'bidiVisual' }
    if ($null -eq $flag -or (IsOff $flag)) {
      $found.tables++
      if ($fix) { if ($null -eq $tblPr) { $tblPr = EnsureFirst $tbl 'tblPr' }; $flag = Ensure $tblPr 'bidiVisual' $TBL_PR $W; ClearVal $flag }
    }
  }
  foreach ($sect in (Desc $root $W 'sectPr')) {
    if ($null -eq (Kid $sect $W 'bidi')) { $found.sections++; if ($fix) { [void](Ensure $sect 'bidi' $SECT_PR $W) } }
  }
}

function WStyles($doc, $found, $fix) {
  $root = $doc.DocumentElement
  $defaults = Kid $root $W 'docDefaults'
  if ($null -ne $defaults) {
    $pDefault = Kid $defaults $W 'pPrDefault'
    $pPr = $null; if ($null -ne $pDefault) { $pPr = Kid $pDefault $W 'pPr' }
    if ($null -eq $pPr -or $null -eq (Kid $pPr $W 'bidi')) {
      $found.styles++
      if ($fix) {
        if ($null -eq $pDefault) { $pDefault = Ensure $defaults 'pPrDefault' @('rPrDefault','pPrDefault') $W }
        if ($null -eq $pPr) { $pPr = EnsureFirst $pDefault 'pPr' }
        [void](Ensure $pPr 'bidi' $P_PR $W)
      }
    }
    $rDefault = Kid $defaults $W 'rPrDefault'
    $rPr = $null; if ($null -ne $rDefault) { $rPr = Kid $rDefault $W 'rPr' }
    if ($null -ne $rPr) {
      if ($fix) { $lang = Ensure $rPr 'lang' $R_PR $W; if ((WAttr $lang 'bidi') -ne 'he-IL') { SetWAttr $lang 'bidi' 'he-IL' } }
      WComplexScript $rPr $found $fix
    }
  }
  foreach ($style in (Kids $root $W 'style')) {
    if (@('paragraph','character','table') -notcontains (WAttr $style 'type')) { continue }
    $pPr = Kid $style $W 'pPr'
    if ($null -ne $pPr) {
      $bidi = Kid $pPr $W 'bidi'; $jc = Kid $pPr $W 'jc'
      if (IsOff $bidi) { $found.styles++; if ($fix) { ClearVal $bidi } }
      if ($null -ne $jc -and (@('right','end') -contains (WAttr $jc 'val'))) { $found.alignment++; if ($fix) { [void]$pPr.RemoveChild($jc) } }
    }
    $rPr = Kid $style $W 'rPr'
    if ($null -ne $rPr) { WComplexScript $rPr $found $fix }
  }
}

function Word($doc, $name, $found, $fix) {
  if ($name -eq 'word/styles.xml') { WStyles $doc $found $fix }
  elseif ($name -match '^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$') { WPart $doc $found $fix }
}

# ---------- PowerPoint ----------

function PowerPoint($doc, $name, $found, $fix) {
  if ($name -notmatch '^ppt/(slides/slide|notesSlides/notesSlide)\d+\.xml$') { return }
  $root = $doc.DocumentElement
  foreach ($p in (Desc $root $A 'p')) {
    if (-not $HebrewRe.IsMatch((TextOf $p $A))) { continue }
    $pPr = Kid $p $A 'pPr'
    if ($null -eq $pPr -or @('1','true') -notcontains $pPr.GetAttribute('rtl') -or @('','l') -contains $pPr.GetAttribute('algn')) {
      $found.paragraphs++
      if ($fix) {
        if ($null -eq $pPr) { $pPr = EnsureFirst $p 'pPr' }
        $pPr.SetAttribute('rtl', '1')
        if (@('','l') -contains $pPr.GetAttribute('algn')) { $pPr.SetAttribute('algn', 'r') }
      }
    }
    foreach ($r in (Kids $p $A 'r')) {
      if (-not $HebrewRe.IsMatch((TextOf $r $A))) { continue }
      $rPr = Kid $r $A 'rPr'
      if ($null -eq $rPr -or $rPr.GetAttribute('lang') -ne 'he-IL' -or ($null -ne (Kid $rPr $A 'latin') -and $null -eq (Kid $rPr $A 'cs'))) {
        $found.runs++
        if ($fix) {
          if ($null -eq $rPr) { $rPr = EnsureFirst $r 'rPr' }
          $rPr.SetAttribute('lang', 'he-IL')
          $latin = Kid $rPr $A 'latin'
          if ($null -ne $latin -and $null -eq (Kid $rPr $A 'cs')) { $cs = Ensure $rPr 'cs' $A_RPR $A; $cs.SetAttribute('typeface', $latin.GetAttribute('typeface')) }
        }
      }
    }
  }
  foreach ($tbl in (Desc $root $A 'tbl')) {
    $tblPr = Kid $tbl $A 'tblPr'
    if ($null -eq $tblPr -or @('1','true') -notcontains $tblPr.GetAttribute('rtl')) {
      $found.tables++
      if ($fix) { if ($null -eq $tblPr) { $tblPr = EnsureFirst $tbl 'tblPr' }; $tblPr.SetAttribute('rtl', '1') }
    }
  }
}

# ---------- Excel ----------

function Excel($doc, $name, $found, $fix) {
  if ($name -notmatch '^xl/worksheets/sheet\d+\.xml$') { return }
  $root = $doc.DocumentElement
  $views = Kid $root $S 'sheetViews'
  if ($null -eq $views -or (Kids $views $S 'sheetView').Count -eq 0) {
    $found.sheets++
    if ($fix) {
      if ($null -eq $views) { $views = Ensure $root 'sheetViews' $WS $S }
      $view = NewChild $views 'sheetView' $S
      $view.SetAttribute('workbookViewId', '0'); $view.SetAttribute('rightToLeft', '1')
      [void]$views.AppendChild($view)
    }
    return
  }
  foreach ($view in (Kids $views $S 'sheetView')) {
    if (@('1','true') -notcontains $view.GetAttribute('rightToLeft')) { $found.sheets++; if ($fix) { $view.SetAttribute('rightToLeft', '1') } }
  }
}

# ---------- files ----------

$Handlers = @{ '.docx' = 'Word'; '.docm' = 'Word'; '.dotx' = 'Word'; '.pptx' = 'PowerPoint'; '.xlsx' = 'Excel'; '.xlsm' = 'Excel' }
$Utf8 = New-Object Text.UTF8Encoding($false)

function NewFound { $f = @{}; foreach ($k in $Labels.Keys) { $f[$k] = 0 }; return $f }
function Total($found) { $n = 0; foreach ($v in $found.Values) { $n += $v }; return $n }

function Office($path, $fix) {
  $handler = $Handlers[[IO.Path]::GetExtension($path).ToLowerInvariant()]
  $found = NewFound
  $entries = New-Object Collections.ArrayList
  $changed = $false
  $zip = [IO.Compression.ZipFile]::OpenRead($path)
  try {
    foreach ($e in $zip.Entries) {
      $ms = New-Object IO.MemoryStream
      $st = $e.Open(); $st.CopyTo($ms); $st.Close()
      $bytes = $ms.ToArray()
      $item = @{ Name = $e.FullName; Bytes = $bytes; Time = $e.LastWriteTime }
      if ($e.FullName.EndsWith('.xml') -and $bytes.Length -gt 0) {
        $text = $Utf8.GetString($bytes)
        if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
        if ($text.Contains('<')) {
          $doc = New-Object Xml.XmlDocument
          $doc.PreserveWhitespace = $true
          $doc.XmlResolver = $null
          $doc.LoadXml($text)
          $before = Total $found
          & $handler $doc $e.FullName $found $fix
          if ($fix -and (Total $found) -ne $before) {
            $xml = $doc.OuterXml
            if ($doc.FirstChild.NodeType -ne [Xml.XmlNodeType]::XmlDeclaration) {
              $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n" + $xml
            }
            $item.Bytes = $Utf8.GetBytes($xml)
            $changed = $true
          }
        }
      }
      [void]$entries.Add($item)
    }
  } finally { $zip.Dispose() }
  if ($fix -and $changed) {
    $tmp = [IO.Path]::Combine([IO.Path]::GetTempPath(), [Guid]::NewGuid().ToString() + [IO.Path]::GetExtension($path))
    $out = [IO.Compression.ZipFile]::Open($tmp, [IO.Compression.ZipArchiveMode]::Create)
    try {
      foreach ($item in $entries) {
        $entry = $out.CreateEntry($item.Name, [IO.Compression.CompressionLevel]::Optimal)
        $entry.LastWriteTime = $item.Time
        $ws = $entry.Open(); $ws.Write($item.Bytes, 0, $item.Bytes.Length); $ws.Close()
      }
    } finally { $out.Dispose() }
    Move-Item -LiteralPath $tmp -Destination $path -Force
  }
  return $found
}

function HtmlFile($path, $fix) {
  $found = NewFound
  $text = [IO.File]::ReadAllText($path, $Utf8)
  $tag = [regex]::Match($text, '<html\b[^>]*>', 'IgnoreCase')
  $ok = $tag.Success -and [regex]::IsMatch($tag.Value, '\bdir\s*=\s*["'']?rtl', 'IgnoreCase')
  if (-not $ok) {
    $found.html++
    if ($fix) {
      if ($tag.Success) {
        $inner = $tag.Value.Substring(0, $tag.Value.Length - 1)
        $new = [regex]::Replace($inner, '\s(dir|lang)\s*=\s*("[^"]*"|''[^'']*''|\S+)', '') + ' lang="he" dir="rtl">'
        $text = $text.Substring(0, $tag.Index) + $new + $text.Substring($tag.Index + $tag.Length)
      } else {
        $text = "<!doctype html>`n<html lang=""he"" dir=""rtl"">`n<meta charset=""utf-8"">`n" + $text + "`n</html>`n"
      }
      [IO.File]::WriteAllText($path, $text, $Utf8)
    }
  }
  return $found
}

if ($args.Count -lt 2 -or @('fix','check') -notcontains $args[0]) {
  Write-Output 'hebrew_rtl.ps1 fix <file> [<file> ...]     תיקון במקום (Word, PowerPoint, Excel, HTML)'
  Write-Output 'hebrew_rtl.ps1 check <file> [<file> ...]   מה עוד משמאל לימין; קוד יציאה 1 אם יש'
  exit 2
}
$mode = $args[0]
$bad = $false
foreach ($arg in $args[1..($args.Count - 1)]) {
  $name = [IO.Path]::GetFileName($arg)
  if (-not (Test-Path -LiteralPath $arg -PathType Leaf)) { Write-Output "✗ ${arg}: הקובץ לא נמצא"; $bad = $true; continue }
  $path = (Resolve-Path -LiteralPath $arg).ProviderPath
  $ext = [IO.Path]::GetExtension($path).ToLowerInvariant()
  try {
    if ($Handlers.ContainsKey($ext)) { $found = Office $path ($mode -eq 'fix') }
    elseif (@('.html','.htm') -contains $ext) { $found = HtmlFile $path ($mode -eq 'fix') }
    else { Write-Output "– ${name}: לא מסמך Word, PowerPoint, Excel או HTML. אין מה לבדוק"; continue }
  } catch {
    Write-Output "✗ ${name}: $($_.Exception.Message)"; $bad = $true; continue
  }
  $issues = @()
  foreach ($k in $Labels.Keys) { if ($found[$k]) { $issues += "$($Labels[$k]): $($found[$k])" } }
  if ($mode -eq 'fix') {
    if ($issues.Count) { Write-Output "✓ ${name}: תוקן מימין לשמאל ($($issues -join ', '))" } else { Write-Output "✓ ${name}: כבר מימין לשמאל" }
  } elseif ($issues.Count) {
    $bad = $true; Write-Output "✗ ${name}: $($issues -join ', ')"
  } else {
    Write-Output "✓ ${name}: מימין לשמאל"
  }
}
if ($bad) { exit 1 } else { exit 0 }
