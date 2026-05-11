param(
  [Parameter(Mandatory=$true, ValueFromRemainingArguments=$true)]
  [string[]]$Paths
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Get-DocxPlainText([string]$DocxPath) {
  # Prefer Word COM automation when available (best text fidelity).
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($DocxPath, $false, $true)
    $text = $doc.Content.Text
    $doc.Close($false) | Out-Null
    $word.Quit() | Out-Null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    return $text
  } catch {
    # Fallback: unzip docx and extract runs from document.xml
    $tmp = Join-Path $env:TEMP ("docx_" + [guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    try {
      $zipPath = Join-Path $tmp "source.zip"
      Copy-Item -LiteralPath $DocxPath -Destination $zipPath -Force
      Expand-Archive -LiteralPath $zipPath -DestinationPath $tmp -Force
      $xmlPath = Join-Path $tmp "word\document.xml"
      $xml = Get-Content -LiteralPath $xmlPath -Raw -Encoding UTF8
      $sb = New-Object System.Text.StringBuilder
      $parts = $xml -split '</w:p>'
      foreach ($p in $parts) {
        $runs = [regex]::Matches($p, '<w:t[^>]*>([^<]*)</w:t>')
        if ($runs.Count -gt 0) {
          foreach ($m in $runs) { [void]$sb.Append($m.Groups[1].Value) }
          [void]$sb.AppendLine()
        }
      }
      return [System.Net.WebUtility]::HtmlDecode($sb.ToString())
    } finally {
      Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
  }
}

foreach ($p in $Paths) {
  if (-not (Test-Path -LiteralPath $p)) {
    Write-Output ("=== MISSING: {0} ===" -f $p)
    continue
  }
  Write-Output ("=== {0} ===" -f (Split-Path -Leaf $p))
  $out = Get-DocxPlainText -DocxPath $p
  # Trim excessive blank lines
  $out = ($out -split "`n") | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ -ne "" } | Out-String
  Write-Output $out
  Write-Output ""
}
