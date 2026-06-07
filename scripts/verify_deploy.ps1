# verify_deploy.ps1 - Deploy verification (ASCII only to avoid PS5.1 encoding issues)
#
# Purpose: when a fix is pushed but the browser still shows the old page,
#          this checks the LIVE GitHub Pages server directly to tell apart
#          "deploy not reflected yet" vs "browser cache problem".
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File scripts/verify_deploy.ps1
#   (run 1-2 min after push; if FAIL due to CDN lag, wait and re-run)
#
# To add a new check: append a @{ url=; marker=; desc= } line to $checks.

$ErrorActionPreference = "Stop"
$base = "https://you0810jmsdf.github.io/cocola-site"

$checks = @(
  @{ url = "$base/member-app/html.html";     marker = "removeAttribute('readonly')"; desc = "member-app search readonly" }
  @{ url = "$base/dao/index.html";           marker = 'id="memberSearch" placeholder'; desc = "dao memberSearch exists" }
  @{ url = "$base/dao/index.html";           marker = "removeAttribute('readonly')"; desc = "dao search readonly" }
  @{ url = "$base/events/index.html";        marker = 'id="search-input"'; desc = "events search exists" }
  @{ url = "$base/organizations/index.html"; marker = 'id="search" type="search"'; desc = "organizations search readonly" }
  @{ url = "$base/dao/service-worker.js";    marker = "cocola-dao-v22"; desc = "dao SW cache v22" }
)

$fail = 0
foreach ($c in $checks) {
  $u = $c.url + "?_v=" + (Get-Date -UFormat %s)
  try {
    $html = Invoke-WebRequest -Uri $u -Headers @{ "Cache-Control" = "no-cache" } -UseBasicParsing | Select-Object -ExpandProperty Content
  } catch {
    Write-Host ("[ERROR] {0} : fetch failed {1}" -f $c.desc, $_.Exception.Message) -ForegroundColor Red
    $fail++; continue
  }
  if ($html -like ("*" + $c.marker + "*")) {
    Write-Host ("[OK]   {0}" -f $c.desc) -ForegroundColor Green
  } else {
    Write-Host ("[FAIL] {0} : marker not found (CDN lag or code not reflected)" -f $c.desc) -ForegroundColor Yellow
    $fail++
  }
}

Write-Host ""
if ($fail -eq 0) {
  Write-Host "ALL OK = server has the new version." -ForegroundColor Green
  Write-Host "If browser still shows old: it is BROWSER CACHE -> F12 -> right-click reload -> Empty Cache and Hard Reload." -ForegroundColor Cyan
  exit 0
} else {
  Write-Host ("{0} check(s) failed. If just pushed, wait 1-2 min (CDN lag) and re-run. If it persists, check the code." -f $fail) -ForegroundColor Yellow
  exit 1
}
