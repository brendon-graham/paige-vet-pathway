# Road to Vet - weekly coaching compute step
# Reads Paige's synced progress (READ-ONLY) and writes a coaching brief describing
# which subjects/levels need new questions. It does NOT invent questions and does NOT
# write anything to her live bank - the brain drafts + verifies questions from this brief,
# they land in the Pending tab, and only reviewed rows go live. (See SETUP.md.)
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File weekly-coach.ps1 [-Endpoint <url>]
# If -Endpoint is omitted it reads coach-config.json  ->  { "endpoint": "https://.../exec" }

param([string]$Endpoint = "")

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

if (-not $Endpoint) {
  $cfgPath = Join-Path $here "coach-config.json"
  if (Test-Path $cfgPath) {
    $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
    $Endpoint = $cfg.endpoint
  }
}
if (-not $Endpoint) {
  Write-Host "No endpoint. Pass -Endpoint or create coach-config.json with { endpoint: '.../exec' }."
  exit 1
}

# --- pull her state (read-only) ---
$state = Invoke-RestMethod -Uri ("{0}?action=state&id=paige&t={1}" -f $Endpoint, [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()) -Method Get
if (-not $state.data -or -not $state.data.quiz) {
  Write-Host "No progress yet for 'paige' (she hasn't synced). Nothing to do."
  exit 0
}

$ability = $state.data.quiz.ability
$missed  = @($state.data.quiz.missed)
$history = @($state.data.quiz.history)

$subjNames = @{ bio="Biology"; chem="Chemistry"; phys="Physics"; vet="Vet & Animal"; math="Vet Maths" }
$subs = @("bio","chem","phys","vet","math")

# rank subjects weakest-first by ability
$ranked = $subs | Sort-Object { [double]$ability.$_ }

# the 3 weakest subjects get fresh questions at their current working level
$weak = @()
foreach ($s in ($ranked | Select-Object -First 3)) {
  $ab  = [double]$ability.$s
  $lvl = [Math]::Max(1, [Math]::Min(3, [Math]::Round($ab)))
  $weak += [pscustomobject]@{
    subj     = $s
    name     = $subjNames[$s]
    ability  = [Math]::Round($ab, 2)
    level    = [int]$lvl
    addCount = 4               # draft this many new questions here
  }
}

$brief = [pscustomobject]@{
  generatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm")
  person      = "paige"
  ability     = $ability
  weakSubjects= $weak
  missedCount = $missed.Count
  quizzesDone = $history.Count
  note        = "Draft the addCount questions per weakSubject at the given level, verify each answer, POST to the Pending tab (type:pending). Do NOT write to live Questions."
}

$briefPath = Join-Path $here "coach-brief.json"
$brief | ConvertTo-Json -Depth 6 | Out-File -FilePath $briefPath -Encoding utf8

# --- log ---
$logPath = Join-Path $here "coach-log.md"
$line = "- {0} - quizzes {1}, review backlog {2}. Weak: {3}" -f `
  $brief.generatedAt, $history.Count, $missed.Count, (($weak | ForEach-Object { "{0} L{1} ({2})" -f $_.name, $_.level, $_.ability }) -join ", ")
Add-Content -Path $logPath -Value $line -Encoding utf8

Write-Host "Coaching brief written: $briefPath"
Write-Host $line
