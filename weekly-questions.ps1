# Road to Vet - autonomous weekly question generator
# Runs unattended. Reads Paige's synced progress, drafts new questions for her weak
# subjects via the Claude API, INDEPENDENTLY re-checks the answer on each one, and
# publishes only the verified ones straight to the live Questions tab (active=TRUE).
# No human approval step. A question that fails self-check is dropped, never shown.
#
# Config: coach-config.json (git-ignored) in this folder:
#   { "endpoint": "https://.../exec", "anthropicKey": "sk-ant-...",
#     "model": "claude-sonnet-5", "perSubject": 4 }
#
# Schedule (weekly):
#   schtasks /Create /TN "Paige Road to Vet Coach" /TR "powershell -ExecutionPolicy Bypass -File C:\Users\BrendonGraham\paige-vet-pathway\weekly-questions.ps1" /SC WEEKLY /D SUN /ST 07:00

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---- config ----
$cfgPath = Join-Path $here "coach-config.json"
if (-not (Test-Path $cfgPath)) { Write-Host "Missing coach-config.json - see coach-config.example.json"; exit 1 }
$cfg = [IO.File]::ReadAllText($cfgPath) | ConvertFrom-Json
$endpoint = $cfg.endpoint
$apiKey   = $cfg.anthropicKey
$model    = if ($cfg.model) { $cfg.model } else { "claude-sonnet-5" }
$perSub   = if ($cfg.perSubject) { [int]$cfg.perSubject } else { 4 }
if (-not $endpoint -or -not $apiKey) { Write-Host "endpoint and anthropicKey required in coach-config.json"; exit 1 }

$subjNames = @{ bio="Biology"; chem="Chemistry"; phys="Physics"; vet="Vet and animal science"; math="Veterinary dosing / numeracy maths" }

function Invoke-Anthropic {
  param([string]$Prompt, [int]$MaxTokens = 2500)
  $payload = @{ model=$model; max_tokens=$MaxTokens; messages=@(@{ role="user"; content=$Prompt }) } | ConvertTo-Json -Depth 8 -Compress
  $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
  $headers = @{ "x-api-key"=$apiKey; "anthropic-version"="2023-06-01" }
  $resp = Invoke-WebRequest -Uri "https://api.anthropic.com/v1/messages" -Method Post -Headers $headers -ContentType "application/json" -Body $bytes -UseBasicParsing
  $txt = [Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
  $obj = $txt | ConvertFrom-Json
  return [string]$obj.content[0].text
}

function Get-JsonArray {
  param([string]$Text)
  $t = $Text -replace '(?s)^.*?```(?:json)?', '' -replace '(?s)```.*$', ''
  $s = $t.IndexOf('['); $e = $t.LastIndexOf(']')
  if ($s -lt 0 -or $e -lt $s) { return @() }
  try { return @(($t.Substring($s, $e - $s + 1)) | ConvertFrom-Json) } catch { return @() }
}

function Post-Publish {
  param($Rows)
  $body = @{ type="publish"; rows=$Rows } | ConvertTo-Json -Depth 8
  $bytes = [Text.Encoding]::UTF8.GetBytes($body)
  return Invoke-RestMethod -Uri $endpoint -Method Post -ContentType "text/plain; charset=utf-8" -Body $bytes
}

# ---- pull her progress + existing questions ----
$state = Invoke-RestMethod -Uri ("{0}?action=state&id=paige&t={1}" -f $endpoint, [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()) -Method Get
if (-not $state.data -or -not $state.data.quiz) { Write-Host "No synced progress yet - nothing to do."; exit 0 }
$ability = $state.data.quiz.ability

$existing = @(Invoke-RestMethod -Uri ("{0}?action=questions&t={1}" -f $endpoint, [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()) -Method Get)

$subs = @("bio","chem","phys","vet","math")
$ranked = $subs | Sort-Object { [double]$ability.$_ }
$weak = @($ranked | Select-Object -First 3)

$stamp = (Get-Date).ToString("yyyyMMdd")
$published = 0
$logLines = @()

foreach ($s in $weak) {
  $ab = [double]$ability.$s
  $lvl = [Math]::Max(1, [Math]::Min(3, [int][Math]::Round($ab)))
  $avoid = (@($existing | Where-Object { $_.subj -eq $s } | ForEach-Object { $_.q }) | Select-Object -First 40) -join " | "

  $genPrompt = @"
You are writing multiple-choice quiz questions for a 15-year-old New Zealand student (Paige) working towards Massey veterinary school. Subject: $($subjNames[$s]). Difficulty: NCEA Level $lvl.
Write exactly $perSub NEW questions. Rules:
- Factually correct, one unambiguous correct answer, 4 plausible options.
- NZ English. Clear, one idea each. For vet/maths, use farm and animal examples.
- Do NOT duplicate or closely rephrase any of these existing questions: $avoid
- A short teaching explanation (one or two sentences) for each.
Return ONLY a JSON array, no prose, each element:
{"q":"...","o":["A","B","C","D"],"answer":1,"ex":"..."}
where "answer" is the 1-based index (1-4) of the correct option.
"@

  $genText = Invoke-Anthropic -Prompt $genPrompt
  $gen = Get-JsonArray -Text $genText
  if (-not $gen.Count) { $logLines += "  $($subjNames[$s]) L$lvl - generation returned nothing, skipped"; continue }

  # ---- independent verification (re-derive the answer, blind to the claimed one) ----
  $vItems = @()
  for ($i = 0; $i -lt $gen.Count; $i++) {
    $vItems += @{ n=$i; q=$gen[$i].q; o=$gen[$i].o }
  }
  $vJson = ($vItems | ConvertTo-Json -Depth 6 -Compress)
  $verPrompt = @"
For each question below, independently work out which option (1-based index) is correct. Ignore any answer you might guess the author intended - solve it yourself.
Return ONLY a JSON array: [{"n":0,"correct":2}, ...] with the correct 1-based option index for each n.
Questions: $vJson
"@
  $verText = Invoke-Anthropic -Prompt $verPrompt
  $ver = Get-JsonArray -Text $verText
  $verMap = @{}
  foreach ($v in $ver) { $verMap[[int]$v.n] = [int]$v.correct }

  $kept = @()
  for ($i = 0; $i -lt $gen.Count; $i++) {
    $q = $gen[$i]
    $claimed = [int]$q.answer
    if (-not $verMap.ContainsKey($i)) { continue }
    if ($verMap[$i] -ne $claimed) { continue }                 # self-check failed -> drop
    if (-not $q.o -or $q.o.Count -lt 2) { continue }
    if ($claimed -lt 1 -or $claimed -gt $q.o.Count) { continue }
    $kept += @{
      id     = "gen_${s}_${stamp}_$i"
      subj   = $s
      lvl    = $lvl
      q      = [string]$q.q
      o      = @($q.o | ForEach-Object { [string]$_ })
      answer = $claimed
      ex     = [string]$q.ex
    }
  }

  if ($kept.Count) {
    $res = Post-Publish -Rows $kept
    $published += [int]$res.added
    $logLines += "  $($subjNames[$s]) L$lvl - generated $($gen.Count), verified $($kept.Count), published $($res.added)"
  } else {
    $logLines += "  $($subjNames[$s]) L$lvl - generated $($gen.Count), none passed self-check"
  }
}

$header = "- {0} - published {1} new questions (ability bio {2} chem {3} phys {4} vet {5} math {6})" -f `
  (Get-Date).ToString("yyyy-MM-dd HH:mm"), $published, `
  [Math]::Round([double]$ability.bio,2), [Math]::Round([double]$ability.chem,2), `
  [Math]::Round([double]$ability.phys,2), [Math]::Round([double]$ability.vet,2), [Math]::Round([double]$ability.math,2)
$logPath = Join-Path $here "coach-log.md"
Add-Content -Path $logPath -Value $header -Encoding utf8
foreach ($l in $logLines) { Add-Content -Path $logPath -Value $l -Encoding utf8 }

Write-Host $header
$logLines | ForEach-Object { Write-Host $_ }
