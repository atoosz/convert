<#
.SYNOPSIS
  Conversion smoke test: generates small sample files and converts each one to every output format
  the app offers for it, using the app's --cli mode.

.DESCRIPTION
    .\smoke-test.ps1 [[-Exe] Convert.exe] [-Strict] [-RequireLibreOffice] [-Out DIR] [-Keep] [-TimeoutSec N]

    -Exe                 the Convert.exe to test (default: out\smoke-app\Convert.exe, which
                         .\check.ps1 build makes; .\check.ps1 smoke builds it for you)
    -Strict              missing ffmpeg/ffprobe is a failure (CI uses this). Without it, formats
                         that need ffmpeg are reported as SKIP.
    -RequireLibreOffice  missing LibreOffice is a failure. Without it, conversions that need it
                         (to or from Word, ODT, RTF) are reported as SKIP. CI doesn't install it.
    -Out DIR             keep samples, outputs and per-conversion logs in DIR (default: a temp
                         folder, deleted afterwards unless something failed or -Keep is given)
    -TimeoutSec N        time per conversion (default: $env:SMOKE_TIMEOUT, else 120); the app and
                         everything it started are killed when it's up

  A conversion passes when the app exits 0, prints no FAIL line, and writes at least one non-empty
  file.

  One conversion of every kind also runs from a read-only folder, where the output has to go to the
  fallback folder instead.

  Font samples are copies of fonts that come with Windows; nothing is committed. Converted fonts are
  also checked for the right first bytes.

  Formats with a hardware encoder (H.264, HEVC, MKV; ProRes has none on Windows) also run with
  hardware encoding off. On a machine without a GPU encoder (e.g. GitHub's Windows runners) the
  hardware run still has to pass: the app falls back to the software encoder by itself.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)][string]$Exe,
    [switch]$Strict,
    [switch]$RequireLibreOffice,
    [string]$Out,
    [switch]$Keep,
    [int]$TimeoutSec = 0
)
$ErrorActionPreference = "Stop"
# Keep this file ASCII: Windows PowerShell reads a .ps1 without a BOM in the ANSI code page.

if (-not $Exe) { $Exe = Join-Path $PSScriptRoot "out\smoke-app\Convert.exe" }
if (-not (Test-Path -LiteralPath $Exe -PathType Leaf)) {
    Write-Host "No Convert.exe at $Exe. Build it with .\check.ps1 build (or run .\check.ps1 smoke, which builds it)."
    exit 2
}
$Exe = (Resolve-Path -LiteralPath $Exe).ProviderPath
if ($TimeoutSec -le 0) {
    $TimeoutSec = 120
    if ($env:SMOKE_TIMEOUT) { $TimeoutSec = [int]$env:SMOKE_TIMEOUT }
}
$HwTargets = @("mp4-h264", "mp4-hevc", "mkv")
$Utf8 = New-Object System.Text.UTF8Encoding $false

# An existing -Out folder is emptied first, but only if it's empty or an earlier run's: a typo
# must not wipe someone's folder.
$Marker = ".convert-smoke"
if (-not $Out) {
    $Out = Join-Path ([IO.Path]::GetTempPath()) ("convert-smoke-" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
} elseif (Test-Path -LiteralPath $Out) {
    $existing = @(Get-ChildItem -LiteralPath $Out -Force)
    if ($existing.Count -gt 0 -and -not (Test-Path -LiteralPath (Join-Path $Out $Marker))) {
        Write-Host "$Out isn't empty and wasn't made by this script. Pick another folder for -Out."
        exit 2
    }
    Remove-Item -LiteralPath $Out -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $Out | Out-Null
$Out = (Resolve-Path -LiteralPath $Out).ProviderPath
[IO.File]::WriteAllText((Join-Path $Out $Marker), "", $Utf8)
$Samples = Join-Path $Out "samples"
$Runs = Join-Path $Out "runs"
New-Item -ItemType Directory -Force -Path $Samples, $Runs | Out-Null

$script:Passed = 0
$script:Skipped = 0
$script:Failures = New-Object System.Collections.Generic.List[string]
$script:Covered = New-Object System.Collections.Generic.HashSet[string]

function Show($status, $label, $name, $detail) {
    # The console may not be UTF-8; the logs keep the app's own text.
    $detail = "$detail".Replace([string][char]0x2192, "->")
    Write-Host ("{0,-4}  {1,-20} {2,-12} {3}" -f $status, $label, $name, $detail)
}
function Pass($label, $name, $detail) { $script:Passed++; Show "PASS" $label $name $detail }
function Skip($label, $name, $detail) { $script:Skipped++; Show "SKIP" $label $name $detail }
function Fail($label, $name, $detail) {
    $script:Failures.Add(("{0,-20} {1,-12} {2}" -f $label, $name, "$detail".Replace([string][char]0x2192, "->")))
    Show "FAIL" $label $name $detail
}

# --- Running programs ------------------------------------------------------------------------

function Format-Arg([string]$a) {
    if ($a -ne "" -and $a -notmatch '[\s"]') { return $a }
    $a = $a -replace '(\\*)"', '$1$1\"'
    $a = $a -replace '(\\+)$', '$1$1'
    return '"' + $a + '"'
}

# Runs a program with its output in $Log and returns its exit code, or $null if it ran out of time
# (it and everything it started are killed then). Start-Process rather than &: Convert.exe is a GUI
# program, which PowerShell doesn't wait for, and this is the one way that redirects to a file, keeps
# the exit code and allows a timeout in both Windows PowerShell and PowerShell 7. Output goes to a
# file, not a pipe, because what the app starts (WebView2, LibreOffice) can outlive it and would
# keep a pipe open. Whatever went to stderr (a crash, ffmpeg's errors) is kept in "$Log.stderr".
function Invoke-Logged {
    param([string]$File, [string[]]$Arguments, [string]$Log, [hashtable]$Environment = @{}, [string]$WorkingDirectory = $Out)
    $stderr = "$Log.stderr"
    if ($File -eq $Exe -and -not $Environment.ContainsKey("WEBVIEW2_USER_DATA_FOLDER")) {
        # A WebView2 folder of the run's own: the app's usual one is shared with every other Convert
        # (and unit test run) on the PC, and one of those using it at the same time makes PDF creation
        # fail (0x8007139F) or hang.
        $Environment = $Environment.Clone()
        $Environment["WEBVIEW2_USER_DATA_FOLDER"] = Join-Path $Out "webview2"
    }
    $saved = @{}
    foreach ($k in $Environment.Keys) {
        $saved[$k] = [Environment]::GetEnvironmentVariable($k)
        [Environment]::SetEnvironmentVariable($k, $Environment[$k])
    }
    try {
        $argLine = ($Arguments | ForEach-Object { Format-Arg $_ }) -join " "
        $p = Start-Process -FilePath $File -ArgumentList $argLine -WorkingDirectory $WorkingDirectory -NoNewWindow -PassThru `
            -RedirectStandardOutput $Log -RedirectStandardError $stderr
    } finally {
        # The child has its own copy by now.
        foreach ($k in $saved.Keys) { [Environment]::SetEnvironmentVariable($k, $saved[$k]) }
    }
    $null = $p.Handle   # without a handle taken while it runs, ExitCode is empty afterwards
    $code = $null
    if ($p.WaitForExit($TimeoutSec * 1000)) {
        $p.WaitForExit()
        $code = $p.ExitCode
    } else {
        $null = Invoke-Quiet taskkill.exe /PID $p.Id /T /F
        $p.WaitForExit()
    }
    if ((Test-Path -LiteralPath $stderr) -and (Get-Item -LiteralPath $stderr).Length -eq 0) {
        Remove-Item -LiteralPath $stderr -ErrorAction SilentlyContinue
    }
    return $code
}

# For the small console tools (icacls, taskkill): with output redirected, Windows PowerShell turns
# their stderr into errors, which "Stop" would make fatal. Their exit code says what happened.
function Invoke-Quiet([string]$File) {
    $ErrorActionPreference = "Continue"
    & $File @args 2>&1 | Out-Null
    return $LASTEXITCODE
}

# What a run printed: its output, then its stderr. Shared reading, because something the app started
# (soffice.bin) can still hold the log file open after the app is gone.
function Read-Log([string]$Log) {
    $lines = @()
    foreach ($path in $Log, "$Log.stderr") {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $stream = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]"ReadWrite, Delete")
        $reader = New-Object IO.StreamReader($stream, $Utf8)
        try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $lines += @($text -split "\r?\n" | Where-Object { $_ -ne "" })
    }
    return $lines
}

function Get-Outputs([string]$Dir, [string]$Except) {
    @(Get-ChildItem -LiteralPath $Dir -Recurse -File | Where-Object {
        $_.Length -gt 0 -and $_.Name -ne "output.log" -and $_.Name -ne "output.log.stderr" -and $_.FullName -ne $Except
    })
}

# Why a run failed: the app's FAIL line, else the last thing it printed.
function Get-Reason([string[]]$Lines) {
    $fail = $Lines | Where-Object { $_ -like "FAIL*" } | Select-Object -First 1
    if ($fail) { return ($fail -replace '^FAIL [^:]*:\s*', '') }
    $last = $Lines | Where-Object { $_.Trim() } | Select-Object -Last 1
    if ($last) { return $last }
    return "no output"
}

# Invoke-Conversion <label> <target> <sample file> [settings JSON]
function Invoke-Conversion([string]$Label, [string]$Target, [string]$Sample, [string]$Settings = "") {
    $name = Split-Path $Sample -Leaf
    $dir = Join-Path $Runs "$name-$Label"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Copy-Item -LiteralPath $Sample -Destination $dir
    $file = Join-Path $dir $name
    $log = Join-Path $dir "output.log"
    $envs = @{}
    if ($Settings) { $envs["CONVERT_SETTINGS"] = $Settings }
    $code = Invoke-Logged $Exe @("--cli", $Target, $file) $log $envs
    $lines = Read-Log $log
    $outputs = Get-Outputs $dir $file
    if ($null -eq $code) {
        Fail $Label $name "timed out after ${TimeoutSec}s"
    } elseif ($code -ne 0 -or ($lines | Where-Object { $_ -like "FAIL*" })) {
        Fail $Label $name "$(Get-Reason $lines) (exit $code)"
    } elseif ($outputs.Count -eq 0) {
        Fail $Label $name "no output file written"
    } else {
        $ok = $lines | Where-Object { $_ -like "OK*" } | Select-Object -First 1
        Pass $Label $name ("$ok" -replace '^OK\s*', '')
    }
}

# --- Tools -----------------------------------------------------------------------------------

# The app looks for its tools in these folders (Tools.Directories() in Tools.cs), so the script
# finds the same ffmpeg the app will use.
function Get-ToolDirectories {
    $dirs = New-Object System.Collections.Generic.List[string]
    $dirs.Add((Join-Path (Split-Path $Exe) "tools"))
    $paths = @($env:PATH)
    foreach ($key in "HKCU:\Environment", "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment") {
        $value = Get-ItemProperty -Path $key -Name Path -ErrorAction SilentlyContinue
        if ($value) { $paths += $value.Path }
    }
    foreach ($path in $paths) {
        foreach ($d in "$path".Split(";")) {
            $d = $d.Trim().Trim('"')
            # The app skips relative entries too.
            if ($d -match '^([A-Za-z]:\\|\\\\)') { $dirs.Add($d) }
        }
    }
    $dirs.Add((Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links"))
    return $dirs
}

function Find-Tool([string]$Name, [string[]]$Extra = @()) {
    foreach ($d in @(Get-ToolDirectories) + $Extra) {
        $p = Join-Path $d $Name
        if (Test-Path -LiteralPath $p -PathType Leaf) { return $p }
    }
    return $null
}

$FFmpeg = Find-Tool "ffmpeg.exe"
$FFprobe = $null
if ($FFmpeg -and (Test-Path -LiteralPath (Join-Path (Split-Path $FFmpeg) "ffprobe.exe"))) { $FFprobe = Join-Path (Split-Path $FFmpeg) "ffprobe.exe" }
if (-not $FFprobe) { $FFprobe = Find-Tool "ffprobe.exe" }
$LibreOffice = Find-Tool "soffice.exe" @(
    (Join-Path $env:ProgramFiles "LibreOffice\program"),
    (Join-Path ${env:ProgramFiles(x86)} "LibreOffice\program"))

$WebView2 = $null
foreach ($key in "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
                 "HKCU:\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}") {
    $value = Get-ItemProperty -Path $key -Name pv -ErrorAction SilentlyContinue
    if ($value -and $value.pv -and $value.pv -ne "0.0.0.0") { $WebView2 = $value.pv; break }
}

function Or-Missing($value) { if ($value) { return $value } return "not installed" }

Write-Host "Convert smoke test"
Write-Host "  app:          $Exe"
Write-Host "  ffmpeg:       $(Or-Missing $FFmpeg)"
Write-Host "  LibreOffice:  $(Or-Missing $LibreOffice)"
Write-Host "  WebView2:     $(Or-Missing $WebView2)"
Write-Host "  output:       $Out"
Write-Host ""

$listLog = Join-Path $Out "targets.log"
$null = Invoke-Logged $Exe @("--cli") $listLog
$AllTargets = @()
foreach ($line in Read-Log $listLog) {
    if ($line -like "targets: *") { $AllTargets = @($line.Substring(9).Split(" ", [StringSplitOptions]::RemoveEmptyEntries)) }
}
if ($AllTargets.Count -eq 0) {
    Write-Host "Couldn't read the target list from '$Exe --cli' (see $listLog)."
    exit 1
}

if (-not $FFmpeg -or -not $FFprobe) {
    if ($Strict) {
        Fail "ffmpeg" "-" "ffmpeg/ffprobe not installed (winget install Gyan.FFmpeg)"
    } else {
        Write-Host "Note: ffmpeg isn't installed, so audio and video formats are skipped (winget install Gyan.FFmpeg)."
    }
}
if (-not $LibreOffice) {
    if ($RequireLibreOffice) {
        Fail "libreoffice" "-" "LibreOffice not installed (winget install TheDocumentFoundation.LibreOffice)"
    } else {
        Write-Host "Note: LibreOffice isn't installed, so conversions to and from Word, ODT and RTF are skipped."
    }
}
# Offered only when this PC can write them (Targets.IsAvailable), so they're not tested otherwise.
if ($AllTargets -notcontains "heic") { Write-Host "Note: HEIC isn't offered here (needs the HEIF and HEVC extensions from the Store)." }
if ($AllTargets -notcontains "he-aac") { Write-Host "Note: HE-AAC isn't offered here (needs an ffmpeg built with libfdk_aac)." }

# --- Samples (generated, nothing committed) --------------------------------------------------

function Write-Sample([string]$Name, [string]$Text) {
    [IO.File]::WriteAllText((Join-Path $Samples $Name), $Text.Replace("`r`n", "`n"), $Utf8)
}

Write-Sample "sample.md" @'
# Smoke test

Some **bold**, *italic* and `code`, plus a [link](https://example.com).

1. First
2. Second
   - nested item

- [x] done
- [ ] open

| Format | Engine   |
|--------|----------|
| PDF    | WebView2 |

> A quote.

```csharp
Console.WriteLine("hello");
```
'@

# A TypeScript file: .ts is also the extension of MPEG transport streams (stream.ts below), so the
# app has to look inside to tell them apart.
Write-Sample "script.ts" @'
export function greet(name: string): string {
  return `Hello, ${name}!`;
}

console.log(greet("smoke test"));
'@

$eAcute = [char]0x00E9; $iUml = [char]0x00EF; $uUml = [char]0x00FC
Write-Sample "sample.txt" @"
Smoke test

Plain text with some accents: caf$eAcute, na${iUml}ve, Z${uUml}rich.
A second paragraph.
"@

# Why a sample is missing, for the failures it causes.
$script:SampleErrors = @{}

# Runs the app on a copy of a sample in a folder of its own and puts what it made into the samples.
function New-SampleWithApp([string]$From, [string]$Target, [string]$Made, [string]$Folder) {
    $dir = Join-Path $Out $Folder
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Copy-Item -LiteralPath (Join-Path $Samples $From) -Destination $dir -Force
    $log = Join-Path $dir "output-$Target.log"
    $code = Invoke-Logged $Exe @("--cli", $Target, (Join-Path $dir $From)) $log
    $result = Join-Path $dir $Made
    if ((Test-Path -LiteralPath $result -PathType Leaf) -and (Get-Item -LiteralPath $result).Length -gt 0) {
        Copy-Item -LiteralPath $result -Destination $Samples
    } elseif ($null -eq $code) {
        $script:SampleErrors[$Made] = "$From to $Target timed out after ${TimeoutSec}s ($log)"
    } else {
        $script:SampleErrors[$Made] = "$From to ${Target}: $(Get-Reason (Read-Log $log)) ($log)"
    }
}

# The PDF comes from the app itself (Markdown -> PDF), so it has real, extractable text. If that
# fails, every PDF format is reported as failed below.
New-SampleWithApp "sample.md" "pdf-doc" "sample.pdf" "pdf-sample"

# A Word file from the app too (through LibreOffice), so Office documents are also read.
if ($LibreOffice) { New-SampleWithApp "sample.md" "docx" "sample.docx" "docx-sample" }

function Invoke-FFmpeg([string]$Name, [string[]]$Arguments) {
    $log = Join-Path $Out "ffmpeg-$Name.log"
    $code = Invoke-Logged $FFmpeg (@("-v", "error", "-y") + $Arguments + @((Join-Path $Samples $Name))) $log
    if ($code -ne 0) {
        Remove-Item -LiteralPath (Join-Path $Samples $Name) -ErrorAction SilentlyContinue
        $script:SampleErrors[$Name] = "ffmpeg: $(Get-Reason (Read-Log $log)) ($log)"
    }
}
if ($FFmpeg) {
    # Hi-res stereo WAV: 24-bit, 96 kHz.
    Invoke-FFmpeg "sample.wav" @("-f", "lavfi", "-i", "sine=frequency=440:sample_rate=96000:duration=2", "-ac", "2", "-c:a", "pcm_s24le")
    # Short video with a soundtrack.
    Invoke-FFmpeg "sample.mp4" @("-f", "lavfi", "-i", "testsrc2=size=320x240:rate=30:duration=2",
        "-f", "lavfi", "-i", "sine=frequency=880:sample_rate=48000:duration=2",
        "-c:v", "mpeg4", "-q:v", "5", "-pix_fmt", "yuv420p", "-c:a", "aac", "-shortest")
    Invoke-FFmpeg "sample.png" @("-f", "lavfi", "-i", "testsrc2=size=320x240", "-frames:v", "1")
    # A real MPEG transport stream, the other meaning of .ts.
    Invoke-FFmpeg "stream.ts" @("-f", "lavfi", "-i", "testsrc2=size=320x240:rate=30:duration=2",
        "-f", "lavfi", "-i", "sine=frequency=660:sample_rate=48000:duration=2",
        "-c:v", "mpeg2video", "-q:v", "5", "-pix_fmt", "yuv420p", "-c:a", "mp2", "-shortest", "-f", "mpegts")
    # Odd width and height, which H.264 and HEVC can't store as they are (issue #29).
    Invoke-FFmpeg "odd.webm" @("-f", "lavfi", "-i", "testsrc2=size=321x241:rate=30:duration=1",
        "-f", "lavfi", "-i", "sine=frequency=550:sample_rate=48000:duration=1",
        "-c:v", "libvpx", "-c:a", "libopus", "-shortest")
} elseif (Test-Path -LiteralPath (Join-Path $Samples "sample.pdf")) {
    New-SampleWithApp "sample.pdf" "pdf-png" "sample.png" "png-sample"
}

# Fonts: a TrueType font and a collection that come with Windows, plus that font as WOFF and WOFF2,
# made by the app itself (so those two formats are also read, not just written).
function Copy-FirstFont([string]$As, [string[]]$Names) {
    foreach ($n in $Names) {
        $p = Join-Path $env:WINDIR "Fonts\$n"
        if (Test-Path -LiteralPath $p -PathType Leaf) {
            Copy-Item -LiteralPath $p -Destination (Join-Path $Samples $As)
            return
        }
    }
    $script:SampleErrors[$As] = "none of $($Names -join ', ') in $env:WINDIR\Fonts"
}
Copy-FirstFont "sample.ttf" @("arial.ttf", "segoeui.ttf", "consola.ttf", "tahoma.ttf", "times.ttf")
Copy-FirstFont "sample.ttc" @("cambria.ttc", "msgothic.ttc", "simsun.ttc", "msyh.ttc", "meiryo.ttc", "YuGothR.ttc")
if (Test-Path -LiteralPath (Join-Path $Samples "sample.ttf")) {
    New-SampleWithApp "sample.ttf" "font-woff" "sample.woff" "font-sample"
    New-SampleWithApp "sample.ttf" "font-woff2" "sample.woff2" "font-sample"
}

# --- Hardware encoders -----------------------------------------------------------------------

if ($FFmpeg) {
    $working = @()
    foreach ($enc in "h264_nvenc", "hevc_nvenc", "h264_qsv", "hevc_qsv", "h264_amf", "hevc_amf") {
        $code = Invoke-Logged $FFmpeg @("-v", "error", "-f", "lavfi", "-i", "testsrc2=size=320x240:rate=30:duration=0.5", "-c:v", $enc, "-f", "null", "-") `
            (Join-Path $Out "encoder-$enc.log")
        if ($code -eq 0) { $working += $enc }
    }
    $list = "none"
    if ($working.Count -gt 0) { $list = $working -join " " }
    Write-Host "Hardware encoders here: $list (formats whose encoder is missing test the software fallback)"
    Write-Host ""
}

# --- Conversions -----------------------------------------------------------------------------

$NeedsFFmpeg = @("sample.mp4", "stream.ts", "odd.webm", "sample.wav")
$Documents = @("sample.md", "sample.txt", "script.ts", "sample.docx")

# Mirrors DocConv.NeedsLibreOffice: Word/ODT/RTF in or out needs it; Markdown and text to PDF, HTML
# or text don't.
function Test-NeedsLibreOffice([string]$Name, [string]$Target) {
    if ($Documents -notcontains $Name) { return $false }
    if ($Name -eq "sample.docx") { return $true }
    if ($Target -eq "pdf-doc" -or $Target -eq "html" -or $Target -eq "txt") { return $false }
    return $true
}

function Skip-OrFail-Missing([string]$Label, [string]$Name, [string]$Target) {
    if (-not $FFmpeg -and $NeedsFFmpeg -contains $Name -and -not $Strict) {
        Skip $Label $Name "ffmpeg not installed"
    } elseif (-not $LibreOffice -and (Test-NeedsLibreOffice $Name $Target) -and -not $RequireLibreOffice) {
        Skip $Label $Name "LibreOffice not installed"
    } elseif ($script:SampleErrors.ContainsKey($Name)) {
        Fail $Label $Name "sample couldn't be created: $($script:SampleErrors[$Name])"
    } else {
        Fail $Label $Name "sample couldn't be created (see the logs in $Out)"
    }
}

foreach ($name in "sample.mp4", "stream.ts", "odd.webm", "sample.wav", "sample.png", "sample.md", "sample.txt", "script.ts",
                  "sample.pdf", "sample.docx", "sample.ttf", "sample.ttc", "sample.woff", "sample.woff2") {
    $sample = Join-Path $Samples $name
    $exists = Test-Path -LiteralPath $sample -PathType Leaf
    # Ask the app which formats it offers for this file (same logic as the window). It may look
    # inside the file, so ask about the sample itself; an empty stand-in is only for a sample that
    # couldn't be generated.
    $probe = $sample
    if (-not $exists) {
        $probeDir = Join-Path $Out "probe"
        New-Item -ItemType Directory -Force -Path $probeDir | Out-Null
        $probe = Join-Path $probeDir $name
        [IO.File]::WriteAllBytes($probe, [byte[]]@())
    }
    $log = Join-Path $Runs "targets-$name.log"
    $code = Invoke-Logged $Exe @("--cli", "--targets", $probe) $log
    $targets = @(Read-Log $log | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($code -ne 0) {
        Fail "-" $name "app doesn't recognise this file type: $($targets -join ' ') (exit $code)"
        continue
    }

    # The two .ts files must be told apart by their contents.
    if ($name -eq "script.ts" -and $targets -notcontains "pdf-doc") {
        Fail "-" $name "a TypeScript file should be offered document formats, got: $($targets -join ' ')"
    } elseif ($name -eq "stream.ts" -and $exists -and $targets -notcontains "mp4-h264") {
        Fail "-" $name "a transport stream should be offered video formats, got: $($targets -join ' ')"
    }

    foreach ($target in $targets) {
        $null = $script:Covered.Add($target)
        if (-not $exists) {
            Skip-OrFail-Missing $target $name $target
            continue
        }
        if (-not $LibreOffice -and (Test-NeedsLibreOffice $name $target)) {
            if ($RequireLibreOffice) { Fail $target $name "LibreOffice not installed" } else { Skip $target $name "LibreOffice not installed" }
            continue
        }
        Invoke-Conversion $target $target $sample
        if ($HwTargets -contains $target) {
            Invoke-Conversion "$target+sw" $target $sample '{"hardware": false}'
        }
    }
}

# --- Fonts -----------------------------------------------------------------------------------
# Every converted font starts with the bytes of the format that was asked for. A collection becomes
# a folder with one font per file.

function Get-Head([string]$Path) {
    $bytes = New-Object byte[] 4
    $f = [IO.File]::OpenRead($Path)
    try { $n = $f.Read($bytes, 0, 4) } finally { $f.Dispose() }
    return (($bytes[0..($n - 1)] | ForEach-Object { $_.ToString("x2") }) -join "")
}

function Test-FontBytes([string]$Sample, [string]$Target, [string[]]$Heads) {
    $dir = Join-Path $Runs "$Sample-$Target"
    $label = "${Target}:bytes"
    if (-not (Test-Path -LiteralPath $dir)) { return }   # skipped
    $files = Get-Outputs $dir (Join-Path $dir $Sample)
    if ($files.Count -eq 0) { return }   # the conversion itself failed, and says so above
    $bad = @()
    foreach ($f in $files) {
        $head = Get-Head $f.FullName
        if ($Heads -notcontains $head) { $bad += "$($f.Name)=$head" }
    }
    if ($bad.Count -gt 0) {
        Fail $label $Sample "not the format asked for: $($bad -join ' ')"
    } elseif ($Sample -eq "sample.ttc" -and $files.Count -lt 2) {
        Fail $label $Sample "a collection should become several fonts, got $($files.Count)"
    } elseif ($files.Count -gt 0) {
        Pass $label $Sample "$($files.Count) file(s) start with $($Heads -join ' ')"
    }
}
foreach ($sample in "sample.ttf", "sample.ttc", "sample.woff", "sample.woff2") {
    Test-FontBytes $sample "font-woff2" @("774f4632")                          # wOF2
    Test-FontBytes $sample "font-woff" @("774f4646")                           # wOFF
    Test-FontBytes $sample "font-desktop" @("00010000", "4f54544f", "74727565")  # TrueType, OTTO, true
}

# --- Document -> PDF -------------------------------------------------------------------------
# The paper size setting is respected, and nothing but the OK line is printed. The file is given by
# a relative path, as from a shell in its folder.

# Page sizes in whole points. WebView2 writes the MediaBox in the clear; should a future version
# move it into a compressed object stream, the app renders the first page at 72 dpi instead, where
# pixels are points.
function Get-PageSizes([string]$Pdf, [string]$Dir) {
    $text = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($Pdf))
    $boxes = [regex]::Matches($text, '/MediaBox\s*\[\s*([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s+([-0-9.]+)\s*\]')
    if ($boxes.Count -gt 0) {
        $inv = [Globalization.CultureInfo]::InvariantCulture
        return @($boxes | ForEach-Object {
            $m = $_
            $v = @(1..4 | ForEach-Object { [double]::Parse($m.Groups[$_].Value, $inv) })
            "{0}x{1}" -f [Math]::Round($v[2] - $v[0]), [Math]::Round($v[3] - $v[1])
        }) | Select-Object -Unique
    }
    $render = Join-Path $Dir "render"
    New-Item -ItemType Directory -Force -Path $render | Out-Null
    Copy-Item -LiteralPath $Pdf -Destination $render
    $null = Invoke-Logged $Exe @("--cli", "pdf-png", (Join-Path $render (Split-Path $Pdf -Leaf))) (Join-Path $render "output.log") @{ CONVERT_SETTINGS = '{"pdfDPI": 72}' }
    $png = Get-ChildItem -LiteralPath $render -Recurse -Filter "*.png" | Sort-Object FullName | Select-Object -First 1
    if (-not $png) { return @("unknown (no MediaBox in the clear, and rendering the page failed)") }
    $b = [IO.File]::ReadAllBytes($png.FullName)
    $w = ([int]$b[16] -shl 24) + ([int]$b[17] -shl 16) + ([int]$b[18] -shl 8) + $b[19]
    $h = ([int]$b[20] -shl 24) + ([int]$b[21] -shl 16) + ([int]$b[22] -shl 8) + $b[23]
    return @("${w}x$h")
}

function Test-Paper([string]$Label, [string]$Folder, [string]$Settings, [string]$Size) {
    $dir = Join-Path $Runs $Folder
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    Copy-Item -LiteralPath (Join-Path $Samples "sample.md") -Destination $dir
    $log = Join-Path $dir "output.log"
    $code = Invoke-Logged $Exe @("--cli", "pdf-doc", "sample.md") $log @{ CONVERT_SETTINGS = $Settings } $dir
    $lines = Read-Log $log
    $pdf = Join-Path $dir "sample.pdf"
    $other = $lines | Where-Object { $_ -notlike "OK *" } | Select-Object -First 1
    if ($null -eq $code) {
        Fail $Label "sample.md" "timed out after ${TimeoutSec}s"
    } elseif ($code -ne 0 -or ($lines | Where-Object { $_ -like "FAIL*" })) {
        Fail $Label "sample.md" "$(Get-Reason $lines) (exit $code)"
    } elseif (-not (Test-Path -LiteralPath $pdf)) {
        Fail $Label "sample.md" "no sample.pdf written"
    } elseif ($null -ne $other) {
        Fail $Label "sample.md" "unexpected output: $other"
    } else {
        $sizes = @(Get-PageSizes $pdf $dir)
        if ($sizes.Count -ne 1 -or $sizes[0] -ne $Size) {
            Fail $Label "sample.md" "wrong paper size: $($sizes -join ', ') pt, expected $Size"
        } else {
            Pass $Label "sample.md" "$Size pt, nothing else printed"
        }
    }
}
Test-Paper "pdf-doc:a4" "paper-a4" '{"paperA4": true}' "595x842"
Test-Paper "pdf-doc:letter" "paper-letter" '{"paperA4": false}' "612x792"

# --- Read-only source folder -----------------------------------------------------------------
# When the original's folder can't be written to, the output goes to a fallback folder (Downloads;
# CONVERT_FALLBACK_DIR points it at a folder of ours) and the app says so. One conversion of every
# kind.

$RO = Join-Path $Out "read-only"
$Fallback = Join-Path $Out "fallback"
$RoCases = @("sample.png:jpg", "sample.wav:mp3", "sample.mp4:gif", "sample.md:pdf-doc", "sample.txt:docx",
             "sample.pdf:pdf-png", "sample.pdf:pdf-txt", "sample.pdf:pdf-docx", "sample.ttf:font-woff2")
New-Item -ItemType Directory -Force -Path $RO, $Fallback | Out-Null
foreach ($case in $RoCases) {
    $src = Join-Path $Samples $case.Split(":")[0]
    if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination $RO -Force }
}

# Read-only without elevation: deny ourselves creating files and folders in it (not inherited, so
# the samples inside stay readable). The owner can always change the ACL back.
$Sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$denied = (Invoke-Quiet icacls.exe $RO /deny "*${Sid}:(WD,AD)") -eq 0
try {
    if ($denied) {
        try {
            [IO.File]::WriteAllText((Join-Path $RO "probe.txt"), "")
            Remove-Item -LiteralPath (Join-Path $RO "probe.txt")
            $denied = $false
        } catch [UnauthorizedAccessException] { }
    }
    foreach ($case in $RoCases) {
        $name, $target = $case.Split(":")
        $label = "read-only:$target"
        $null = $script:Covered.Add($target)
        if (-not $denied) {
            Fail $label $name "couldn't make $RO read-only (icacls)"
            continue
        }
        if (-not (Test-Path -LiteralPath (Join-Path $RO $name))) {
            Skip-OrFail-Missing $label $name $target
            continue
        }
        if (-not $LibreOffice -and (Test-NeedsLibreOffice $name $target)) {
            if ($RequireLibreOffice) { Fail $label $name "LibreOffice not installed" } else { Skip $label $name "LibreOffice not installed" }
            continue
        }
        $log = Join-Path $Runs "read-only-$name-$target.log"
        $before = @(Get-Outputs $Fallback "").Count
        $code = Invoke-Logged $Exe @("--cli", $target, (Join-Path $RO $name)) $log @{ CONVERT_FALLBACK_DIR = $Fallback }
        $after = @(Get-Outputs $Fallback "").Count
        $lines = Read-Log $log
        if ($null -eq $code) {
            Fail $label $name "timed out after ${TimeoutSec}s"
        } elseif ($code -ne 0 -or ($lines | Where-Object { $_ -like "FAIL*" })) {
            Fail $label $name "$(Get-Reason $lines) (exit $code)"
        } elseif ($after -le $before) {
            Fail $label $name "nothing was written to the fallback folder"
        } elseif (-not ($lines | Where-Object { $_.Contains("saved to $Fallback") })) {
            Fail $label $name "the app didn't say where it saved the file"
        } else {
            Pass $label $name "saved to the fallback folder"
        }
    }
} finally {
    $null = Invoke-Quiet icacls.exe $RO /remove:d "*$Sid"
}

# Every format the app lists must have been tried by at least one sample.
foreach ($target in $AllTargets) {
    if (-not $script:Covered.Contains($target)) {
        Fail $target "-" "no sample covers this format; add one to windows\smoke-test.ps1"
    }
}

# --- Summary ---------------------------------------------------------------------------------

Write-Host ""
Write-Host "$($script:Passed) passed, $($script:Failures.Count) failed, $($script:Skipped) skipped"
if ($script:Failures.Count -gt 0) {
    Write-Host ""
    Write-Host "Failed formats:"
    foreach ($f in $script:Failures) { Write-Host "  $f" }
    Write-Host ""
    Write-Host "Samples, outputs and logs: $Out"
    exit 1
}
if (-not $Keep) {
    try { Remove-Item -LiteralPath $Out -Recurse -Force } catch { Write-Host "Couldn't remove $Out completely: $($_.Exception.Message)" }
}
exit 0
