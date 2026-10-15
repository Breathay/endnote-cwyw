[CmdletBinding()]
param(
    [ValidateSet('Probe','Format')][string]$Action = 'Probe',
    [string]$Document,
    [string]$StylePath,
    [int]$ExpectedCitations = -1,
    [int]$ExpectedReferences = -1,
    [ValidateRange(5,300)][int]$TimeoutSeconds = 60,
    [string]$PdfPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false

# GetActiveObject is available in Windows PowerShell 5.1, not in .NET-based
# PowerShell 7. Re-execute the same source with safely quoted arguments; do not
# alter execution policy, start a shell UI, or install another PowerShell.
if ($PSVersionTable.PSVersion.Major -ge 6) {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw 'Native Word CWYW requires Windows.'
    }
    function Quote-Literal([object]$Value) { return "'" + ([string]$Value).Replace("'", "''") + "'" }
    $taskEngine = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $taskSource = [IO.File]::ReadAllText($PSCommandPath, [Text.Encoding]::UTF8)
    $taskArguments = ''
    foreach ($taskKey in $PSBoundParameters.Keys) {
        $taskArguments += ' -' + $taskKey + ' ' + (Quote-Literal $PSBoundParameters[$taskKey])
    }
    $taskCommand = "& {`n" + $taskSource + "`n}" + $taskArguments
    & $taskEngine -NoProfile -Command $taskCommand
    exit $LASTEXITCODE
}

function Get-ActiveWord {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
    catch { return $null }
}

function Get-CwywState($Doc) {
    $taskCites = 0
    $taskBibliographies = 0
    $taskBibliographyText = ''
    foreach ($taskField in $Doc.Fields) {
        $taskCode = [string]$taskField.Code.Text
        if ($taskCode -match '^\s*ADDIN\s+EN\.CITE(?:\s|$)') { $taskCites++ }
        if ($taskCode -match '^\s*ADDIN\s+EN\.REFLIST\b') {
            $taskBibliographies++
            $taskBibliographyText += [string]$taskField.Result.Text
        }
    }
    $taskRefs = 0
    try {
        [xml]$taskLibraries = $Doc.Variables.Item('EN.Libraries').Value
        $taskRefs = @($taskLibraries.SelectNodes('//record-ids/item')).Count
    } catch { $taskRefs = 0 }
    $taskUnresolved = [regex]::Matches([string]$Doc.Content.Text, '\{[^{}]*#\d+[^{}]*\}').Count
    $taskHasher = [Security.Cryptography.SHA256]::Create()
    try {
        $taskDigest = [BitConverter]::ToString($taskHasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($taskBibliographyText)))
    } finally { $taskHasher.Dispose() }
    return [pscustomobject]@{
        citations = $taskCites
        bibliographies = $taskBibliographies
        references = $taskRefs
        unresolved = $taskUnresolved
        bibliography_chars = $taskBibliographyText.Length
        bibliography_sha256 = $taskDigest
    }
}

$taskWord = Get-ActiveWord
if ($Action -eq 'Probe') {
    $taskDocuments = @()
    $taskConnected = $false
    $taskTemplates = @()
    if ($null -ne $taskWord) {
        foreach ($taskAddin in $taskWord.COMAddIns) {
            if ($taskAddin.ProgId -eq 'EndNote.WordAddin.Connect') { $taskConnected = [bool]$taskAddin.Connect }
        }
        foreach ($taskTemplate in $taskWord.Templates) {
            if ($taskTemplate.Name -like 'EndNote Cwyw*.dotm') { $taskTemplates += $taskTemplate.FullName }
        }
        foreach ($taskDoc in $taskWord.Documents) {
            $taskDocuments += [pscustomobject]@{ name=$taskDoc.Name; path=$taskDoc.FullName }
        }
    }
    [pscustomobject]@{
        action='probe'; word_running=($null -ne $taskWord); cwyw_connected=$taskConnected
        cwyw_templates=$taskTemplates; documents=$taskDocuments
        style_exists=([bool]$StylePath -and (Test-Path -LiteralPath $StylePath -PathType Leaf))
        word_mutated=$false
    } | ConvertTo-Json -Depth 5
    exit 0
}

if (-not $Document -or -not $StylePath) { throw 'Format requires -Document and an installed -StylePath.' }
$taskDocumentPath = (Resolve-Path -LiteralPath $Document).Path
$taskStylePath = (Resolve-Path -LiteralPath $StylePath).Path
if ([IO.Path]::GetExtension($taskDocumentPath) -ne '.docx') { throw 'The prepared output must be .docx.' }
if ([IO.Path]::GetExtension($taskStylePath) -ne '.ens') { throw 'StylePath must be an installed .ens file.' }
$taskStyleName = [IO.Path]::GetFileNameWithoutExtension($taskStylePath)
if ($null -eq $taskWord) { $taskWord = New-Object -ComObject Word.Application }
$taskWord.Visible = $true
$taskAddin = $taskWord.COMAddIns.Item('EndNote.WordAddin.Connect')
if (-not $taskAddin.Connect) { throw 'Enable the installed EndNote CWYW add-in in Word before formatting.' }
$taskDoc = $null
foreach ($taskCandidate in $taskWord.Documents) {
    if ([string]::Equals($taskCandidate.FullName,$taskDocumentPath,[StringComparison]::OrdinalIgnoreCase)) {
        $taskDoc = $taskCandidate
        break
    }
}
if ($null -eq $taskDoc) { $taskDoc = $taskWord.Documents.Open($taskDocumentPath) }
if ($taskDoc.ReadOnly) { throw 'The prepared output is read-only or open in a different Word instance.' }
$taskDoc.Activate()

$taskLayoutVar = $null
foreach ($taskVariable in $taskDoc.Variables) {
    if ($taskVariable.Name -eq 'EN.Layout') { $taskLayoutVar=$taskVariable;break }
}
if ($null -ne $taskLayoutVar) {
    [xml]$taskLayout = $taskLayoutVar.Value
    if ($null -eq $taskLayout.ENLayout.Style) { throw 'Existing EN.Layout has no Style; inspect it before changing settings.' }
    $taskLayout.ENLayout.Style = $taskStyleName
    $taskLayoutVar.Value = $taskLayout.OuterXml
} else {
    # Observed CWYW layout schema. Existing documents retain their own fonts
    # and bibliography preferences; these defaults apply only on first format.
    [xml]$taskLayout = '<ENLayout><Style></Style><LeftDelim>{</LeftDelim><RightDelim>}</RightDelim><FontName>Times New Roman</FontName><FontSize>11</FontSize><ReflistTitle></ReflistTitle><StartingRefnum>1</StartingRefnum><FirstLineIndent>0</FirstLineIndent><HangingIndent>720</HangingIndent><LineSpacing>0</LineSpacing><SpaceAfter>0</SpaceAfter><HyperlinksEnabled>0</HyperlinksEnabled><HyperlinksVisible>0</HyperlinksVisible><EnableBibliographyCategories>0</EnableBibliographyCategories></ENLayout>'
    $taskLayout.ENLayout.Style = $taskStyleName
    [void]$taskDoc.Variables.Add('EN.Layout',$taskLayout.OuterXml)
}

# This is the native CWYW entry point. Never substitute an invented macro,
# create imitation ADDIN fields, or call a *Dlg macro that waits on a modal UI.
$taskWord.Run('ENFormatBibliography')
$taskDeadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
$taskPrevious = ''
$taskStable = 0
$taskReady = $false
do {
    try {
        $taskState = Get-CwywState $taskDoc
        $taskValid = $taskState.bibliographies -eq 1 -and $taskState.bibliography_chars -gt 0 -and $taskState.unresolved -eq 0
        if ($ExpectedCitations -ge 0) { $taskValid = $taskValid -and $taskState.citations -eq $ExpectedCitations }
        if ($ExpectedReferences -ge 0) { $taskValid = $taskValid -and $taskState.references -eq $ExpectedReferences }
        $taskSignature = $taskState | ConvertTo-Json -Compress
        if ($taskValid -and $taskSignature -eq $taskPrevious) { $taskStable++ } else { $taskStable=0 }
        $taskPrevious = $taskSignature
        if ($taskStable -ge 2) { $taskReady=$true;break }
    } catch [Runtime.InteropServices.COMException] {
        # Word may reject a read briefly while CWYW is updating the document.
        $taskStable = 0
    }
    Start-Sleep -Milliseconds 750
} while ([DateTime]::UtcNow -lt $taskDeadline)
if (-not $taskReady) {
    throw 'CWYW did not reach a stable complete bibliography. Word was left open for inspection; do not rerun blindly.'
}
$taskDoc.Save()
if ($PdfPath) {
    $taskPdf = [IO.Path]::GetFullPath($PdfPath)
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($taskPdf))
    $taskDoc.Repaginate()
    $taskDoc.ExportAsFixedFormat($taskPdf,17)
}
[pscustomobject]@{
    action='format';document=$taskDocumentPath;style=$taskStyleName
    native_citations=$taskState.citations;native_bibliographies=$taskState.bibliographies
    traveling_library_records=$taskState.references;unresolved=$taskState.unresolved
    stable=$true;saved=[bool]$taskDoc.Saved
} | ConvertTo-Json -Depth 4
# Keep the user's Word documents and EndNote library open.
