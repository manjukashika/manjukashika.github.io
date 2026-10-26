param(
    [switch]$Watch
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

$workspaceDir = $PSScriptRoot
if (-not $workspaceDir) {
    $workspaceDir = "c:\Users\manju\Desktop\Academia\website_portfolio"
}

$excelPath = Join-Path $workspaceDir "GoldfishMemory.xlsx"
$jsOutPath = Join-Path $workspaceDir "goldfish-data.js"
$htmlPath  = Join-Path $workspaceDir "index.html"

function HtmlEncode([string]$str) {
    if ([string]::IsNullOrEmpty($str)) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($str.Trim())
}

function FormatResourceLinks([string]$raw) {
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return '<span style="color: var(--ink-muted); font-size: 0.82rem;">&mdash;</span>'
    }
    $trimmed = $raw.Trim()
    # Extract URLs if present
    $urlPattern = 'https?://[^\s,;]+'
    $matches = [regex]::Matches($trimmed, $urlPattern)
    if ($matches.Count -gt 0) {
        $linksHtml = @()
        $idx = 1
        foreach ($m in $matches) {
            $url = $m.Value.TrimEnd('.', ',', ';', ')')
            $label = "Link"
            if ($url -match 'youtube\.com|youtu\.be') { $label = "YouTube" }
            elseif ($url -match 'bit\.ly') { $label = "Playlist / Link" }
            elseif ($url -match 'github\.com') { $label = "GitHub" }
            elseif ($url -match 'mit\.edu') { $label = "MIT OCW" }
            else {
                try {
                    $uri = [System.Uri]$url
                    $hostName = $uri.Host -replace '^www\.', ''
                    $label = $hostName
                } catch {
                    $label = "Resource $idx"
                }
            }
            $safeUrl = [System.Net.WebUtility]::HtmlEncode($url)
            $safeLabel = [System.Net.WebUtility]::HtmlEncode($label)
            $linksHtml += "<a href=`"$safeUrl`" target=`"_blank`" rel=`"noopener`" class=`"resource-link`">$safeLabel &nearr;</a>"
            $idx++
        }
        return ($linksHtml -join " ")
    } else {
        return HtmlEncode $trimmed
    }
}

function Resolve-AnswerIndex([string]$ansRaw, [string[]]$options) {
    if ([string]::IsNullOrWhiteSpace($ansRaw)) { return 0 }
    $clean = $ansRaw.Trim()
    switch -Regex ($clean) {
        '^[Aa]$|^1$' { return 0 }
        '^[Bb]$|^2$' { return 1 }
        '^[Cc]$|^3$' { return 2 }
        '^[Dd]$|^4$' { return 3 }
    }
    for ($i = 0; $i -lt $options.Count; $i++) {
        if ($options[$i].Trim().ToLower() -eq $clean.ToLower()) {
            return $i
        }
    }
    return 0
}

function Parse-GoldfishExcel {
    if (-not (Test-Path $excelPath)) {
        Write-Warning "GoldfishMemory.xlsx not found at $excelPath"
        return $null
    }

    # Open with ReadWrite share so it works even while Excel is open
    $fs = New-Object System.IO.FileStream($excelPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)

    try {
        # 1. Shared Strings (explicitly read as UTF8 so math symbols like ∑ never get garbled)
        $sharedStrings = @()
        $ssEntry = $zip.GetEntry('xl/sharedStrings.xml')
        if ($ssEntry) {
            $sr = New-Object System.IO.StreamReader($ssEntry.Open(), [System.Text.Encoding]::UTF8)
            $ssXml = [xml]$sr.ReadToEnd()
            $sr.Close()
            foreach ($si in $ssXml.sst.si) {
                if ($null -ne $si.t) {
                    if ($si.t -is [string]) { $sharedStrings += $si.t }
                    else { $sharedStrings += $si.t.InnerText }
                } elseif ($null -ne $si.r) {
                    $parts = foreach ($r in $si.r) {
                        if ($r.t -is [string]) { $r.t } else { $r.t.InnerText }
                    }
                    $sharedStrings += ($parts -join '')
                } else {
                    $sharedStrings += ''
                }
            }
        }

        # 2. Styles (to detect medium-shade chapter rows even if unmerged)
        $xfTints = @{}
        $stylesEntry = $zip.GetEntry('xl/styles.xml')
        if ($stylesEntry) {
            $sr = New-Object System.IO.StreamReader($stylesEntry.Open(), [System.Text.Encoding]::UTF8)
            $stylesXml = [xml]$sr.ReadToEnd()
            $sr.Close()
            $sIdx = 0
            foreach ($xf in $stylesXml.styleSheet.cellXfs.xf) {
                $fillId = [int]$xf.fillId
                $fill = $stylesXml.styleSheet.fills.fill[$fillId]
                $fg = $fill.patternFill.fgColor
                $tintVal = ""
                if ($fg -and $fg.tint) { $tintVal = [string]$fg.tint }
                $xfTints[$sIdx] = $tintVal
                $sIdx++
            }
        }

        # 3. Workbook & Relationship mapping
        $wbEntry = $zip.GetEntry('xl/workbook.xml')
        $sr = New-Object System.IO.StreamReader($wbEntry.Open(), [System.Text.Encoding]::UTF8)
        $wbXml = [xml]$sr.ReadToEnd()
        $sr.Close()

        $relsEntry = $zip.GetEntry('xl/_rels/workbook.xml.rels')
        $sr = New-Object System.IO.StreamReader($relsEntry.Open(), [System.Text.Encoding]::UTF8)
        $relsXml = [xml]$sr.ReadToEnd()
        $sr.Close()

        $relMap = @{}
        foreach ($rel in $relsXml.Relationships.Relationship) {
            $relMap[$rel.Id] = $rel.Target
        }

        $subjects = @()
        $quizList = @()

        foreach ($sheet in $wbXml.workbook.sheets.sheet) {
            $sheetName = [string]$sheet.name
            $rId = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = $relMap[$rId]
            if (-not $target.StartsWith('xl/')) {
                $target = "xl/" + $target.TrimStart('/')
            }

            $sheetEntry = $zip.GetEntry($target)
            if (-not $sheetEntry) { continue }

            $sr = New-Object System.IO.StreamReader($sheetEntry.Open(), [System.Text.Encoding]::UTF8)
            $sheetXml = [xml]$sr.ReadToEnd()
            $sr.Close()

            $keyLower = $sheetName.Trim().ToLower()

            # Check if this sheet is the Quiz sheet
            if ($keyLower -eq 'quiz') {
                foreach ($row in $sheetXml.worksheet.sheetData.row) {
                    $rowNum = [int]$row.r
                    if ($rowNum -eq 1) { continue } # Skip header row (Question | A | B | C | D | correct option)

                    $cells = @{}
                    foreach ($c in $row.c) {
                        $cellRef = [string]$c.r
                        $col = $cellRef -replace '\d+', ''
                        $val = ""
                        if ($c.t -eq 's') {
                            $idx = [int]$c.v
                            if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) {
                                $val = $sharedStrings[$idx]
                            }
                        } elseif ($c.t -eq 'inlineStr' -and $c.is) {
                            $val = $c.is.InnerText
                        } elseif ($null -ne $c.v) {
                            $val = [string]$c.v
                        }
                        if (-not [string]::IsNullOrWhiteSpace($val)) {
                            $cells[$col] = $val.Trim()
                        }
                    }

                    if (-not $cells.ContainsKey('A')) { continue }
                    $qText = $cells['A']
                    if ($qText.ToLower() -eq 'question') { continue }

                    $optA = if ($cells['B']) { $cells['B'] } else { "" }
                    $optB = if ($cells['C']) { $cells['C'] } else { "" }
                    $optC = if ($cells['D']) { $cells['D'] } else { "" }
                    $optD = if ($cells['E']) { $cells['E'] } else { "" }
                    $ansRaw = if ($cells['F']) { $cells['F'] } else { "A" }

                    $opts = @($optA, $optB, $optC, $optD)
                    $ansIdx = Resolve-AnswerIndex -ansRaw $ansRaw -options $opts

                    $quizList += [PSCustomObject]@{
                        question    = $qText
                        options     = $opts
                        answerIndex = $ansIdx
                    }
                }
                continue
            }

            # Otherwise parse as a Subject Sheet
            $mergedRows = @{}
            if ($sheetXml.worksheet.mergeCells -and $sheetXml.worksheet.mergeCells.mergeCell) {
                foreach ($mc in $sheetXml.worksheet.mergeCells.mergeCell) {
                    $ref = [string]$mc.ref
                    if ($ref -match '^([A-Z]+)(\d+):([A-Z]+)(\d+)$') {
                        $startCol = $Matches[1]
                        $startRow = [int]$Matches[2]
                        $endCol   = $Matches[3]
                        if ($startCol -eq 'A' -and $endCol -ne 'A') {
                            $mergedRows[$startRow] = $true
                        }
                    }
                }
            }

            $headers = @{
                A = "Concept"
                B = "Definition"
                C = "Application"
                D = "Resources"
            }
            $chapters = @()
            $currentChapter = $null

            foreach ($row in $sheetXml.worksheet.sheetData.row) {
                $rowNum = [int]$row.r
                $cells = @{}
                $cellStyles = @{}

                foreach ($c in $row.c) {
                    $cellRef = [string]$c.r
                    $col = $cellRef -replace '\d+', ''
                    $styleIdx = if ($c.s) { [int]$c.s } else { 0 }
                    $cellStyles[$col] = $styleIdx

                    $val = ""
                    if ($c.t -eq 's') {
                        $idx = [int]$c.v
                        if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) {
                            $val = $sharedStrings[$idx]
                        }
                    } elseif ($c.t -eq 'inlineStr' -and $c.is) {
                        $val = $c.is.InnerText
                    } elseif ($null -ne $c.v) {
                        $val = [string]$c.v
                    }
                    if (-not [string]::IsNullOrWhiteSpace($val)) {
                        $cells[$col] = $val.Trim()
                    }
                }

                if ($cells.Count -eq 0) { continue }

                # Row 1 is Sheet Title
                if ($rowNum -eq 1) { continue }

                # Row 2 is Column Headers
                if ($rowNum -eq 2) {
                    if ($cells['A']) { $headers.A = $cells['A'] }
                    if ($cells['B']) { $headers.B = $cells['B'] }
                    if ($cells['C']) { $headers.C = $cells['C'] }
                    if ($cells['D']) { $headers.D = $cells['D'] }
                    continue
                }

                # Determine if this row is a Chapter Header (merged row starting at A, OR medium shade ~0.60 tint with only col A text)
                $isMergedBanner = $mergedRows.ContainsKey($rowNum)
                $aStyle = if ($cellStyles.ContainsKey('A')) { $cellStyles['A'] } else { 0 }
                $aTint = if ($xfTints.ContainsKey($aStyle)) { $xfTints[$aStyle] } else { "" }
                $isMediumShadeBanner = ($aTint -like "0.599*" -or $aTint -like "0.399*") -and $cells.ContainsKey('A') -and (-not $cells.ContainsKey('B')) -and (-not $cells.ContainsKey('C')) -and (-not $cells.ContainsKey('D'))

                if ($isMergedBanner -or $isMediumShadeBanner) {
                    $chapTitle = if ($cells['A']) { $cells['A'] } else { ($cells.Values | Select-Object -First 1) }
                    if (-not [string]::IsNullOrWhiteSpace($chapTitle)) {
                        $currentChapter = [PSCustomObject]@{
                            title = $chapTitle.Trim()
                            rows  = @()
                        }
                        $chapters += $currentChapter
                    }
                } else {
                    # Regular topic row
                    if ($null -eq $currentChapter) {
                        $currentChapter = [PSCustomObject]@{
                            title = "General Topics"
                            rows  = @()
                        }
                        $chapters += $currentChapter
                    }
                    $currentChapter.rows += [PSCustomObject]@{
                        concept     = if ($cells['A']) { $cells['A'] } else { "" }
                        definition  = if ($cells['B']) { $cells['B'] } else { "" }
                        application = if ($cells['C']) { $cells['C'] } else { "" }
                        resources   = if ($cells['D']) { $cells['D'] } else { "" }
                    }
                }
            }

            $subjKey = $keyLower -replace '\s+', '-'
            $colorCss = "#2E5E88"
            if ($keyLower -match 'math') {
                $subjKey = "maths"
                $colorCss = "#2E5E88"
            } elseif ($keyLower -match 'neuro') {
                $subjKey = "neuroimaging"
                $colorCss = "#4F6B46"
            } elseif ($keyLower -match 'method') {
                $subjKey = "methods"
                $colorCss = "#B58B24"
            }

            $subjects += [PSCustomObject]@{
                sheetName = $sheetName.Trim()
                key       = $subjKey
                color     = $colorCss
                headers   = $headers
                chapters  = $chapters
            }
        }

        return [PSCustomObject]@{
            subjects = $subjects
            quiz     = $quizList
        }
    } finally {
        $zip.Dispose()
        $fs.Close()
    }
}

function Update-WebsiteFiles {
    $parsed = Parse-GoldfishExcel
    if ($null -eq $parsed) { return }

    $subjects = $parsed.subjects
    $quizList = $parsed.quiz
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    # 1. Write goldfish-data.js (works on file:// and https://)
    $subjectsJson = ConvertTo-Json -InputObject @($subjects) -Depth 10
    $quizJson     = ConvertTo-Json -InputObject @($quizList) -Depth 10
    $jsContent = @"
// Auto-generated from GoldfishMemory.xlsx at $timestamp
window.GOLDFISH_EXCEL_DATA = $subjectsJson;
window.GOLDFISH_QUIZ_DATA = $quizJson;
"@
    [System.IO.File]::WriteAllText($jsOutPath, $jsContent, [System.Text.Encoding]::UTF8)

    # 2. Build static HTML for the Subject Groups so index.html also has the exact fallback markup
    $sb = New-Object System.Text.StringBuilder
    foreach ($subj in $subjects) {
        $chapCount = $subj.chapters.Count
        $subLabel = if ($chapCount -eq 1) { "1 Chapter • Expandable Table" } else { "$chapCount Chapters • Expandable Tables" }
        $safeSheetName = HtmlEncode $subj.sheetName

        [void]$sb.AppendLine("          <div class=`"subject-group`" data-subject=`"$($subj.key)`">")
        [void]$sb.AppendLine("            <div class=`"subject-group-header`">")
        [void]$sb.AppendLine("              <div class=`"subject-title-wrap`">")
        [void]$sb.AppendLine("                <h3 style=`"color: $($subj.color);`">$safeSheetName</h3>")
        [void]$sb.AppendLine("              </div>")
        [void]$sb.AppendLine("              <span class=`"subject-meta`">$subLabel</span>")
        [void]$sb.AppendLine("            </div>")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("            <div class=`"subsection-list`">")

        $cIdx = 0
        foreach ($chap in $subj.chapters) {
            $openAttr = if ($cIdx -eq 0) { " open" } else { "" }
            $rowCount = $chap.rows.Count
            $topicCountText = if ($rowCount -eq 1) { "1 Topic" } else { "$rowCount Topics" }
            $safeChapTitle = HtmlEncode $chap.title

            $hA = HtmlEncode $subj.headers.A
            $hB = HtmlEncode $subj.headers.B
            $hC = HtmlEncode $subj.headers.C
            $hD = HtmlEncode $subj.headers.D

            [void]$sb.AppendLine("              <details class=`"subsection-item`"$openAttr>")
            [void]$sb.AppendLine("                <summary class=`"subsection-summary`">")
            [void]$sb.AppendLine("                  <span class=`"summary-left`">")
            [void]$sb.AppendLine("                    <span class=`"summary-chevron`">&#9654;</span>")
            [void]$sb.AppendLine("                    $safeChapTitle")
            [void]$sb.AppendLine("                  </span>")
            [void]$sb.AppendLine("                  <span class=`"summary-count`">$topicCountText</span>")
            [void]$sb.AppendLine("                </summary>")
            [void]$sb.AppendLine("                <div class=`"table-scroll-wrap`">")
            [void]$sb.AppendLine("                  <table class=`"topic-table`">")
            [void]$sb.AppendLine("                    <thead>")
            [void]$sb.AppendLine("                      <tr>")
            [void]$sb.AppendLine("                        <th>$hA</th>")
            [void]$sb.AppendLine("                        <th>$hB</th>")
            [void]$sb.AppendLine("                        <th>$hC</th>")
            [void]$sb.AppendLine("                        <th>$hD</th>")
            [void]$sb.AppendLine("                      </tr>")
            [void]$sb.AppendLine("                    </thead>")
            [void]$sb.AppendLine("                    <tbody>")

            if ($rowCount -eq 0) {
                [void]$sb.AppendLine("                      <tr>")
                [void]$sb.AppendLine("                        <td colspan=`"4`" style=`"color: var(--ink-muted); font-style: italic;`">No topics added under this chapter yet.</td>")
                [void]$sb.AppendLine("                      </tr>")
            } else {
                foreach ($r in $chap.rows) {
                    $cText = if ($r.concept) { HtmlEncode $r.concept } else { "&mdash;" }
                    $dText = if ($r.definition) { HtmlEncode $r.definition } else { "&mdash;" }
                    $aText = if ($r.application) { HtmlEncode $r.application } else { "&mdash;" }
                    $rHtml = FormatResourceLinks $r.resources

                    [void]$sb.AppendLine("                      <tr>")
                    [void]$sb.AppendLine("                        <td class=`"topic-name`">$cText</td>")
                    [void]$sb.AppendLine("                        <td class=`"topic-def`">$dText</td>")
                    [void]$sb.AppendLine("                        <td class=`"topic-app`">$aText</td>")
                    [void]$sb.AppendLine("                        <td class=`"topic-res`">$rHtml</td>")
                    [void]$sb.AppendLine("                      </tr>")
                }
            }

            [void]$sb.AppendLine("                    </tbody>")
            [void]$sb.AppendLine("                  </table>")
            [void]$sb.AppendLine("                </div>")
            [void]$sb.AppendLine("              </details>")
            $cIdx++
        }

        [void]$sb.AppendLine("            </div>")
        [void]$sb.AppendLine("          </div>")
    }

    # Update index.html between <!-- EXCEL_SUBJECTS_START --> and <!-- EXCEL_SUBJECTS_END --> if markers exist
    if (Test-Path $htmlPath) {
        $htmlRaw = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
        $pattern = '(?s)(<!-- EXCEL_SUBJECTS_START -->).*?(<!-- EXCEL_SUBJECTS_END -->)'
        if ($htmlRaw -match $pattern) {
            $replacement = "`$1`r`n" + $sb.ToString().TrimEnd() + "`r`n          `$2"
            $htmlRaw = [regex]::Replace($htmlRaw, $pattern, $replacement)
        }
        $quizPattern = '(?s)(// <!-- EXCEL_QUIZ_START -->).*?(// <!-- EXCEL_QUIZ_END -->)'
        if ($htmlRaw -match $quizPattern) {
            $quizReplacement = "`$1`r`n    let quizBank = $quizJson;`r`n    `$2"
            $htmlRaw = [regex]::Replace($htmlRaw, $quizPattern, $quizReplacement)
        }
        [System.IO.File]::WriteAllText($htmlPath, $htmlRaw, [System.Text.Encoding]::UTF8)
    }

    Write-Host "[$timestamp] Synced GoldfishMemory.xlsx -> goldfish-data.js & index.html ($($subjects.Count) subject sheets, $($quizList.Count) quiz questions)"
}

Update-WebsiteFiles

if ($Watch) {
    Write-Host "Watching GoldfishMemory.xlsx for changes... (Press Ctrl+C to stop)"
    $lastWrite = (Get-Item $excelPath).LastWriteTimeUtc
    while ($true) {
        Start-Sleep -Milliseconds 900
        if (Test-Path $excelPath) {
            $currWrite = (Get-Item $excelPath).LastWriteTimeUtc
            if ($currWrite -ne $lastWrite) {
                Start-Sleep -Milliseconds 350
                try {
                    Update-WebsiteFiles
                    $lastWrite = $currWrite
                } catch {
                    # Wait for Excel save lock to settle
                }
            }
        }
    }
}
