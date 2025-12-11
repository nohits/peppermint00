$newFilePath3 = '.\hometemplate.html'
$htmlContent = Get-Content -Path .\hometemplate11.html -Raw
$targetDate = '2025-11-01'
$matchupList = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?sportId=1&date=$targetDate").dates.games

$postHtml = ""
$i = 0
foreach ($matchup in $matchupList) {
    $i++
    $href = '$Link' + $i + 'game'
    $aLabel = '$' + $i + 'game'
    $divDate = '$Date' + $i + 'game'

    $postHtml += @"
  <div class="post">
    <h3><a href="$($href)" title="Permalink to this article">$($aLabel)</a></h3>

    <div class="postinfo">
      <p class="published" title="2015-08-23T00:00:00-07:00">$($divDate)</p>
    </div>
  </div>

"@
}

$pattern = '(?s)(?<=<div id="postlist">).*?(?=</div>)'

if ($htmlContent -match '<div id="postlist">') {
    $updatedHtml1 = [regex]::Replace($htmlContent, $pattern, "`r`n$postHtml`r`n", 'Singleline')
    Set-Content -Path $newFilePath3 -Value $updatedHtml1
    Write-Host "Inserted $($matchupList.Count) posts into <div id='postlist'>."
} else {
    Write-Host "Could not find <div id='postlist'> in the HTML file."
}

