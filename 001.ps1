$targetDate = '2025-11-01'
#$targetDate = Get-Date -f yyyy-MM-dd
$matchupList = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?sportId=1&date=$targetDate").dates.games
$i = 0

foreach ($game in $matchupList) {
    $secondDate = Get-Date -Format "d" $game.gameDate
    $fiveDaysBack = (Get-Date $secondDate).AddDays(-6)
    $tenDaysBack = (Get-Date $secondDate).AddDays(-10)
    $formatDateFiveDays = Get-Date -Format "d" $fiveDaysBack
    $formatDateTenDays = Get-Date -Format "d" $tenDaysBack
    $gameId = $game.gamePk
    $gameInfo = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$gameId/feed/live").gameData
    $weather = $gameInfo.weather
    $awayTeam = $game.teams.away.team.name
    $awayTeamId = $game.teams.away.team.id
    $homeTeam = $game.teams.home.team.name
    $homeTeamId = $game.teams.home.team.id
    $Matchup = $awayTeam + ' @ ' + $homeTeam
    $i++

    if (-not [string]::IsNullOrEmpty($gameInfo.probablePitchers.away.fullName)) {
        $awaySpId = $gameInfo.probablePitchers.away.id
        $awaySpStats = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=season&group=pitching&season=2025" -ErrorAction SilentlyContinue).stats.splits.stat | Select-Object -First 1
    }
    if (-not [string]::IsNullOrEmpty($gameInfo.probablePitchers.home.fullName)) {
        $homeSpId = $gameInfo.probablePitchers.home.id
        $homeSpStats = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=season&group=pitching" -ErrorAction SilentlyContinue).stats.splits.stat | Select-Object -First 1
    }
    if (-not [string]::IsNullOrEmpty($gameInfo.probablePitchers.away.fullName) -and $null -ne $awaySpStats) {
        $awayTeamSP = $gameInfo.probablePitchers.away.fullName
        $awaySpProfile = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId").people
        $awaySpCaStats = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=career&group=pitching").stats.splits.stat
        $awaySpGameLog = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=gameLog&group=pitching&seasons=2025,2024,2023").stats.splits
        $awaySpHand = $awaySpProfile.pitchHand.code + 'HP'
        $awaySpArs = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=pitchArsenal&group=pitching").stats.splits.stat | Sort-Object -Property percentage -Descending
        $awaySpArs1 = $awaySpArs[0].type.description 
        $awaySpArs2 = $awaySpArs[1].type.description 
        $awaySpArs3 = $awaySpArs[2].type.description 
        $awaySpSpin = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=metricAverages&group=pitching&metrics=releaseSpinRate,releaseExtension,releaseSpeed,effectiveSpeed,launchSpeed,launchAngle").stats.splits.stat
        $awaySpSpin1 = $awaySpSpin | Where-Object {$_.event.details.type.description -match $awaySpArs[0].type.description}
        $awaySpSpin2 = $awaySpSpin | Where-Object {$_.event.details.type.description -match $awaySpArs[1].type.description}
        $awaySpSpin3 = $awaySpSpin | Where-Object {$_.event.details.type.description -match $awaySpArs[2].type.description}
        $awaySpAdv = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=seasonAdvanced&group=pitching").stats.splits.stat[0]
        $awaySpBar = ($awaySpAdv.lineOuts + $awaySpAdv.lineHits + $awaySpAdv.flyHits) / $awaySpAdv.totalSwings
        $awaySpvsHomeTm = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=gameLog&group=pitching&seasons=2025,2024,2023,2022").stats.splits | Where-Object {$_.opponent.id -eq $homeTeamId}
        $awaySpFip = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=statSplits&group=pitching&gameType=R&sitCodes=fip").stats.splits.stat | Select-Object -First 1
        $awaySpFP = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=statSplits&group=pitching&gameType=R&sitCodes=fp").stats.splits.stat.strikePercentage | Select-Object -First 1
        $awaySpRon = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=statSplits&group=pitching&gameType=R&sitCodes=ron").stats.splits.stat | Select-Object -First 1
        $awaySpTotalPitches = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=season&group=pitching&gameType=R").stats.splits.stat.numberOfPitches | Select-Object -First 1
        $awaySpZn05 = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=statSplits&group=pitching&gameType=R&sitCodes=zn05").stats.splits.stat | Select-Object -First 1
        $awaySpMeatballPerc = ($awaySpZn05.strikes / $awaySpTotalPitches).ToString(".000")
        $awaySpDY = [string]$awaySpProfile.draftYear + "?playerId=" + $awaySpId 
        $awaySpDraft = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/draft/$awaySpDY").drafts.rounds.picks
        $awaySpMeta = 'AGE ' + $awaySpProfile.currentAge + ', PICK #' + $awaySpDraft.pickNumber + ' ' + $awaySpDraft.year

        if ($awaySpvsHomeTm.Count -gt 0) {
            $aSpvs = '{0:M/d/yy}' -f [datetime]$awaySpvsHomeTm[-1].date + ' vs ' + $awaySpvsHomeTm[-1].opponent.name.split()[-1]
            $aSpvsD = $awaySpvsHomeTm[-1].stat.summary + ', ' + $awaySpvsHomeTm[-1].stat.hits + ' H'

        }
        else {
            $aSpvs = "never faced $homeTeam"
            $aSpvsD = ''
        }
        if ($awaySpvsHomeTm.Count -gt 1) {
            $aSpvs2 = '{0:M/d/yy}' -f [datetime]$awaySpvsHomeTm[-2].date + ' vs ' + $awaySpvsHomeTm[-2].opponent.name.split()[-1]
            $aSpvsD2 = $awaySpvsHomeTm[-2].stat.summary + ', ' + $awaySpvsHomeTm[-2].stat.hits + ' H' 
        }
        else {
            $aSpvs2 = "."
            $aSpvsD2 = ''
        }

        if ($awaySpGameLog[-1].isHome -eq 'False') {
            $as1 = '{0:MM/dd}' -f [datetime]$awaySpGameLog[-1].date + ' vs ' + $awaySpGameLog[-1].opponent.name.split()[-1]
        }
        else { 
            $as1 = '{0:MM/dd}' -f [datetime]$awaySpGameLog[-1].date + '  @ ' + $awaySpGameLog[-1].opponent.name.split()[-1]
        }
        if ($awaySpGameLog[-2].isHome -eq 'False') {
            $as2 = '{0:MM/dd}' -f [datetime]$awaySpGameLog[-2].date + ' vs ' + $awaySpGameLog[-2].opponent.name.split()[-1]
        }
        else { 
            $as2 = '{0:MM/dd}' -f [datetime]$awaySpGameLog[-2].date + '  @ ' + $awaySpGameLog[-2].opponent.name.split()[-1]
        }
        if ($awaySpGameLog[-3].isHome -eq 'False') {
            $as3 = '{0:MM/dd}' -f [datetime]$awaySpGameLog[-3].date + ' vs ' + $awaySpGameLog[-3].opponent.name.split()[-1]
        }
        else { 
            $as3 = '{0:MM/dd}' -f [datetime]$awaySpGameLog[-3].date + '  @ ' + $awaySpGameLog[-3].opponent.name.split()[-1]
        }
        if ($awaySpGameLog[-4].isHome -eq 'False') {
            $as4 = '{0:MM/dd}' -f [datetime]$awaySpGameLog[-4].date + ' vs ' + $awaySpGameLog[-4].opponent.name.split()[-1]
        } else { 
            $as4 = '{0:MM/dd}' -f [datetime]$awaySpGameLog[-4].date + '  @ ' + $awaySpGameLog[-4].opponent.name.split()[-1]
        }


        $homeTmvsAwaySp = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awaySpId/stats?stats=vsPlayer5Y&group=pitching&opposingTeamId=$homeTeamId&rosterType=Active").stats.splits | Sort-Object -Descending {$_.stat.atBats}
        $homeRoster = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/$homeTeamId/roster").roster
        $homeTmHittersss = $homeRoster | Where-Object { $_.position -notmatch 'Pitcher'}
        $homeTmIds = $homeTmHittersss.person.id -join ','
        $homeTmvsSp = $homeTmvsAwaySp | Where-Object { $_.stat.atBats -ge 5 -and $homeTmHittersss.person.id -contains $_.batter.id}

        if ($homeTmvsSp.Count -gt 0) {
            $homeTmvsSpNm1 = '..' + $homeTmvsSp[-1].batter.fullName
            $homeTmvsSpSt1 = "{0:d2}" -f $homeTmvsSp[-1].stat.atBats + ' AB, ' + $homeTmvsSp[-1].stat.hits + ' H, ' + $homeTmvsSp[-1].stat.baseOnBalls + ' BB, ' + $homeTmvsSp[-1].stat.homeRuns + ' HR, ' + 
                             $homeTmvsSp[-1].stat.strikeOuts + ' K'
        }
        if ($homeTmvsSp.Count -ge 1) {
            $homeTmvsSpNm2 = '..' + $homeTmvsSp[-2].batter.fullName
            $homeTmvsSpSt2 = "{0:d2}" -f $homeTmvsSp[-2].stat.atBats + ' AB, ' + $homeTmvsSp[-2].stat.hits + ' H, ' + $homeTmvsSp[-2].stat.baseOnBalls + ' BB, ' + $homeTmvsSp[-2].stat.homeRuns + ' HR, ' + 
                             $homeTmvsSp[-2].stat.strikeOuts + ' K'
        }
        if ($homeTmvsSp.Count -eq 0 -or [string]::IsNullOrEmpty($homeTmvsSp)) {
            $homeTmvsSpNm1 = 'na2'
            $homeTmvsSpNm2 = 'na1'
        }
    }

    if ($homeTmvsSp.Count -eq 0 -or [string]::IsNullOrEmpty($homeTmvsSp)) {
        $homeTmvsSpNm1 = 'na2'
        $homeTmvsSpNm2 = 'na1'
    }

    if (-not [string]::IsNullOrEmpty($gameInfo.probablePitchers.home.fullName) -and $null -ne $homeSpStats) {
        $homeTeamSP = $gameInfo.probablePitchers.home.fullName
        $homeSpProfile = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId").people
        $homeSpCaStats = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=career&group=pitching").stats.splits.stat
        $homeSpGameLog = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=gameLog&group=pitching&seasons=2025,2024,2023").stats.splits
        $homeSpHand = $homeSpProfile.pitchHand.code + 'HP'
        $homeSpArs = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=pitchArsenal&group=pitching").stats.splits.stat | Sort-Object -Property percentage -Descending
        $homeSpArs1 = $homeSpArs[0].type.description
        $homeSpArs2 = $homeSpArs[1].type.description 
        $homeSpArs3 = $homeSpArs[2].type.description
        $homeSpSpin = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=metricAverages&group=pitching&metrics=releaseSpinRate,releaseExtension,releaseSpeed,effectiveSpeed,launchSpeed,launchAngle").stats.splits.stat
        $homeSpSpin1 = $homeSpSpin | Where-Object {$_.event.details.type.description -match $homeSpArs[0].type.description}
        $homeSpSpin2 = $homeSpSpin | Where-Object {$_.event.details.type.description -match $homeSpArs[1].type.description}
        $homeSpSpin3 = $homeSpSpin | Where-Object {$_.event.details.type.description -match $homeSpArs[2].type.description}
        $homeSpAdv = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=seasonAdvanced&group=pitching").stats.splits.stat[0]
        $homeSpBar = ($homeSpAdv.lineOuts + $homeSpAdv.lineHits + $homeSpAdv.flyHits) / $homeSpAdv.totalSwings
        $homeSpvsAwayTm = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=gameLog&group=pitching&seasons=2025,2024,2023,2022").stats.splits | Where-Object {$_.opponent.id -eq $awayTeamId}
        $homeSpFip = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=statSplits&group=pitching&gameType=R&sitCodes=fip").stats.splits.stat | Select-Object -First 1
        $homeSpFP = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=statSplits&group=pitching&gameType=R&sitCodes=fp").stats.splits.stat.strikePercentage | Select-Object -First 1
        $homeSpRon = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=statSplits&group=pitching&gameType=R&sitCodes=ron").stats.splits.stat | Select-Object -First 1
        $homeSpTotalPitches = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=season&group=pitching&gameType=R").stats.splits.stat.numberOfPitches | Select-Object -First 1
        $homeSpZn05 = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=statSplits&group=pitching&gameType=R&sitCodes=zn05").stats.splits.stat | Select-Object -First 1
        $homeSpMeatballPerc = ($homeSpZn05.strikes / $homeSpTotalPitches).ToString(".000")
        #579328
        $homeSpDY = [string]$homeSpProfile.draftYear + "?playerId=" + $homeSpId 
        $homeSpDraft = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/draft/$homeSpDY").drafts.rounds.picks
        $homeSpMeta = 'AGE ' + $homeSpProfile.currentAge + ', PICK #' + $homeSpDraft.pickNumber + ' ' + $homeSpDraft.year

        if ($null -ne $homeSpvsAwayTm) {
            $hSpvs = '{0:M/d/yy}' -f [datetime]$homeSpvsAwayTm[-1].date + ' vs ' + $homeSpvsAwayTm[-1].opponent.name.split()[-1]
            $hSpvsD = $homeSpvsAwayTm[-1].stat.summary + ', ' + $homeSpvsAwayTm[-1].stat.hits + ' H'
        } else {
            $hSpvs = "never faced $awayTeam"
            $hSpvsD = ''
        }

        if ($homeSpvsAwayTm.Count -ge 2) {
            $hSpvs2 = '{0:M/d/yy}' -f [datetime]$homeSpvsAwayTm[-2].date + ' vs ' + $homeSpvsAwayTm[-2].opponent.name.split()[-1]
            $hSpvsD2 = $homeSpvsAwayTm[-2].stat.summary + ', ' + $homeSpvsAwayTm[-2].stat.hits + ' H' 
        } else {
            $hSpvs2 = "."
            $hSpvsD2 = ''
        }

        if ($homeSpGameLog[-1].isHome -eq 'False') {
            $hs1 = '{0:MM/dd}' -f [datetime]$homeSpGameLog[-1].date + ' vs ' + $homeSpGameLog[-1].opponent.name.split()[-1]
        } else { 
            $hs1 = '{0:MM/dd}' -f [datetime]$homeSpGameLog[-1].date + '  @ ' + $homeSpGameLog[-1].opponent.name.split()[-1]
        }
        if ($homeSpGameLog[-2].isHome -eq 'False') {
            $hs2 = '{0:MM/dd}' -f [datetime]$homeSpGameLog[-2].date + ' vs ' + $homeSpGameLog[-2].opponent.name.split()[-1] #| .\ErrorDetails.dll
        } else { 
            $hs2 = '{0:MM/dd}' -f [datetime]$homeSpGameLog[-2].date + '  @ ' + $homeSpGameLog[-2].opponent.name.split()[-1]
        }
        if ($homeSpGameLog[-3].isHome -eq 'False') {
            $hs3 = '{0:MM/dd}' -f [datetime]$homeSpGameLog[-3].date + ' vs ' + $homeSpGameLog[-3].opponent.name.split()[-1]
        } else { 
            $hs3 = '{0:MM/dd}' -f [datetime]$homeSpGameLog[-3].date + '  @ ' + $homeSpGameLog[-3].opponent.name.split()[-1]
        }
        if ($homeSpGameLog[-4].isHome -eq 'False') {
            $hs4 = '{0:MM/dd}' -f [datetime]$homeSpGameLog[-4].date + ' vs ' + $homeSpGameLog[-4].opponent.name.split()[-1]
        } else { 
            $hs4 = '{0:MM/dd}' -f [datetime]$homeSpGameLog[-4].date + '  @ ' + $homeSpGameLog[-4].opponent.name.split()[-1]
        }


        $awayTmvsHomeSp = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeSpId/stats?stats=vsPlayer5Y&group=pitching&opposingTeamId=$awayTeamId&rosterType=Active").stats.splits | Sort-Object -Descending {$_.stat.atBats}
        $awayRoster = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/$awayTeamId/roster").roster
        $awayTmHittersss = $awayRoster | Where-Object {$_.position -notmatch 'Pitcher'}
        $awayTmIds = $awayTmHittersss.person.id -join ','
        $awayTmvsSp = $awayTmvsHomeSp | Where-Object {$_.stat.atBats -ge 5 -and $awayTmHittersss.person.id -contains $_.batter.id}

        if ($awayTmvsSp.Count -gt 0) {
            $awayTmvsSpNm1 = '..' + $awayTmvsSp[-1].batter.fullName
            $awayTmvsSpSt1 = "{0:d2}" -f $awayTmvsSp[-1].stat.atBats + ' AB, ' + $awayTmvsSp[-1].stat.hits + ' H, ' + $awayTmvsSp[-1].stat.baseOnBalls + ' BB, ' + $awayTmvsSp[-1].stat.homeRuns + ' HR, ' + 
                             $awayTmvsSp[-1].stat.strikeOuts + ' K'
        }
        if ($awayTmvsSp.Count -gt 1) {
            $awayTmvsSpNm2 = '..' + $awayTmvsSp[-2].batter.fullName
            $awayTmvsSpSt2 = "{0:d2}" -f $awayTmvsSp[-2].stat.atBats + ' AB, ' + $awayTmvsSp[-2].stat.hits + ' H, ' + $awayTmvsSp[-2].stat.baseOnBalls + ' BB, ' + $awayTmvsSp[-2].stat.homeRuns + ' HR, ' + 
                             $awayTmvsSp[-2].stat.strikeOuts + ' K'
        }
        if ($awayTmvsSp.Count -eq 0 -or [string]::IsNullOrEmpty($awayTmvsSp)) {
            $awayTmvsSpNm1 = 'na2'
            $awayTmvsSpNm2 = 'na1'
        }
    }

    if ($awayTmvsSp.Count -eq 0 -or [string]::IsNullOrEmpty($awayTmvsSp)) {
        $awayTmvsSpNm1 = 'na2'
        $awayTmvsSpNm2 = 'na1'
    }


    $awayTeamStat = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/$awayTeamId/stats?stats=season&group=hitting").stats.splits.stat
    $awayTeamRecents = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?stats=byDateRange&sportId=1&teamId=$awayTeamId&startDate=$formatDateFiveDays&endDate=$secondDate").dates.games
    $awayTeamRecord = [string]$game.teams.away.leagueRecord.wins + '-' + $game.teams.away.leagueRecord.losses
    $awayPlayersBatS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/stats?stats=season&group=hitting&teamId=$awayTeamId").stats.splits
    $awayRS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=runs&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $awayTeam}
    $awayAvgS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=battingAverage&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $awayTeam}
    $awayHrS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=homeRuns&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $awayTeam}
    $awaySlgS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=slg&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $awayTeam}
    $awayOpS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=ops&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $awayTeam}
    $awayTmBatRec = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/stats?stats=lastXGames&group=hitting&teamId=$awayTeamId").stats.splits | Sort-Object -Descending {$_.stat.hits}
    $awayTmTopBats = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/$awayTeamId/leaders?leaderCategories=hits&limit=8").teamLeaders[0].leaders.person
    $awayTmRecNm = $awayTmBatRec.player.fullname
    $awayTmRecSt = $awayTmBatRec.stat
    $awayTeamRecIdsAll = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?stats=byDateRange&sportId=1&teamId=$awayTeamId&startDate=$formatDateFiveDays&endDate=$secondDate").dates.games.gamePk
    $awayTeamRecAwayR = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?stats=byDateRange&sportId=1&teamId=$awayTeamId&startDate=$formatDateFiveDays&endDate=$secondDate").dates.games.teams.away.score
    $awayTeamRecHomeR = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?stats=byDateRange&sportId=1&teamId=$awayTeamId&startDate=$formatDateFiveDays&endDate=$secondDate").dates.games.teams.home.score
    $awayTmHitters = @()
    $awayTmBatNames = @()
    $arrayawayTeamHome = @()
    $arrayawayTeamAway = @()
    $awayTeamRecWP = @()
    $awayTeamRecL = @()
    $awayTeamRecBatStats = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/$awayTeamId/stats?stats=byDateRange&group=hitting&startDate=$formatDateTenDays&endDate=$secondDate").stats.splits.stat
    $awayTeamVsLHP = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&season=2025&sportIds=1&stats=statSplits&sitCodes=vl").stats.splits | Where-Object { $_.team.id -match $awayTeamId}
    $awayTeamVsRHP = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&season=2025&sportIds=1&stats=statSplits&sitCodes=vr").stats.splits | Where-Object { $_.team.id -match $awayTeamId}
    $awayTeamAheadCnt = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&season=2025&sportIds=1&stats=statSplits&sitCodes=ac").stats.splits | Where-Object { $_.team.id -match $awayTeamId}
    $awayTeamBehindCnt = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&season=2025&sportIds=1&stats=statSplits&sitCodes=bc").stats.splits | Where-Object { $_.team.id -match $awayTeamId}

    foreach ($awayTmBat in $awayTmTopBats) {
        $awayTmBatId = $awayTmBat.id
        $awayTmBatName = $awayTmBat.fullname
        $awayTmBatStat = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$awayTmBatId/stats?stats=season&group=hitting").stats.splits.stat
        $awayTmHitters += $awayTmBatStat  
        $awayTmBatNames += $awayTmBatName
    }
    foreach ($awayTeamRecPk in $awayTeamRecIdsAll) {
        $awayTeamRecWinners = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$awayTeamRecPk/feed/live").livedata.decisions.winner.fullname
        $awayTeamRecLosers = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$awayTeamRecPk/feed/live").livedata.decisions.loser.fullname
        $awayTeamRecHomeAbv = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$awayTeamRecPk/feed/live").gamedata.teams.home.abbreviation
        $awayTeamRecAwayAbv = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$awayTeamRecPk/feed/live").gamedata.teams.away.abbreviation
        $awayTeamRecWP += "W- " + $awayTeamRecWinners 
        $awayTeamRecL += "  L- " + $awayTeamRecLosers
        $arrayawayTeamHome += $awayTeamRecHomeAbv  + ' ' 
        $arrayawayTeamAway += $awayTeamRecAwayAbv + ' ' 
    }

    $awayTeamRecScore1 = '.' + $arrayawayTeamAway[4] + $awayTeamRecAwayR[4] + ' @ ' + $arrayawayTeamHome[4] + $awayTeamRecHomeR[4]
    $awayTeamRecScore2 = '..' + $arrayawayTeamAway[3] + $awayTeamRecAwayR[3] + ' @ ' + $arrayawayTeamHome[3] + $awayTeamRecHomeR[3]
    $awayTeamRecScore3 = '...' + $arrayawayTeamAway[2] + $awayTeamRecAwayR[2] + ' @ ' + $arrayawayTeamHome[2] + $awayTeamRecHomeR[2]
    $awayTeamRecScore4 = $arrayawayTeamAway[1] + $awayTeamRecAwayR[1] + ' @ ' + $arrayawayTeamHome[1] + $awayTeamRecHomeR[1]

    $homeTeamStat = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/$homeTeamId/stats?stats=season&group=hitting").stats.splits.stat
    $homeTeamRecents = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?stats=byDateRange&sportId=1&teamId=$homeTeamId&startDate=$formatDateFiveDays&endDate=$secondDate").dates.games
    $homeTeamRecord = [string]$game.teams.home.leagueRecord.wins + '-' + $game.teams.home.leagueRecord.losses
    $homePlayersBatS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/stats?stats=season&group=hitting&teamId=$homeTeamId").stats.splits
    $homeRS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=runs&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $homeTeam}
    $homeAvgS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=battingAverage&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $homeTeam}
    $homeHrS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=homeRuns&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $homeTeam}
    $homeSlgS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=slg&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $homeTeam}
    $homeOpS = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats/leaders?leaderCategories=ops&statGroup=hitting&limit=33").leagueLeaders.leaders | Where-Object {$_.team.name -eq $homeTeam}
    $homeTmBatRec = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/stats?stats=lastXGames&group=hitting&teamId=$homeTeamId").stats.splits | Sort-Object -Descending {$_.stat.hits}
    $homeTmTopBats = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/$homeTeamId/leaders?leaderCategories=hits&limit=8").teamLeaders[0].leaders.person
    $homeTmRecNm = $homeTmBatRec.player.fullname
    $homeTmRecSt = $homeTmBatRec.stat
    $homeTeamRecIdsAll = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?stats=byDateRange&sportId=1&teamId=$homeTeamId&startDate=$formatDateFiveDays&endDate=$secondDate").dates.games.gamePk
    $homeTeamRecAwayR = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?stats=byDateRange&sportId=1&teamId=$homeTeamId&startDate=$formatDateFiveDays&endDate=$secondDate").dates.games.teams.away.score
    $homeTeamRecHomeR = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/schedule?stats=byDateRange&sportId=1&teamId=$homeTeamId&startDate=$formatDateFiveDays&endDate=$secondDate").dates.games.teams.home.score
    $homeTeamRecIdsForm = $homeTeamRecIdsAll -join ','
    $homeTmHitters = @()
    $homeTmBatNames = @()
    $arrayHomeTeamHome = @()
    $arrayHomeTeamAway = @()
    $homeTeamRecWP = @()
    $homeTeamRecL = @()
    $homeTeamRecBatStats = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/$homeTeamId/stats?stats=byDateRange&group=hitting&startDate=$formatDateTenDays&endDate=$secondDate").stats.splits.stat
    $homeTeamVsLHP = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&season=2025&sportIds=1&stats=statSplits&sitCodes=vl").stats.splits | Where-Object { $_.team.id -match $homeTeamId}
    $homeTeamVsRHP = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&season=2025&sportIds=1&stats=statSplits&sitCodes=vr").stats.splits | Where-Object { $_.team.id -match $homeTeamId}
    $homeTeamAheadCnt = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&season=2025&sportIds=1&stats=statSplits&sitCodes=ac").stats.splits | Where-Object { $_.team.id -match $homeTeamId}
    $homeTeamBehindCnt = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/teams/stats?group=hitting&season=2025&sportIds=1&stats=statSplits&sitCodes=bc").stats.splits | Where-Object { $_.team.id -match $homeTeamId}

    foreach ($homeTmBat in $homeTmTopBats) {
        $homeTmBatId = $homeTmBat.id
        $homeTmBatName = $homeTmBat.fullname
        $homeTmBatStat = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1/people/$homeTmBatId/stats?stats=season&group=hitting").stats.splits.stat
        $homeTmHitters += $homeTmBatStat  
        $homeTmBatNames += $homeTmBatName
    }
    foreach ($homeTeamRecPk in $homeTeamRecIdsAll) {
        $homeTeamRecWinners = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$homeTeamRecPk/feed/live").livedata.decisions.winner.fullname 
        $homeTeamRecLosers = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$homeTeamRecPk/feed/live").livedata.decisions.loser.fullname
        $homeTeamRecHomeAbv = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$homeTeamRecPk/feed/live").gamedata.teams.home.abbreviation
        $homeTeamRecAwayAbv = (Invoke-RestMethod -Uri "https://statsapi.mlb.com/api/v1.1/game/$homeTeamRecPk/feed/live").gamedata.teams.away.abbreviation
        $homeTeamRecWP += "W- " + $homeTeamRecWinners
        $homeTeamRecL += "  L- " + $homeTeamRecLosers #.Split()[0].substring(0,1) +'.' + $homeTeamRecLosers.Split()[-1] 
        $arrayHomeTeamHome += $homeTeamRecHomeAbv  + ' '
        $arrayHomeTeamAway += $homeTeamRecAwayAbv + ' '
    }

    $homeTeamRecScore1 = '.' + $arrayhomeTeamAway[4] + $homeTeamRecAwayR[4] + ' @ ' + $arrayhomeTeamHome[4] + $homeTeamRecHomeR[4]
    $homeTeamRecScore2 = '..' + $arrayhomeTeamAway[3] + $homeTeamRecAwayR[3] + ' @ ' + $arrayhomeTeamHome[3] + $homeTeamRecHomeR[3]
    $homeTeamRecScore3 = '...' + $arrayhomeTeamAway[2] + $homeTeamRecAwayR[2] + ' @ ' + $arrayhomeTeamHome[2] + $homeTeamRecHomeR[2]
    $homeTeamRecScore4 = $arrayhomeTeamAway[1] + $homeTeamRecAwayR[1] + ' @ ' + $arrayhomeTeamHome[1] + $homeTeamRecHomeR[1]
    
    $matchupStatss = [PSCustomObject]@{
        'Start' = "{0:h:mm}" -f [DateTime]$game.gameDate #+ ' CST'
        'Matchup' = $awayTeam + ' @ ' + $homeTeam 
        'Probables' = $gameInfo.probablePitchers.away.fullName + ' vs ' + "$homeTeamSP"
        'Weather' = $weather.condition + ', ' + $weather.wind
    }
    
    if ($null -ne $awaySpStats) {
        $awayPitcher = [PSCustomObject] @{
            ($awayTeamSP + ' ' + $awaySpHand) = $awaySpMeta 
            'WL  | ERA | IP' = [string]$awaySpStats.wins + '-' + [string]$awaySpStats.losses + '  | ' + $awaySpStats.era + ' | ' + $awaySpStats.inningsPitched
            'AVG | SLG' = $awaySpStats.avg + ' | ' + $awaySpAdv.slg
            'K   | BB' = $awaySpAdv.strikeoutsPer9 + ' | ' + $awaySpAdv.baseOnBallsPer9
            'BARREL | WIFF' = $awaySpBar.ToString(".000") + ' | ' + ($awaySpAdv.swingAndMisses / $awaySpAdv.totalSwings).ToString(".000")
            'STRIKE 1 | MEATBALL' = $awaySpFP + ' | ' + $awaySpMeatballPerc
            'LAST 4' = ''
            $as1 = $awaySpGameLog[-1].stat.summary + ', ' + $awaySpGameLog[-1].stat.hits + ' H'
            $as2 = $awaySpGameLog[-2].stat.summary + ', ' + $awaySpGameLog[-2].stat.hits + ' H'
            $as3 = $awaySpGameLog[-3].stat.summary + ', ' + $awaySpGameLog[-3].stat.hits + ' H'
            $as4 = $awaySpGameLog[-4].stat.summary + ', ' + $awaySpGameLog[-4].stat.hits + ' H' 
            'VS OPPONENT' = ''
            $aSPvs = $aSpvsD
            $aSPvs2 = $aSpvsD2
            'CAREER' = $awaySpCaStats.inningsPitched + ' IP, ' + $awaySpCaStats.earnedRuns + ' ER, ' + $awaySpCaStats.strikeOuts + ' K, ' +  $awaySpCaStats.baseOnBalls + ' BB, ' +
                        $awaySpCaStats.hits + ' H' 
            '1ST INNING' = $awaySpFip.inningsPitched + ' IP, ' + $awaySpFip.earnedRuns + ' ER, ' + $awaySpFip.strikeOuts + ' K, ' + $awaySpFip.baseOnBalls + ' BB, ' + $awaySpFip.hits + ' H' 
            'RUNNER ON' = 'AVG ' + $awaySpRon.avg + ', OBP ' + $awaySpRon.obp + ', SLG ' + $awaySpRon.slg + ', BB% ' + $awaySpRon.walksPer9Inn
            $awaySpArs1 = $awaySpArs[0].percentage.ToString("00%") + ', ' + ([int]$awaySpArs[0].averageSpeed).ToString() + ' MPH, ' +  ([int]$awaySpSpin1.metric[0].averageValue).ToString() + ' SPIN'  
            $awaySpArs2 = $awaySpArs[1].percentage.ToString("00%") + ', ' + ([int]$awaySpArs[1].averageSpeed).ToString() + ' MPH, ' +  ([int]$awaySpSpin2.metric[0].averageValue).ToString() + ' SPIN'  
            $awaySpArs3 = $awaySpArs[2].percentage.ToString("00%") + ', ' + ([int]$awaySpArs[2].averageSpeed).ToString() + ' MPH, ' +  ([int]$awaySpSpin3.metric[0].averageValue).ToString() + ' SPIN'
        }
    } else {
        $awayPitcher = "No Stats for Away SP " + $gameinfo.probablePitchers.away.fullName
        $awayTeamStatss.PSObject.Properties.Remove($awayTmvsSpNm1)
        $awayTeamStatss.PSObject.Properties.Remove($awayTmvsSpNm2)
    }

    $awayTeamStatss = [PSCustomObject] @{
        ($awayTeam).ToUpper() = $awayTeamRecord
        "R" = '#' + "{0:d2}" -f $awayRS.rank + ' ' + "{0:d2}" -f $awayRS.value
        "OPS" = '#' + "{0:d2}" -f $awayOpS.rank + ' ' + $awayOps.value
        $awayTmBatNames[0] = 'AVG ' + $awayTmHitters[0].avg + ', HR ' + "{0:d2}" -f $awayTmHitters[0].homeRuns + 
                                       ', OBP ' + $awayTmHitters[0].obp + ', SLG ' + $awayTmHitters[0].slg 
        $awayTmBatNames[1] = 'AVG ' + $awayTmHitters[1].avg + ', HR ' + "{0:d2}" -f $awayTmHitters[1].homeRuns + 
                                       ', OBP ' + $awayTmHitters[1].obp + ', SLG ' + $awayTmHitters[1].slg 
        $awayTmBatNames[2] = 'AVG ' + $awayTmHitters[2].avg + ', HR ' + "{0:d2}" -f $awayTmHitters[2].homeRuns + 
                                       ', OBP ' + $awayTmHitters[2].obp + ', SLG ' + $awayTmHitters[2].slg 
        $awayTmBatNames[3] = 'AVG ' + $awayTmHitters[3].avg + ', HR ' + "{0:d2}" -f $awayTmHitters[3].homeRuns + 
                                       ', OBP ' + $awayTmHitters[3].obp + ', SLG ' + $awayTmHitters[3].slg 
        $awayTmBatNames[4] = 'AVG ' + $awayTmHitters[4].avg + ', HR ' + "{0:d2}" -f $awayTmHitters[4].homeRuns + 
                                       ', OBP ' + $awayTmHitters[4].obp + ', SLG ' + $awayTmHitters[4].slg 
        $awayTmBatNames[5] = 'AVG ' + $awayTmHitters[5].avg + ', HR ' + "{0:d2}" -f $awayTmHitters[5].homeRuns + 
                                       ', OBP ' + $awayTmHitters[5].obp + ', SLG ' + $awayTmHitters[5].slg 
        $awayTmBatNames[6] = 'AVG ' + $awayTmHitters[6].avg + ', HR ' + "{0:d2}" -f $awayTmHitters[6].homeRuns + 
                                       ', OBP ' + $awayTmHitters[6].obp + ', SLG ' + $awayTmHitters[6].slg + " `n"
        #$awayTmBatNames[7] = 'AVG ' + $awayTmHitters[7].avg + ', HR ' + "{0:d2}" -f $awayTmHitters[7].homeRuns + 
        #                               ', OBP ' + $awayTmHitters[7].obp + ', SLG ' + $awayTmHitters[7].slg #+ " `n"
        ".LAST 10" = 'AVG ' + $awayTeamRecBatStats.avg + ', HR ' + "{0:d2}" -f $awayTeamRecBatStats.homeRuns + ', OPS ' + $awayTeamRecBatStats.ops
        ('.' + $awayTmRecNm[0]) = 'AVG ' + $awayTmRecSt[0].avg + ', HR ' + "{0:d2}" -f $awayTmRecSt[0].homeRuns + 
                                            ', OBP ' + $awayTmRecSt[0].obp + ', SLG ' + $awayTmRecSt[0].slg #+ ', K ' + "{0:d2}" -f $awayTmRecSt[0].strikeOuts
        ('.' + $awayTmRecNm[1]) = 'AVG ' + $awayTmRecSt[1].avg + ', HR ' + "{0:d2}" -f $awayTmRecSt[1].homeRuns + 
                                            ', OBP ' + $awayTmRecSt[1].obp + ', SLG ' + $awayTmRecSt[1].slg #+ ', K ' + "{0:d2}" -f $awayTmRecSt[1].strikeOuts
        ('.' + $awayTmRecNm[2]) = 'AVG ' + $awayTmRecSt[2].avg + ', HR ' + "{0:d2}" -f $awayTmRecSt[2].homeRuns + 
                                            ', OBP ' + $awayTmRecSt[2].obp + ', SLG ' + $awayTmRecSt[2].slg #+ ', K ' + "{0:d2}" -f $awayTmRecSt[2].strikeOuts
        ('.' + $awayTmRecNm[3]) = 'AVG ' + $awayTmRecSt[3].avg + ', HR ' + "{0:d2}" -f $awayTmRecSt[3].homeRuns + 
                                            ', OBP ' + $awayTmRecSt[3].obp + ', SLG ' + $awayTmRecSt[3].slg #+ ', K ' + "{0:d2}" -f $awayTmRecSt[3].strikeOuts
        ('.' + $awayTmRecNm[4]) = 'AVG ' + $awayTmRecSt[4].avg + ', HR ' + "{0:d2}" -f $awayTmRecSt[4].homeRuns + 
                                            ', OBP ' + $awayTmRecSt[4].obp + ', SLG ' + $awayTmRecSt[4].slg #+ ', K ' + "{0:d2}" -f $awayTmRecSt[4].strikeOuts
        ('.' + $awayTmRecNm[5]) = 'AVG ' + $awayTmRecSt[5].avg + ', HR ' + "{0:d2}" -f $awayTmRecSt[5].homeRuns + 
                                            ', OBP ' + $awayTmRecSt[5].obp + ', SLG ' + $awayTmRecSt[5].slg #+ ', K ' + "{0:d2}" -f $awayTmRecSt[5].strikeOuts
        'VS SP' = ''
        #$awayTmvsSpNm1 = $awayTmvsSpSt1
        'RECENT SCORES' = ''
        $awayTeamRecScore1 = $awayTeamRecWP[4] + $awayTeamRecL[4]
        $awayTeamRecScore2 = $awayTeamRecWP[3] + $awayTeamRecL[3]
        #$awayTeamRecScore3 = $awayTeamRecWP[2] + $awayTeamRecL[2] 
        #$awayTeamRecScore4 = $awayTeamRecWP[1] + $awayTeamRecL[1] 
                                 
        'vs RHP' = 'AVG ' + $awayTeamVsRHP.stat.avg + ', OBP ' + $awayTeamVsRHP.stat.obp + ', OPS ' + $awayTeamVsRHP.stat.ops
        'vs LHP' = 'AVG ' + $awayTeamVsLHP.stat.avg + ', OBP ' + $awayTeamVsLHP.stat.obp + ', OPS ' + $awayTeamVsLHP.stat.ops
        'Ahead /Count' = 'AVG ' + $awayTeamAheadCnt.stat.avg + ', OBP ' + $awayTeamAheadCnt.stat.obp + ', OPS ' + $awayTeamAheadCnt.stat.ops
        'Behind /Count' = 'AVG ' + $awayTeamBehindCnt.stat.avg + ', OBP ' + $awayTeamBehindCnt.stat.obp + ', OPS ' + $awayTeamBehindCnt.stat.ops
    }

    if ($null -ne $homeSpStats) {
        $homePitcher = [PSCustomObject] @{
            ($homeTeamSp + ' ' + $homeSpHand) = $homeSpMeta
            'WL  | ERA | IP' = [string]$homeSpStats.wins + '-' + [string]$homeSpStats.losses + '  | ' + $homeSpStats.era + ' | ' + $homeSpStats.inningsPitched 
            'AVG | SLG' = $homeSpStats.avg + ' | ' + $homeSpStats.slg
            'K   | BB' = $homeSpAdv.strikeoutsPer9 + ' | ' + $homeSpAdv.baseOnBallsPer9
            'BARREL | WIFF' = $homeSpBar.ToString(".000") + ' | ' + ($homeSpAdv.swingAndMisses / $homeSpAdv.totalSwings).ToString(".000")
            'STRIKE 1 | MEATBALL' = $homeSpFP + ' | ' + $homeSpMeatballPerc
            'LAST 4' = ' '
            $hs1 = $homeSpGameLog[-1].stat.summary + ', ' + $homeSpGameLog[-1].stat.hits + ' H'
            $hs2 = $homeSpGameLog[-2].stat.summary + ', ' + $homeSpGameLog[-2].stat.hits + ' H'
            $hs3 = $homeSpGameLog[-3].stat.summary + ', ' + $homeSpGameLog[-3].stat.hits + ' H' 
            $hs4 = $homeSpGameLog[-4].stat.summary + ', ' + $homeSpGameLog[-4].stat.hits + ' H' 
            'VS OPPONENT' = ''
            $hSPvs = $hSpvsD
            $hSPvs2 = $hSpvsD2
            'CAREER' = $homeSpCaStats.inningsPitched + ' IP, ' + $homeSpCaStats.earnedRuns + ' ER, ' + $homeSpCaStats.strikeOuts + ' K, ' +  $homeSpCaStats.baseOnBalls + ' BB, ' +
                       $homeSpCaStats.hits + ' H'
            '1ST INNING' = $homeSpFip.inningsPitched + ' IP, ' + $homeSpFip.earnedRuns + ' ER, ' + $homeSpFip.strikeOuts + ' K, ' + $homeSpFip.baseOnBalls + ' BB, ' + $homeSpFip.hits + ' H' 
            'RUNNER ON' = 'AVG ' + $homeSpRon.avg + ', OBP ' + $homeSpRon.obp + ', SLG ' + $homeSpRon.slg + ', BB% ' + $homeSpRon.walksPer9Inn
            $homeSpArs1 = $homeSpArs[0].percentage.ToString("00%") + ', ' + ([int]$homeSpArs[0].averageSpeed).ToString() + ' MPH, ' +  ([int]$homeSpSpin1.metric[0].averageValue).ToString() + ' SPIN'  
            $homeSpArs2 = $homeSpArs[1].percentage.ToString("00%") + ', ' + ([int]$homeSpArs[1].averageSpeed).ToString() + ' MPH, ' +  ([int]$homeSpSpin2.metric[0].averageValue).ToString() + ' SPIN'  
            $homeSpArs3 = $homeSpArs[2].percentage.ToString("00%") + ', ' + ([int]$homeSpArs[2].averageSpeed).ToString() + ' MPH, ' +  ([int]$homeSpSpin3.metric[0].averageValue).ToString() + ' SPIN'
        }
    } else {$homePitcher = "No stats this season for Home SP"
    }

    $homeTeamStatss = [PSCustomObject] @{
        ($homeTeam).ToUpper() = $homeTeamRecord
        "R" = '#' + "{0:d2}" -f $homeRS.rank + ' ' + $homeRS.value
        "OPS" = '#' + "{0:d2}" -f $homeOpS.rank + ' ' + $homeOps.value        
        $homeTmBatNames[0] = 'AVG ' + $homeTmHitters[0].avg + ', HR ' + "{0:d2}" -f $homeTmHitters[0].homeRuns + 
                                       ', OBP ' + $homeTmHitters[0].obp + ', SLG ' + $homeTmHitters[0].slg 
        $homeTmBatNames[1] = 'AVG ' + $homeTmHitters[1].avg + ', HR ' + "{0:d2}" -f $homeTmHitters[1].homeRuns + 
                                       ', OBP ' + $homeTmHitters[1].obp + ', SLG ' + $homeTmHitters[1].slg 
        $homeTmBatNames[2] = 'AVG ' + $homeTmHitters[2].avg + ', HR ' + "{0:d2}" -f $homeTmHitters[2].homeRuns + 
                                       ', OBP ' + $homeTmHitters[2].obp + ', SLG ' + $homeTmHitters[2].slg 
        $homeTmBatNames[3] = 'AVG ' + $homeTmHitters[3].avg + ', HR ' + "{0:d2}" -f $homeTmHitters[3].homeRuns + 
                                       ', OBP ' + $homeTmHitters[3].obp + ', SLG ' + $homeTmHitters[3].slg 
        $homeTmBatNames[4] = 'AVG ' + $homeTmHitters[4].avg + ', HR ' + "{0:d2}" -f $homeTmHitters[4].homeRuns + 
                                       ', OBP ' + $homeTmHitters[4].obp + ', SLG ' + $homeTmHitters[4].slg 
        $homeTmBatNames[5] = 'AVG ' + $homeTmHitters[5].avg + ', HR ' + "{0:d2}" -f $homeTmHitters[5].homeRuns + 
                                       ', OBP ' + $homeTmHitters[5].obp + ', SLG ' + $homeTmHitters[5].slg 
        $homeTmBatNames[6] = 'AVG ' + $homeTmHitters[6].avg + ', HR ' + "{0:d2}" -f $homeTmHitters[6].homeRuns + 
                                       ', OBP ' + $homeTmHitters[6].obp + ', SLG ' + $homeTmHitters[6].slg + " `n"
        ".LAST 10" = 'AVG ' + $homeTeamRecBatStats.avg + ', HR ' + "{0:d2}" -f $homeTeamRecBatStats.homeRuns + ', OPS ' + $homeTeamRecBatStats.ops
        ('.' + $homeTmRecNm[0]) = 'AVG ' + $homeTmRecSt[0].avg + ', HR ' + "{0:d2}" -f $homeTmRecSt[0].homeRuns + 
                                  ', OBP ' + $homeTmRecSt[0].obp + ', SLG ' + $homeTmRecSt[0].slg 
        ('.' + $homeTmRecNm[1]) = 'AVG ' + $homeTmRecSt[1].avg + ', HR ' + "{0:d2}" -f $homeTmRecSt[1].homeRuns + 
                                  ', OBP ' + $homeTmRecSt[1].obp + ', SLG ' + $homeTmRecSt[1].slg 
        ('.' + $homeTmRecNm[2]) = 'AVG ' + $homeTmRecSt[2].avg + ', HR ' + "{0:d2}" -f $homeTmRecSt[2].homeRuns + 
                                  ', OBP ' + $homeTmRecSt[2].obp + ', SLG ' + $homeTmRecSt[2].slg 
        ('.' + $homeTmRecNm[3]) = 'AVG ' + $homeTmRecSt[3].avg + ', HR ' + "{0:d2}" -f $homeTmRecSt[3].homeRuns + 
                                  ', OBP ' + $homeTmRecSt[3].obp + ', SLG ' + $homeTmRecSt[3].slg 
        ('.' + $homeTmRecNm[4]) = 'AVG ' + $homeTmRecSt[4].avg + ', HR ' + "{0:d2}" -f $homeTmRecSt[4].homeRuns + 
                                  ', OBP ' + $homeTmRecSt[4].obp + ', SLG ' + $homeTmRecSt[4].slg 
        ('.' + $homeTmRecNm[5]) = 'AVG ' + $homeTmRecSt[5].avg + ', HR ' + "{0:d2}" -f $homeTmRecSt[5].homeRuns + 
                                  ', OBP ' + $homeTmRecSt[5].obp + ', SLG ' + $homeTmRecSt[5].slg 
        'VS SP' = ''
        #$homeTmvsSpNm1 = $homeTmvsSpSt1
        'RECENT SCORES' = ''
         $homeTeamRecScore1 = $homeTeamRecWP[4] + $homeTeamRecL[4]
         $homeTeamRecScore2 = $homeTeamRecWP[3] + $homeTeamRecL[3]
         #$homeTeamRecScore3 = $homeTeamRecWP[2] + $homeTeamRecL[2]
         #$homeTeamRecScore4 = $homeTeamRecWP[1] + $homeTeamRecL[1]
        'vs RHP' = 'AVG ' + $homeTeamVsRHP.stat.avg + ', OBP ' + $homeTeamVsRHP.stat.obp + ', OPS ' + $homeTeamVsRHP.stat.ops
        'vs LHP' = 'AVG ' + $homeTeamVsLHP.stat.avg + ', OBP ' + $homeTeamVsLHP.stat.obp + ', OPS ' + $homeTeamVsLHP.stat.ops
        'Ahead  /Count' = 'AVG ' + $homeTeamAheadCnt.stat.avg + ', OBP ' + $homeTeamAheadCnt.stat.obp + ', OPS ' + $homeTeamAheadCnt.stat.ops
        'Behind /Count' = 'AVG ' + $homeTeamBehindCnt.stat.avg + ', OBP ' + $homeTeamBehindCnt.stat.obp + ', OPS ' + $homeTeamBehindCnt.stat.ops
    }
        
    if ($null -eq $awaySpGameLog[-1]) {
        $awayPitcher.psobject.Properties.Remove($as1)
    }
    if ($null -eq $awaySpGameLog[-2]) {
        $awayPitcher.psobject.Properties.Remove($as2)
    }
    if ($null -eq $awaySpGameLog[-3]) {
        $awayPitcher.psobject.Properties.Remove($as3)
    }
    if ($null -eq $awaySpGameLog[-4]) {
        $awayPitcher.psobject.Properties.Remove($as4)
    }
    if ($null -eq $homeSpGameLog[-1]) {
        $homePitcher.psobject.Properties.Remove($hs1)
    }
    if ($null -eq $homeSpGameLog[-2]) {
        $homePitcher.psobject.Properties.Remove($hs2)
    }
    if ($null -eq $homeSpGameLog[-3]) {
        $homePitcher.psobject.Properties.Remove($hs3)
    }
    if ($null -eq $homeSpGameLog[-4]) {
        $homePitcher.psobject.Properties.Remove($hs4)
    }

    if ($null -eq $awaySpGameLog) {
        $awayPitcher.psobject.Properties.Remove($as1)
        $awayPitcher.psobject.Properties.Remove($as2)
        $awayPitcher.psobject.Properties.Remove($as3)
        $awayPitcher.psobject.Properties.Remove($as4)
    }
    if ($null -eq $homeSpGameLog) {
        $homePitcher.psobject.Properties.Remove($hs1)
        $homePitcher.psobject.Properties.Remove($hs2)
        $homePitcher.psobject.Properties.Remove($hs3)
        $homePitcher.psobject.Properties.Remove($hs4)
    }

    if ($homeSpvsAwayTm.count -eq 0) {
        $homePitcher.psobject.Properties.Remove('VS OPPONENT')
    }

    if ($awaySpvsHomeTm.count -eq 0) {
        $awayPitcher.psobject.Properties.Remove('VS OPPONENT')
    }
    
    if ($null -eq $awaySpvsHomeTm) {
        $awayPitcher.psobject.Properties.Remove($aSpvs2)
    }
    if ($null -eq $awaySpvsHomeTm.Count -or $awaySpvsHomeTm.Count -lt 1) {
        $awayPitcher.psobject.Properties.Remove('.')
    }

    if ($null -eq $homeSpvsAwayTm) {
        $homePitcher.psobject.Properties.Remove($hSpvs2)
    }
    if ($null -eq $homeSpvsAwayTm.Count -or $homeSpvsAwayTm.Count -lt 1) {
        $homePitcher.psobject.Properties.Remove('.')
    }

    if ($awayTmvsSp.Count -eq 0 -or [string]::IsNullOrEmpty($awayTmvsSp)) {
        $awayTeamStatss.psobject.Properties.Remove($awayTmvsSpNm1)
        $awayTeamStatss.psobject.Properties.Remove($awayTmvsSpNm2)
        $awayTeamStatss.psobject.Properties.Remove('na1')
        $awayTeamStatss.psobject.Properties.Remove('na2')
    }

    if ($homeTmvsSp.Count -eq 0 -or [string]::IsNullOrEmpty($homeTmvsSp)) {
        $homeTeamStatss.psobject.Properties.Remove($homeTmvsSpNm1)
        $homeTeamStatss.psobject.Properties.Remove($homeTmvsSpNm2)
        $homeTeamStatss.psobject.Properties.Remove('na2')
        $homeTeamStatss.psobject.Properties.Remove('na1')
    }
    if ($homeTmvsSp.Count -eq 1) {
        $homeTeamStatss.psobject.Properties.Remove($homeTmvsSpNm2)
    }

    if ($awayTmvsSp.Count -eq 1) {
        $awayTeamStatss.psobject.Properties.Remove($awayTmvsSpNm2)
    }

    if ([string]::IsNullOrEmpty($gameInfo.probablePitchers.away.fullName)) {
        $awayPitcher = 'Away Sp Undecided'
    }
    if ([string]::IsNullOrEmpty($gameInfo.probablePitchers.home.fullName)) {
        $homePitcher = 'Home Sp Undecided'
    }

    $matchupStatss | Format-List
    $awayPitcher | Format-List 
    $homePitcher | Format-List 
    $awayTeamStatss | Format-List 
    $homeTeamStatss | Format-List 

    Write-Host "Next Matchup..." -ForegroundColor Green

    $originalFilePath = ".\gametemplate.html"
    $newFilePath = ".\games\$Matchup.html"
    $fileContent = Get-Content -Path $originalFilePath -Raw
    $replacements = @{
        '$table1' = $awayPitcher | ConvertTo-Html -As List -Fragment
        '$table2' = $homePitcher | ConvertTo-Html -As List -Fragment
        '$table3' = $awayTeamStatss | ConvertTo-Html -As List -Fragment
        '$table4' = $homeTeamStatss | ConvertTo-Html -As List -Fragment
        '$header1' = $awayTeamSp 
        '$header2' = $homeTeamSp 
        '$header3' = $awayTeam 
        '$header4' = $homeTeam  
        '$title1' = $matchupStatss.Matchup
        '$title2' = $matchupStatss.Probables
        '$title3' = $matchupStatss.Weather
        '$title4' = $matchupStatss.Start
    }

    foreach ($oldText in $replacements.Keys) {
        $newText = $replacements[$oldText]

        if ($fileContent -match [regex]::Escape($oldText)) {
            Write-Host "Replacing $oldText"  #'$oldText' with '$newText'"
            $fileContent = $fileContent -replace [regex]::Escape($oldText), $newText
        } else {
            Write-Host "'$oldText' not found in the file content" -ForegroundColor Yellow
        }
    }

    Set-Content -Path $newFilePath -Value $fileContent
    Write-Host "HTML content replaced and saved to new file successfully."


    if ($i -le 1) {
        $originalFilePath2 = ".\hometemplate.html"
        $newFilePath2 = ".\updatedtemplate.html"
    } else {
        $originalFilePath2 = ".\updatedtemplate.html"
        $newFilePath2 = ".\updatedtemplate.html" 
    }

    $fileContent2 = Get-Content -Path $originalFilePath2 -Raw
    $gameI = '$' + $i + 'game' 
    $LinkIgame = '$Link' + $i + 'game'
    $DateIgame = '$Date' + $i + 'game' 
    $replacements2 = @{
        $gameI = $Matchup 
        $LinkIgame = '.\games\' + $Matchup + '.html'
        $DateIgame = $secondDate
    }

    foreach ($oldText2 in $replacements2.Keys) {
        $newText2 = $replacements2[$oldText2]
        $escapedOldText2 = [regex]::Escape($oldText2)
    
        if ($fileContent2 -match $escapedOldText2) {
            Write-Host "Replacing $oldText2" 
            $fileContent2 = $fileContent2 -replace [regex]::Escape($oldText2), $newText2 
        } else {
            Write-Host "$oldText2 not found in the file content" -ForegroundColor Yellow
        }
    }

    Set-Content -Path $newFilePath2 -Value $fileContent2
    Write-Host "HTML content replaced and saved to new file successfully."

}

#Read-Host "report has finished generating, press enter to exit"



