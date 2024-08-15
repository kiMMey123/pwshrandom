param(
    [Alias("sted", "plass", "l", "q")]
    [Parameter(Mandatory=$true)]
    [String] $Location,

    [Alias("Hour", "H")]
    [Parameter()]
    [switch] $Hourly
)

$script:ModuleRoot = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$script:WeatherTable = Import-Clixml "$ModuleRoot/weather-legend.xml"

function Confirm-Input {
    param( 
        [object] 
        $Setting,

        $Value 
    )
    
    $inputAccepted = $false

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @{ Value = ""; Accepted = $false; Reason = "Empty input" }
    }

    switch ( $Setting.Type ) {
        "Length" { $InputAccepted = $Value.Length -in $Setting.Range }
        "Color" { $InputAccepted = $Value -in [enum]::GetValues([System.ConsoleColor]) }
        "Font" {
            [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
            $fonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families
            if ( $Value -in $fonts ) { $InputAccepted = $true }
            else {
                $font = ( $fonts | Where-Object { $_ -like "*$Value*" } )
                if ($font -and $value.Length -ge 5) { $value = $font[0].Name; $inputAccepted = $true }
            }
        }

        "Bool" {
            $AcceptedTrue = @("Yes", "Y", "Yup", "1", "true")
            $AcceptedFalse = @( "No", "N", "Nope", "0", "false" )
    
            if ( $Value -in $AcceptedTrue ) { $value = $true; $inputAccepted = $true }
            elseif ( $Value -in $AcceptedFalse ) { $value = $false; $inputAccepted = $true } 
        }

        "Int" {
            $num = $value -as [int]
            if ( $null -ne $num) { $value = $num; $inputaccepted = ($setting.Range ? $num -in $Setting.Range : $true) } 
        }

        "Double" {
            $num = $Value -as [double]
            if ($null -ne $num) { 
                $inputAccepted = $true

                if ($Setting.Range) {
                    if ($num -eq [int]$Value) {
                        $value = $num
                    }
                    else {
                        $whole = ($Value.ToString()).Split(".")[0]
                        $dec = ($value.ToString().split(".")[-1])[0..($Setting.range - 1)] -Join ""
                        $Value = ("$whole.$dec") -as [double]
                    }
                }
                else {
                    $value = $num
                }
            }
        }

        "Location" {
            $Query = Confirm-Input -Setting @{Type = "length"; Range = (1..60) } -Value = $Value

            if ($Query.Accepted) {
                $LocationSearch = Get-LocationData -Query $Value

                if ($LocationSearch.Count -gt 1) {
                    Write-Host "Found multiple matches for $Value"
                    for ($i = 0; $i -lt 9 -and $i -lt $LocationSearch.count; $i++) {
                        $Name = $LocationSearch[$i].display_name
                        Write-Host "  $($i + 1): $Name - $($LocationSearch[$i].addresstype)"
                    }

                    $locationSelect = Read-Host "`n[1-$i] Select location | [B]ack"

                    if ($locationSelect -eq "B" ) {
                        $Reason = "Location select cancelled"
                    }

                    $LocationOK = Confirm-Input @{Type = "int"; Range = (1..$i) } -Value $locationSelect -Multi

                    if ($LocationOK.Accepted) {
                        $Location = $LocationSearch[$LocationOK.Value - 1]
                    }
                    else {
                        $Reason = "Invalid input"
                    }
                }
                else {
                    $Location = $LocationSearch[0]
                }

                $Name = $Location.display_name
                $Lat = (Confirm-Input -Setting @{Type = "double"; Range = 4 } -Value $Location.lat).Value
                $Lon = (Confirm-Input -Setting @{Type = "double"; Range = 4 } -Value $Location.lon).Value
                $value = @(
                    @{Ref = "WeatherLocation_Lat"; Name = "Latitude"; Value = $Lat }
                    @{Ref = "WeatherLocation_Long"; Name = "Longitude"; Value = $Lon }
                    @{Ref = "WeatherLocation_Name"; Name = "Location name"; Value = $Name }
                )
                $inputAccepted = $true
                
            }
            else {
                $Reason = $Query.Reason
            }

        }

        "Alias" {
            $existingAlias = $setting.ExistingAlias.PSObject.Copy()

            if ($value -eq "N") {
                $NewCmdLet = Read-Host "Enter existing cmdlet to create alias | [B]ack"

                if ( $NewCmdLet -eq "B" -or $NewCmdLet.Length -lt 2 ) { 
                    $Reason = "Exited"; Break
                }
            }
            else {
                $NewCmdlet = $setting.Name
            }

            
            $hasCmdlet = ($existingAlias | Where-Object { $NewCmdlet -in $_.Command })
            $TestCmdlet = Get-Command $NewCmdLet -ErrorAction SilentlyContinue
            
            if ($null -eq $TestCmdlet) {
                $reason = "$NewCmdlet does not exist"; Break
            }
            else {
                $NewCmdlet = $TestCmdlet.Name
            }

            if ($value -in @("A", "N") ) {
                
                $NewAlias = Read-Host "Enter new alias for $NewCmdLet | [B] to go back"
                
                if ( $Newalias -eq "B" -or $NewAlias.Length -lt 2 ) { 
                    $Reason = "not good input"; Break
                }
                
                $TestAlias = Get-Command $NewAlias -ErrorAction SilentlyContinue
                
                if ($null -ne $TestAlias -or $NewAlias -in $existingAlias.Alias ) {
                    $reason = "Cannot add $NewAlias as it already exists"; Break
                }

                if ($hasCmdlet) {
                    if ($hasCmdlet.Alias -is [string]) {
                        $hasCmdlet.Alias = @($hasCmdlet.Alias)
                    }

                    $hasCmdlet.Alias += $NewAlias
                    $UpdatedEntry = $hasCmdlet
                }
                else {
                    $UpdatedEntry = @{ Command = $NewCmdlet; Alias = @($NewAlias) }
                    $existingAlias += $UpdatedEntry
                }

                $Reason = "Added alias for $NewCmdLet -> $NewAlias"
                $inputAccepted = $true
            }

            elseif ($value -eq "D") {
                $UpdatedEntry = @{ Command = $NewCmdlet; Alias = @() }
                $existingAlias = $existingAlias | Where-Object { $_.Command -ne $NewCmdlet }
                
                $reason = "Removed all custom alias from $NewCmdLet"
                $inputAccepted = $true
            }
            
            elseif ( $null -ne $value -as [int] ) {
                $num = $value -as [int]
                if ( $null -ne $num ) {
                    if ( $num -in $setting.Range ) {
                        $removeAlias = $hasCmdlet.Alias[$num - 1]

                        if ($hasCmdlet.Alias.Count -eq 1 -or $hasCmdlet.Alias -is [string] ) {
                            $UpdatedEntry = @{ Command = $NewCmdlet; Alias = @() }
                            $existingAlias = $existingAlias | Where-Object { $_.Command -ne $hasCmdlet.Command }
                        }
                        else {
                            $hasCmdlet.Alias = $hasCmdlet.Alias | Where-Object { $_ -ne $removeAlias }
                            $UpdatedEntry = $hasCmdlet
                        }
                        $reason = "Alias '$RemoveAlias' removed from $NewCmdlet "
                        $inputAccepted = $true
                    }
                } 
            }
        }
    }

    if ( $setting.Type -eq "alias" ) {
        $Value = @{ UpdatedEntry = $UpdatedEntry; AliasUser = $existingAlias }
    }
    
    if (-not $reason) {
        if (-not $inputAccepted) {
            $reason = switch ($Setting.Type) {
                "Length" { "Input too long ($($value.Length)) - value needs to be between $($setting.Range[0]) and $($setting.Range[-1]) characters" }
                "Font" { "Could not find any font family matching $value" }
                "Bool" { "Input is not a boolean (Yes / No) answer" }
                "Color" { "Input is not a valid color" }
                "Int" { "Input is not a valid number or number is out of range $($Setting.Range ? "($($setting.Range[0]-$Setting.Range[-1]))" : $null)" }
                "Double" { "Could not convert input to valid double" }
                Default { "Input is not valid" }
            }
        }
        else {
            $reason = "$($setting.Name) updated"
        }
    }

    $toReturn = @{ Value = $Value; Accepted = $inputAccepted; Reason = $reason }
    return $toReturn
}

function Get-LocationData {
    param(
        $Query,
        $Latitude,
        $Longitude,
        [switch] $Multi
    )

    $WriteLocation = $false

    if ($Query) {
        $Query = $Query.Trim()
    }

    $StoredLocations = "$ModuleRoot\weather-locations.xml"

    if (Test-Path $StoredLocations) {
        $LocationData = Import-Clixml $StoredLocations
    } else {
        $LocationData = @(
            @{Lat = 59.9435; Lon = 10.72; Name = "Oslo-Blindern"; Query = @("blindern") }
            @{Lat = 59.9423; Lon = 10.7173; Name = "USIT"; Query = @("usit") }
        )
        $WriteLocation = $true
    }

    if ($Query) {
        $Location = $LocationData | Where-Object {$_.query -eq $Query}
        
        if (-not $Location) {
            $QueryToUrl = $Query -Replace " ", "+"
            $uri = "https://nominatim.openstreetmap.org/search?q=$QueryToUrl&format=json"
        
            try {
                $LocationQuery = Invoke-RestMethod -Method Get -Uri $uri
            }
            catch {
                Throw $_
            }
        
            if ($LocationQuery) {
                if ($Multi) {
                    return $LocationQuery
                }

                $Place = $LocationQuery[0]
                $Latitude  = (Confirm-Input -Setting @{Type = "double"; Range = 4} -Value $place.lat).Value
                $Longitude = (Confirm-Input -Setting @{Type = "double"; Range = 4} -Value $place.lon).Value
    
                $Location = $LocationData | Where-Object {$_.Lat -eq $latitude -and $_.Lon -eq $longitude}
    
                if ($Location) {
                    if ($query -notin $Location.Query) {
                        $Location.Query += $Query
                        $WriteLocation = $true
                    }
                } else {
                    $Location = @{ Name = $Place.display_name; Lat = $Latitude; Lon = $Longitude; Query = @($Query) }
                    $LocationData += $Location
                    $WriteLocation = $true
                }
            }
        }
    } elseif ($Latitude -and $Longitude) {
        $Location = $LocationData | Where-Object {$_.Lat -eq $Latitude -and $_.Lon -eq $Longitude }

        if (-not $Location) {
            $ReverseLocationUri = "https://nominatim.openstreetmap.org/reverse?lat=$Latitude&lon=$longitude"
            $LocationLookup = Invoke-RestMethod -Uri $ReverseLocationUri

            if ($LocationLookup) {
                $Location = @{ Name = $LocationLookup.reversegeocode.result."#name"; Lat = $Latitude; Lon = $Longitude; Query = @() }
                $LocationData += $Location
                $WriteLocation = $true
            }
        }
    }

    if ($WriteLocation) {
        $LocationData | Export-Clixml -Path $StoredLocations -Force
    }

    

    return $Location
}

function Get-WeatherData {
    <#
        .SYNOPSIS
        Returns weather and location data from met.no

        .DESCRIPTION
        If there is a weatherdata.xml file (CliXML formatted), imports and checks that the data was fetched in the last 5 minutes.
        If the file does not exist, data is too old or the coordinates do not match, re-fetch weather data and overwrite the xml file.
        The -Refresh switch will also fetch new weather data regardless of when it was last retrieved.
    #>
    
    param(
        [Double] $Latitude,
        [Double] $Longitude,
        [String] $Query,
        [Switch] $Refresh
    )

    $sitename = $ENV:USERDNSDOMAIN

    $Headers = @{
        "User-Agent"   = $sitename
        "Content-Type" = "application/json"
    }

    if ($Query) {
        $LocationData = Get-LocationData -Query $Query
        $Latitude = $LocationData.Lat
        $Longitude = $LocationData.Lon
    } else {
        $LocationData = Get-LocationData -Latitude $Latitude -Longitude $Longitude
    }

    if (-not $Refresh) {
        try {
            $WeatherDataFile = "$ModuleRoot\weatherdata.xml"
            $WeatherData = $WeatherDataFile | Import-Clixml -ErrorAction SilentlyContinue
            $OldWeatherData = $WeatherData -and (New-TimeSpan -Start (Get-Date($WeatherData.LastFetchTime)) -End (Get-Date)).TotalMinutes -ge 5
            $LocationMatch = ($WeatherData.Longitude -eq $Longitude) -and ($WeatherData.Latitude -eq $Latitude)
            
            $Refresh = -not $WeatherData -or $OldWeatherData -or -not $LocationMatch
            
        }
        catch {
            $Refresh = $true
        }
    }
    
    if ($Refresh) {
        $WeatherUri = "https://api.met.no/weatherapi/locationforecast/2.0/compact?lat=$Latitude&lon=$Longitude"
        $RainUri = "https://api.met.no/weatherapi/nowcast/2.0/complete?lat=$Latitude&lon=$Longitude"
        
        $WeatherData = @{
            LastFetchTime = (Get-Date)
            Latitude      = $Latitude
            Longitude     = $Longitude
            LocationData  = $LocationData
        }

        try {
            $WeatherData.Forecast = Invoke-RestMethod -Method Get -Headers $Headers -Uri $WeatherUri

            if ($WeatherData.Forecast) {
                try {
                    $Precipitation = Invoke-RestMethod -Method Get -Headers $Headers -Uri $RainUri
                }
                catch {
                    $Precipitation = $null
                }
                $WeatherData.Precipitation = $Precipitation
            }
        }
        catch {
            Throw $_
        }

        try {
            $WeatherData | Export-Clixml -Path "$ModuleRoot\weatherdata.xml" -Force
        }

        catch {
            $_
        }
    }
    return $WeatherData
}

function Get-Precipitation {
    param($WeatherResponse = (Get-WeatherData).Precipitation)

    if (-not $WeatherResponse) {
        return $null
    }

    $rainStrings = @()
    $amountStrings = @()
    
    $ItIsGoingToRain = $false
    
    foreach ($amount in $weatherResponse.properties.timeseries.data.instant.details.precipitation_rate) {
        if ($Amount -gt 0) {
            $ItIsGoingToRain = $true
        }

        if ($amount -gt 6) {
            $amount = 6
        }
        
        $thisAmount = ""
        for ($i = 0; $i -lt 6; $i++) {
            if ($amount -gt $i) {
                if (($amount - $i) -gt 0.5) {
                    $thisAmount += "█"
                }
                else {
                    $thisAmount += "▄"
                }
                
            }
            else {
                if (($i + 1) % 2 -eq 0 ) {
                    $thisAmount += "¯"
                }
                else {
                    $thisAmount += " "
                }
            }
        }
        $amountStrings += $thisAmount
    }
        
    for ($i = 5; $i -ge 0; $i--) {
        $precString = ($amountStrings | ForEach-Object { [string]$_[$i] * 2 } ) -Join ""
        $rainStrings += $precString
    }

    $startTime = $weatherResponse.properties.timeseries.time.ToLocalTime()[0]
    $endTime = $weatherResponse.properties.timeseries.time.ToLocalTime()[-1]
    $seconds = (New-TimeSpan -Start $startTime -End $endTime).TotalSeconds
    $timeString = ((@(0..5) | ForEach-Object { $startTime.AddSeconds(($seconds / 5) * $_).ToString("HH:mm") } ) -Join (" " * 4))
    
    $rainStrings += $timeString

    if ($ItIsGoingToRain) {
        return $rainStrings
    }
    else {
        return @("Ingen nedbør forventet mellom $($startTime.ToString("HH:mm")) - $($endTime.ToString("HH:mm"))")
    }
}

function New-WeatherTable {
    param($Layout, $Values, $Precipitation, $ColSpaces = 3)
    
    $TotalWidth = ($Layout.Size | Measure-Object -Sum).Sum + ($ColSpaces * ($Layout.Count - 1)) + 2

    if ($Precipitation) {
        $PrecWidth = (($Precipitation | ForEach-Object { $_.Length } ) | Measure-Object -Maximum).Maximum
        if ($PrecWidth -gt $TotalWidth) {
            $TotalWidth = $PrecWidth
        }
    }

    $separator = "─" * ($TotalWidth + 2)

    $TableRows = @()

    $Header = "  " + (($Layout | ForEach-Object { $_.Header + " " * ($_.Size - $_.Header.Length + $ColSpaces) }) -Join "")
    $TableRows += $Header
    
    foreach ($val in $Values) {
        $RowString = "  "
        
        foreach ($row in $Layout) {
            $Entry = $val.($row.Key)
            
            if ($row.Adjust -eq "I") {
                $NumSpaces = ($row.Size - $val.IconWidth)
            }
            else {
                $numSpaces = ($row.Size - $Entry.Length)
            }
            
            $col = switch -regex ($row.Adjust) {
                "R|I" { " " * $numSpaces + $entry }
                "C" { " " * ([int]($numSpaces / 2)) + $entry + " " * ($numSpaces - ([int]($numSpaces / 2))) }
                "L" { ($entry + " " * $numSpaces ) }
            }
            $col += " " * $ColSpaces
            
            $RowString += $col
        }
        if ($val.Split) {
            $TableRows += $Separator
        }
        $TableRows += $RowString + "  "
    }
    
    if ($Precipitation) {
        $TableRows += "`n"
        if ($Precipitation.Count -gt 1) {
            $Precipitation = @("Nedbør") + $Precipitation
    
        }
    
        foreach ($Line in $Precipitation) {
            $NumSpaces = $TotalWidth - $Line.Length
            $TableRows += (" " * ([int]($NumSpaces / 2))) + $Line + (" " * ([int]($NumSpaces / 2)))
        }
        $TableRows += "`n"
    }
    
    return $TableRows
}

function Get-Weather {
    [CmdletBinding()]
    [Alias("yr", "weather")]
    
    param(
        [Alias("Hour", "H")]
        [Parameter()]
        [Bool] $24Hour = $false,

        [Alias("Q")]
        [Parameter()]
        $Query,

        [Double]
        $Latitude,

        [Double]
        $Longitude
    )

    if ($Query) {
        $Validate = Confirm-Input -Setting @{Type = "Length"; Range = (1..100) } -Value $Query
        if ($Validate.Accepted) {
            try {
                $WeatherData = Get-WeatherData -Query $Query
                $Latitude = $WeatherData.Latitude
                $Longitude = $WeatherData.Longitude
            }
            catch {
                Throw "Unable to fetch weather data for $Query - $_"
            }
        }
        else {
            Throw "Unable to process input: $($validate.Reason)"
        }
    }
    else {
        if (-not $Latitude -and -not $Longitude) {
            $Latitude = $WapSettings.WeatherLocation_Lat
            $Longitude = $WapSettings.WeatherLocation_Long
            $WeatherData = Get-WeatherData -Latitude $Latitude -Longitude $Longitude 
        }
    }

    $weatherResponse = $WeatherData.Forecast
    $WeatherLocation = $WeatherData.LocationData

    if (-not $weatherResponse) {
        Throw "Missing weather data"
    }
    
    $SelectedForecast = @()
    $now = Get-Date
    
    if ($24Hour) {
        $series = $weatherResponse.properties.timeseries[($now.Minute -gt 30) ? 1..24 : 0..23]
        $key = "next_1_hours"
    }
    else {
        $series = @()
        $i = ($now.Minute -gt 30) ? 1 : 0
        $series += $weatherResponse.properties.timeseries[$i]
        $key = "next_6_hours"

        foreach ($hour in $weatherResponse.properties.timeseries[$i..($weatherResponse.properties.timeseries.count - 1)]) {
            if ($hour.time.Hour -in @(0, 6, 12, 18) -and (New-TimeSpan -Start $now -End $hour.Time).TotalDays -le 7 -and $hour.data.$key -and $i -ge 1) {
                $series += $hour
            }
            $i++
        }
    }

    $usingDate = $null
    $Prevtext = $null
    $printDay = $false
    $Widths = @()

    foreach ($hour in $series) {
        $Day = $Date = $null
        $symbol = ($hour.data.$key.summary.symbol_code).Split("_")[0]
        $legend = $WeatherTable | Where-Object { $_.Name -eq $symbol }

        $hourDate = $hour.time.ToLocalTime()

        if ($usingDate.Date -ne $hourDate.Date) {
            $printDay = $true
            $usingDate = $hour.time.ToLocalTime()
            $Date = $hourDate.ToString("dd'/'MM")
        }
        elseif ($printDay) {
            $Date = $hourDate.ToString("ddd")[0..2] -Join ''
            $printDay = $false
        }

        if ($24Hour) {
            $Time = "$($hour.Time.ToLocaltime().ToString("HH")):00 "
        }
        else {
            $nextHour = $hour.Time.AddHours(1)

            while ($nextHour.Hour -notin 0, 6, 12, 18) {
                $nextHour = $nextHour.AddHours(1)
            }

            $Time = "$($hour.Time.ToLocalTime().ToString("HH"))-$($nextHour.ToLocalTime().ToString("HH"))"
        }

        $ThisText = $legend.Text.Bokmål
        $Text = ($Prevtext -ne $ThisText -or $printDay) ? $legend.Text.Bokmål : ""

        $WindSpeed = $hour.data.instant.details.wind_speed
        $WindDirection = $hour.data.instant.details.wind_from_direction

        $WindSymbol = switch ($WindDirection) {
            { $_ -gt 22.5 -and $_ -le 67.5 } { "↗" }
            { $_ -gt 67.5 -and $_ -le 112.5 } { "→" }
            { $_ -gt 112.5 -and $_ -le 157.5 } { "↘" }
            { $_ -gt 157.5 -and $_ -le 202.5 } { "↓" }
            { $_ -gt 202.5 -and $_ -le 247.5 } { "↙" }
            { $_ -gt 247.5 -and $_ -le 292.5 } { "←" }
            { $_ -gt 292.5 -and $_ -le 337.5 } { "↖" }
            Default { "↑" }
        }
        
        $wind = "$([int]$WindSpeed) m/s $WindSymbol"
    
        $precipitation = ""

        if ($hour.data.$key.details.precipitation_amount) {
            $precString = [string]($hour.data.$key.details.precipitation_amount)
            $precipitation = $precString + " " + $weatherResponse.properties.meta.units.precipitation_amount
        }

        $Temperature = ([int]$hour.data.instant.details.air_temperature).ToString()

        $ThisForecast = @{ 
            IconWidth = $legend.IconCount * 2
            Date      = $date
            Time      = $Time
            Icon      = $legend.Icon
            Text      = $Text
            Temp      = $Temperature
            Wind      = $wind
            Prec      = $precipitation
            Split     = $printDay
        }

        $ThisForecast.Weather
        $SelectedForecast += $ThisForecast

        $Widths += @{
            Date = $Date.Length + $Day.Length
            Icon = $legend.IconCount * 2
            Text = $Text.Length
            Time = $Time.Length
            Temp = $ThisForecast.Temp.Length
            Prec = $precipitation.Length
            Wind = $Wind.Length
        }

        $Prevtext = $legend.Text.Bokmål
        $i++
    }

    $Widths += @{ Date = "Dato".Length; Icon = 4; Text = "Vær".Length; Temp = 2; Prec = "Nedbør".Length; Wind = 4 }
    
    $MaxWidths = @{}
    
    $Widths.Keys | ForEach-Object { $MaxWidths.$_ = ($Widths.$_ | Measure-Object -Maximum).Maximum }
    $MaxWidths.Total = (($MaxWidths.Keys | ForEach-Object { $MaxWidths.$_ }) | Measure-Object -Sum).Sum
    
    $TableLayout = @(
        @{Key = "date"; Size = $MaxWidths.Date; Adjust = "C"; Header = "Dato"; }
        @{Key = "time"; Size = $MaxWidths.Time; Adjust = "L"; Header = "Tid"; }
        @{Key = "icon"; Size = $MaxWidths.Icon; Adjust = "I"; Header = "  ☀️"; }
        @{Key = "text"; Size = $MaxWidths.Text; Adjust = "L"; Header = "Vær" }
        @{Key = "temp"; Size = $MaxWidths.Temp; Adjust = "R"; Header = "°C" }
        @{Key = "prec"; Size = $MaxWidths.Prec; Adjust = "R"; Header = "Nedbør" }
        @{Key = "wind"; Size = $MaxWidths.Wind; Adjust = "R"; Header = "Vind" }
    )

    if ($WeatherData.Precipitation) {
        $precipitation = Get-Precipitation -WeatherResponse $WeatherData.Precipitation
    }
    else {
        $precipitation = $null
    }
    $WeatherTable = New-WeatherTable -Layout $TableLayout -Values $SelectedForecast -Precipitation $precipitation
    $LocationName = $WeatherLocation.Name
    return $LocationName + "`n`n" + ($WeatherTable -Join "`n") + "`nOppdatert $($WeatherData.LastFetchTime)"
}


Get-Weather -Query $Location -h $Hourly
