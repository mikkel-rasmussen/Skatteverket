# Helper for acquiring the archival exchange rate for a given date.
function Get-SEKUSDPMIAtDate([string]$startDate, [string]$endDate)
{
	# https://developer.api.riksbank.se/api-details
	$groups = Invoke-RestMethod -Uri https://api.riksbank.se/swea/v1/CrossRates/SEKUSDPMI/SEKETT/$startDate/$endDate -Method Get
	return $groups
}

function Get-ExchangeRateOnDate([datetime]$date, $exchangeRates)
{
	$retryCount = 5
	$rateOnDate = $null
	$altRateOnDate = $null
	[datetime]$altDate = $date
	do 
	{
		foreach ($rate in $exchangeRates)
		{
			$exchangeDate = [datetime]::ParseExact($rate.date, "yyyy-MM-dd", $null)
			if ($exchangeDate -eq $date)
			{
				$rateOnDate = $rate.value
			}

			if ($exchangeDate -eq $altDate)
			{
				$altRateOnDate = $rate.value
			}
		}

		if (($null -eq $rateOnDate) -and ($null -eq $altRateOnDate)) { 
			if ($retryCount -eq 0) { return 0 }
			$date = $date.AddDays(-1.0)
			$altDate = $altDate.AddDays(1.0)
			$retryCount--
		}
	}
	while ($null -eq $rateOnDate -and $null -eq $altRateOnDate)

	if ($null -ne $rateOnDate -and $rateOnDate -gt $altRateOnDate) { return $rateOnDate	}
	return $altRateOnDate
}

$gainsLosses = $(get-content -raw -path gainsLosses.json | ConvertFrom-Json).data.gainsAndLosses.list.gainsLossDtlsList

$k4 = @()
$totalClosing = 0
$totalOpening = 0
$totalGain = 0
$totalLoss = 0

[datetime]$firstDate = [datetime]::MinValue.Date
[datetime]$lastDate = [datetime]::MinValue.Date

foreach ($gl in $gainsLosses)
{
	$openingDate = [datetime]::ParseExact($gl.openingTransDateAcquired, "MM/dd/yyyy", $null)
	$closingDate = [datetime]::ParseExact($gl.closingTransDateSold, "MM/dd/yyyy", $null)

	# set an initial value for the first date
	if ($firstDate -eq [datetime]::MinValue) { $firstDate = $openingDate }

	# find the earliest date
	if ($openingDate -lt $firstDate) { $firstDate = $openingDate }

	# find the latest date
	if ($closingDate -gt $lastDate) { $lastDate = $closingDate }
}

$date1 = Get-Date $firstDate -format yyyy-MM-dd
Write-Output $([string]::Format("Date of first entry (from opening date transaction acquired): {0}", $date1))
$date2 = Get-Date $lastDate -format yyyy-MM-dd
Write-Output $([string]::Format("Date of last entry (from closing dates sold): {0}", $date2))
Write-Output "---------"

# making a single api call for all the dates since the allowed api call count is limited without an api key
$exchangeRates = Get-SEKUSDPMIAtDate $date1 $date2

# Write-Output $exchangeRates

foreach ($gl in $gainsLosses)
{
	$openingDate = [datetime]::ParseExact($gl.openingTransDateAcquired, "MM/dd/yyyy", $null)
	$closingDate = [datetime]::ParseExact($gl.closingTransDateSold, "MM/dd/yyyy", $null)

	$openingRate = Get-ExchangeRateOnDate $openingDate $exchangeRates
	$closingRate = Get-ExchangeRateOnDate $closingDate $exchangeRates
	
	$openingSEK = [int][math]::Round($gl.openingTransAdjCostBasis * $openingRate)
	$closingSEK = [int][math]::Round($gl.closingTransTotalProceeds * $closingRate)
	$diffSEK = $closingSEK - $openingSEK
	
	$gain = [math]::Max(0, $diffSEK)
	$loss = [math]::Min(0, $diffSEK)
	
	$totalClosing += $closingSEK
	$totalOpening += $openingSEK
	$totalGain += $gain
	$totalLoss += $loss
	
	$k4Row = New-Object PSObject |
		Add-Member -Type NoteProperty -Name 'Antal' -Value $gl.quantity -PassThru |
		Add-Member -Type NoteProperty -Name 'Beteckning' -Value $gl.symbol -PassThru |
		Add-Member -Type NoteProperty -Name 'Försäljningspris' -Value $closingSEK -PassThru |
		Add-Member -Type NoteProperty -Name 'Omkostnadsbelopp' -Value $openingSEK -PassThru |
		Add-Member -Type NoteProperty -Name 'Vinst' -Value $gain -PassThru |
		Add-Member -Type NoteProperty -Name 'Förlust' -Value $loss -PassThru |
		Add-Member -Type NoteProperty -Name 'Öppningskurs' -Value $openingRate -PassThru |
		Add-Member -Type NoteProperty -Name 'Slutkurs' -Value $closingRate -PassThru
	$k4 += $k4Row
}

$totals = New-Object PSObject |
	Add-Member -Type NoteProperty -Name 'Summa försäljningspris' -Value $totalClosing -PassThru |
	Add-Member -Type NoteProperty -Name 'Summa omkostnadsbelopp' -Value $totalOpening -PassThru |
	Add-Member -Type NoteProperty -Name 'Summa vinst' -Value $totalGain -PassThru |
	Add-Member -Type NoteProperty -Name 'Summa förlust' -Value $totalLoss -PassThru

$k4 | Format-Table
$totals