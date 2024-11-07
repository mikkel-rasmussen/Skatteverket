# Helper for acquiring the archival exchange rate for a given date.
function Get-SEKUSDPMIAtDate([string]$startDate, [string]$endDate)
{
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

$cashTransactions = $(get-content -raw -path getCashTransactions.json | ConvertFrom-Json).data.value.cashTransactionActivities

$p72 = @()
$totalDividendUSD = 0.0
$totalTaxUSD = 0.0
$totalDividendSEK = 0
$totalTaxSEK = 0

[datetime]$firstDate = [datetime]::MinValue.Date
[datetime]$lastDate = [datetime]::MinValue.Date

foreach ($ct in $cashTransactions)
{
	$txDate = [datetime]::ParseExact($ct.transactionDate, "MM/dd/yyyy", $null)

	# set an initial value for the first date
	if ($firstDate -eq [datetime]::MinValue) { $firstDate = $txDate }

	# find the earliest date
	if ($txDate -lt $firstDate) { $firstDate = $txDate }

	# find the latest date
	if ($txDate -gt $lastDate) { $lastDate = $txDate }
}

$date1 = Get-Date $firstDate -format yyyy-MM-dd
$date2 = Get-Date $lastDate -format yyyy-MM-dd

# making a single api call for all the dates since the allowed api call count is limited without an api key
$exchangeRates = Get-SEKUSDPMIAtDate $date1 $date2

foreach ($ct in $cashTransactions)
{
	if ($ct.transactionType -eq "$+-DIV")
	{
		$txDate = [datetime]::ParseExact($ct.transactionDate, "MM/dd/yyyy", $null)
		$txRate = Get-ExchangeRateOnDate $txDate $exchangeRates
		
		$tx = $ct.amount
		$txSEK = [int][math]::Round($tx * $txRate)
		
		$dividendUSD = [math]::Max(0.0, $tx)
		$taxUSD = [math]::Min(0.0, $tx)
		$dividendSEK = [math]::Max(0, $txSEK)
		$taxSEK = [math]::Min(0, $txSEK)
		
		$totalDividendUSD += $dividendUSD
		$totalTaxUSD += $taxUSD
		$totalDividendSEK += $dividendSEK
		$totalTaxSEK += $taxSEK
	}
}

$totalDividendUSD = [math]::Round($totalDividendUSD)
$totalTaxUSD = [math]::Round($totalTaxUSD)

Write-Warning "Please note that this script only looks at the cash transaction register, which only looks 365 days back. Make sure the USD totals correspond with your 1042-S form."
$totals = New-Object PSObject |
	Add-Member -Type NoteProperty -Name 'Gross Income (Box 2) [Cross-check with 1042-S!]' -Value $totalDividendUSD -PassThru |
	Add-Member -Type NoteProperty -Name 'Federal Tax Withheld (Box 7a) [Cross-check with 1042-S!]' -Value $totalTaxUSD -PassThru |
	Add-Member -Type NoteProperty -Name 'Ränteinkomster, utdelningar m.m. 7.2' -Value $totalDividendSEK -PassThru |
	Add-Member -Type NoteProperty -Name 'Övriga upplysningar -> Avräkning av utländsk skatt' -Value $totalTaxSEK -PassThru

$totals | Format-List
