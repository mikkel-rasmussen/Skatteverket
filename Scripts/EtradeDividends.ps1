# Helper for acquiring the archival exchange rate for a given date.
function Get-SEKUSDPMIAtDate ([datetime]$Date)
{
	$groups = $null
	do
	{
		$body = [string]::Format('<soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" xmlns:xsd="http://swea.riksbank.se/xsd">
     <soap:Header/>
     <soap:Body>
         <xsd:getInterestAndExchangeRates>
             <searchRequestParameters>
                 <aggregateMethod>D</aggregateMethod>
                 <datefrom>{0}</datefrom>
                 <dateto>{0}</dateto>
                 <languageid>en</languageid>
                 <min>false</min>
                 <avg>true</avg>
                 <max>true</max>
                 <ultimo>false</ultimo>
                 <!--1 or more repetitions:-->
                 <searchGroupSeries>
                     <groupid>11</groupid>
                     <seriesid>SEKUSDPMI</seriesid>
                 </searchGroupSeries>
             </searchRequestParameters>
         </xsd:getInterestAndExchangeRates>
     </soap:Body>
 </soap:Envelope>', $Date.ToString("yyyy-MM-dd"))
		
		$groups = $null
		if ($host.Version -gt "7.0")
		{
			$groups = $([xml]$(Invoke-WebRequest -Method Post -SkipHeaderValidation -Headers @{'Content-Type'='application/soap+xml;charset=utf-8;action=urn:getInterestAndExchangeRates'} -Body $body -Uri 'https://swea.riksbank.se/sweaWS/services/SweaWebServiceHttpSoap12Endpoint' -UseBasicParsing)."Content").Envelope.body.getInterestAndExchangeRatesResponse.return.groups
		}
		else
		{
			$groups = $([xml]$(Invoke-WebRequest -Method Post -Headers @{'Content-Type'='application/soap+xml;charset=utf-8;action=urn:getInterestAndExchangeRates'} -Body $body -Uri 'https://swea.riksbank.se/sweaWS/services/SweaWebServiceHttpSoap12Endpoint' -UseBasicParsing)."Content").Envelope.body.getInterestAndExchangeRatesResponse.return.groups
		}
		# Scan backwards in time for the latest rate, if the response for the current is empty.
		if ($null -eq $groups) { Write-Host $([string]::Format("No exchange rate for {0}, trying the previous day", $Date.ToString("yyyy-MM-dd"))) }
		$Date = $Date.AddDays(-1.0)
	}
	while ($null -eq $groups)
	return [double]$groups.series.resultrows.value.'#text'
}

$cashTransactions = $(get-content -raw -path getCashTransactions.json | ConvertFrom-Json).data.value.cashTransactionActivities

$p72 = @()
$totalDividendUSD = 0.0
$totalTaxUSD = 0.0
$totalDividendSEK = 0
$totalTaxSEK = 0

foreach ($ct in $cashTransactions)
{
	if ($ct.transactionType -eq "$+-DIV")
	{
		$txDate = [datetime]::ParseExact($ct.transactionDate, "MM/dd/yyyy", $null)
		$txRate = Get-SEKUSDPMIAtDate($txDate)
		
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
