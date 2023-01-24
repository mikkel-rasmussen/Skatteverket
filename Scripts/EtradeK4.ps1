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
		$groups = $([xml]$(Invoke-WebRequest -Method Post -SkipHeaderValidation -Headers @{'Content-Type'='application/soap+xml;charset=utf-8;action=urn:getInterestAndExchangeRates'} -Body $body -Uri 'https://swea.riksbank.se/sweaWS/services/SweaWebServiceHttpSoap12Endpoint' -UseBasicParsing)."Content").Envelope.body.getInterestAndExchangeRatesResponse.return.groups
		# Scan backwards in time for the latest rate, if the response for the current is empty.
		if ($null -eq $groups) { Write-Host $([string]::Format("No exchange rate for {0}, trying the previous day", $Date.ToString("yyyy-MM-dd"))) }
		$Date = $Date.AddDays(-1.0)
	}
	while ($null -eq $groups)
	return [double]$groups.series.resultrows.value.'#text'
}

$gainsLosses = $(get-content -raw -path gainsLosses.json | ConvertFrom-Json).data.gainsAndLosses.list.gainsLossDtlsList

$k4 = @()
$totalClosing = 0
$totalOpening = 0
$totalGain = 0
$totalLoss = 0

foreach ($gl in $gainsLosses)
{
	$openingDate = [datetime]::ParseExact($gl.openingTransDateAcquired, "MM/dd/yyyy", $null)
	$closingDate = [datetime]::ParseExact($gl.closingTransDateSold, "MM/dd/yyyy", $null)
	$openingRate = Get-SEKUSDPMIAtDate($openingDate)
	$closingRate = Get-SEKUSDPMIAtDate($closingDate)
	
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
		Add-Member -Type NoteProperty -Name 'Förlust' -Value $loss -PassThru
	$k4 += $k4Row
}

$totals = New-Object PSObject |
	Add-Member -Type NoteProperty -Name 'Summa försäljningspris' -Value $totalClosing -PassThru |
	Add-Member -Type NoteProperty -Name 'Summa omkostnadsbelopp' -Value $totalOpening -PassThru |
	Add-Member -Type NoteProperty -Name 'Summa vinst' -Value $totalGain -PassThru |
	Add-Member -Type NoteProperty -Name 'Summa förlust' -Value $totalLoss -PassThru

$k4 | Format-Table
$totals