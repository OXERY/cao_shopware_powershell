#Variablen - CAO-Verbindung
$caoserver = "127.0.0.1"
$caousername = "USER"
$caopassword = "PASSWORD"
$caodatabase = "DBNAME"

#Variablen - Shopverbindung
$shopwareserver= "SWDBSERVER"
$shopwareusername= "SWDBUSER"
$shopwarepassword= "SWDBPASSWORD"
$shopwaredatabase= "SWDBNAME"

#Variablen - Pfade
$outputdir = $(Get-Location).Path
[void][system.reflection.Assembly]::LoadFrom($outputdir+"\Assemblies\v4.5\MySQL.Data.dll")
$StartTime = Get-Date

#Funktionen
function global:Set-SqlConnection ( $SqlConnection, $server = $(Read-Host "SQL Server Name"), $username = $(Read-Host "Username"), $password = $(Read-Host "Password"), $database = $(Read-Host "Default Database") ) {
	$SqlConnection.ConnectionString = "server=$server;user id=$username;password=$password;database=$database;pooling=false;Allow Zero Datetime=True;"
}

function global:Get-SqlDataTable( $SqlConnection, $Query = $(if (-not ($Query -gt $null)) {Read-Host "Query to run"}) ) {
	if (-not ($SqlConnection.State -like "Open")) { $SqlConnection.Open() }
	$SqlCmd = New-Object MySql.Data.MySqlClient.MySqlCommand $Query, $SqlConnection
	$SqlAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet) | Out-Null
	$SqlConnection.Close()
	return $DataSet.Tables[0]
}

#CAO
$DBConnection=New-Object System.Data.Odbc.OdbcConnection
$DBConnection.ConnectionString="Driver=MySQL ODBC 3.51 Driver;Server=$caoserver;Database=$caodatabase;UID=$caousername;PWD=$caopassword;"
$DBConnection.Open()
$DBCommand=New-Object System.Data.Odbc.OdbcCommand
$DBCommand.Connection=$DBConnection
$DBCommand.CommandText = 'SELECT ARTNUM, MENGE_AKT, VK2 AS PREISNETTO, VK2B AS PREISBRUTTO, USERFELD_08 AS LIEFERZEIT FROM  ARTIKEL ORDER BY ARTNUM'
$caoarticles = New-Object System.Data.DataSet
(New-Object System.Data.Odbc.OdbcDataAdapter($DBCommand)).Fill($caoarticles) | Out-Null
$DBConnection.Close()
Write-Host Anzahl CAO-Artikel: $caoarticles.tables.rows.Count

#shopware
Set-Variable shopwareSqlConnection (New-Object MySql.Data.MySqlClient.MySqlConnection) -Scope Global -Option AllScope -Description "Personal variable for Sql Query functions"
Set-SqlConnection $shopwareSqlConnection $shopwareserver $shopwareusername $shopwarepassword $shopwaredatabase
$shopwarearticles = Get-SqlDataTable $shopwareSqlConnection 'SELECT dt.articleID, dt.ordernumber, dt.instock, dt.active, dt.shippingtime, pr.price FROM s_articles AS a INNER JOIN s_articles_details AS dt ON a.id = dt.articleID INNER JOIN s_articles_prices AS pr ON a.id = pr.articleID'
Write-Host Anzahl Shopware-Artikel: $shopwarearticles.count

$updatecount = 0
$updatearticles = ""
$updatearticledetails = ""
$updatearticleprices = ""
foreach ($shopwarearticle in $shopwarearticles) {
    $price = 0
    $active = 1
    $caopreisnetto = 0
    $update = $false
    $price =  [math]::Round($($shopwarearticle.price*1.19),2)
    $caoarticle = ""

    foreach ($caoarticle in $caoarticles.tables.rows) {
        if ($caoarticle.ARTNUM -eq $shopwarearticle.ordernumber) {
            $caopreisnetto = [math]::Round($($caoarticle.PREISBRUTTO/1.19),10)
            if ($caoarticle.MENGE_AKT -ne $shopwarearticle.instock) {
                Write-Host Menge abweichend: $caoarticle.ARTNUM $caoarticle.MENGE_AKT $shopwarearticle.ordernumber $shopwarearticle.instock
                $update = $true
            }
            if ($shopwarearticle.shippingtime -ne $caoarticle.LIEFERZEIT) {
                Write-Host Lieferzeit abweichend: $caoarticle.ARTNUM $caoarticle.LIEFERZEIT $shopwarearticle.ordernumber $shopwarearticle.shippingtime
                $update = $true
            }
            if ($caoarticle.MENGE_AKT -eq 0 -and $shopwarearticle.active -eq 1) {
                Write-Host Aktiv, obwohl 0 Lager: $caoarticle.ARTNUM $caoarticle.MENGE_AKT $shopwarearticle.ordernumber $shopwarearticle.instock
                $active = 0
                $update = $true
            }
            
            if ($caoarticle.PREISBRUTTO -lt 0.01 -and $shopwarearticle.active -eq 1) {
                Write-Host Preis auf 0 und aktiv!
                $active = 0
                $update = $true
            }
            if ($caoarticle.PREISBRUTTO -ne $price) {
                Write-Host Preis weicht ab: $shopwarearticle.articleID $caoarticle.ARTNUM $caoarticle.PREISBRUTTO $swprice.articleID $swprice.price $price
                $update = $true
            }

            if ($update -eq $true) {
                $updatearticles += 'UPDATE s_articles SET shippingtime="' + $caoarticle.LIEFERZEIT + '", active='+$active+' WHERE id = ' + $shopwarearticle.articleID + ';'
                $updatearticledetails += 'UPDATE s_articles_details SET instock=' + [convert]::ToInt32($caoarticle.MENGE_AKT) + ', active=' + $active + ', shippingtime="' + $caoarticle.LIEFERZEIT + '" WHERE articleID='+$shopwarearticle.articleID + ';'
                $updatearticleprices += 'UPDATE s_articles_prices SET price=' + $caopreisnetto + ' WHERE articleID = '+ $shopwarearticle.articleID + ';'
                $updatecount++
            }
            break
        }
    }
}

if ($updatearticles.Length -gt 0) {
    Write-Host Es werden $updatecount Artikel aktualisiert
    Get-SqlDataTable $shopwareSqlConnection $updatearticles
    Get-SqlDataTable $shopwareSqlConnection $updatearticledetails
    Get-SqlDataTable $shopwareSqlConnection $updatearticleprices
}

$EndTime = (Get-Date) - $StartTime
Write-Host Durchgeführte Updates: $updatecount in $($Endtime.Minutes)Min $($Endtime.Seconds)s
