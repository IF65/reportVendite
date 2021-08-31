<?php

ini_set('memory_limit','8192M');

require(realpath(__DIR__ . '/if65') . '/vendor/autoload.php');

// costanti
// -----------------------------------------------------------

// LRP
$sourceHostname = '10.11.14.109';
$sourceUser = 'test';
$sourcePassword = 'user1';
$service_name   = "ORCL";
$sid            = "ORCL";
$port           = 1521;
$dbtns          = "(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = $sourceHostname)(PORT = $port)) (CONNECT_DATA = (SERVICE_NAME = $service_name) (SID = $sid)))";

// MTX
$destinationHostname = '10.11.14.248';
$destinationUser = 'root';
$destinationPassword = 'mela';

$timeZone = new DateTimeZone('Europe/Rome');
$currentDate = new DateTime('now', $timeZone);
$startingDate = (clone $currentDate)->sub(new DateInterval('P14D'));
//$startingDate = DateTime::createFromFormat('Y-m-d', '2020-01-01', $timeZone);

// inizio
// -----------------------------------------------------------

try {
	// apro i datasource LRP / MTX
	// ------------------------------------------------------------------------------
    $sourceDb = new PDO("oci:dbname=" . $dbtns . ";charset=utf8", $sourceUser, $sourcePassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION, PDO::ATTR_EMULATE_PREPARES => false, PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC]);
    $destinationDb = new PDO("mysql:host=$destinationHostname", $destinationUser, $destinationPassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);

    $stmt = "select P.CODICE AS negozio,
					to_char(to_date(t.ID_TEMPO,'J'),'YYYY-MM-DD') AS ddate,
                    --lpad(t.NUMERO_CASSA,3,'0') AS cassa,
					--lpad(t.NUMERO_SCONTRINO,4,'0') AS transazione,
					sum(nvl(t.IMPORTO,0)) AS totale_importo
                from DWHD.testate_202105 t  
                JOIN DWHD.pdv P ON t.ID_PDV = P.ID
			 where to_char(to_date(t.ID_TEMPO,'J'),'YYYY-MM-DD') = '2021-05-10'
			GROUP BY P.CODICE, to_char(to_date(t.ID_TEMPO,'J'),'YYYY-MM-DD')";
    $select_lrp = $sourceDb->prepare($stmt);

	$stmt = "	insert into mtx.control 
					(`store`,`ddate`,`lrptotalamount`)
				values
					(:store,:ddate,:lrptotalamount)
				on duplicate key update 
					`lrptotalamount`=:lrptotalamount";
	$insert_control = $destinationDb->prepare($stmt);

	$stmt = "	update mtx.control set lrptotalamount = :lrptotalamount where ddate = :ddate and store = :store";
	$update_control = $destinationDb->prepare($stmt);

    $select_lrp->execute();
    while ($row = $select_lrp->fetch()) {

        $insert_control->execute([
            ':store' => $row['negozio'],
            ':ddate' => (new DateTime($row['data'], $timeZone))->format('Y-m-d'),
            ':lrptotalamount' => $row['totale_importo']
        ]);
        // aggiornamento control con info da LRP
        //$update_control->execute([':ddate' => (new DateTime($row['ddate'], $timeZone))->format('Y-m-d'), ':lrptotalamount' => $row['totale_importo']]);
    }

} catch (PDOException $e) {
	echo "Errore: " . $e->getMessage();
	die();
}