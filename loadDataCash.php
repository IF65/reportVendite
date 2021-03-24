<?php

ini_set('memory_limit','8192M');

require 'vendor/autoload.php';

$timeZone = new DateTimeZone('Europe/Rome');

if ($argc == 1) {
	$dataInizio = new DateTime('2020-01-01', $timeZone);
	$dataFine = new DateTime('2020-12-31', $timeZone);
} elseif ($argc == 2) {
	$dataInizio = new DateTime($argv[1], $timeZone);
	$dataFine = new DateTime($argv[1], $timeZone);
} else {
	$dataInizio = new DateTime($argv[1], $timeZone);
	$dataFine = new DateTime($argv[2], $timeZone);
}


// identificazione server
// -------------------------------------------------------------------------------
$sourceHostname = '10.11.14.128';
$sourceUser = 'root';
$sourcePassword = 'mela';

$destinationHostname = '10.11.14.248';
$destinationUser = 'root';
$destinationPassword = 'mela';

try {
    // apro i datasource
    // -------------------------------------------------------------------------------
    $sourceDb = new PDO( "mysql:host=$sourceHostname", $sourceUser, $sourcePassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION] );
    $destinationDb = new PDO( "mysql:host=$destinationHostname", $destinationUser, $destinationPassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION] );

    // creo la tabella sales sul db destinazione
    // -------------------------------------------------------------------------------
    $stmt = "   CREATE TABLE IF NOT EXISTS mtx.sales (
                      `store` varchar(4) NOT NULL DEFAULT '',
                      `ddate` date NOT NULL,
                      `reg` varchar(3) NOT NULL DEFAULT '',
                      `trans` varchar(4) NOT NULL DEFAULT '',
                      `department` smallint(5) unsigned NOT NULL,
                      `barcode` varchar(13) NOT NULL,
                      `articledepartment` varchar(4) NOT NULL DEFAULT '',
                      `articlecode` varchar(7) NOT NULL DEFAULT '',
                      `weight` decimal(11,3) NOT NULL DEFAULT '0.000',
                      `rowCount` smallint(5) unsigned NOT NULL DEFAULT '0',
                      `quantity` decimal(11,3) NOT NULL DEFAULT '0.000',
                      `totalamount` decimal(11,2) NOT NULL DEFAULT '0.00',
                      `totaltaxableamount` decimal(11,2) NOT NULL DEFAULT '0.00',
                      `fidelityCard` varchar(13) NOT NULL DEFAULT '',
                      PRIMARY KEY (`store`,`ddate`,`reg`,`trans`,`department`,`barcode`),
                      KEY `barcode` (`barcode`),
                      KEY `articlecode` (`articlecode`),
                      KEY `articledepartment` (`articledepartment`),
                      KEY `ddate` (`ddate`,`store`,`articledepartment`),
                      KEY `weight` (`weight`),
                      KEY `fidelityCard` (`fidelityCard`)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
    $h_query = $destinationDb->prepare( $stmt );
    $h_query->execute();

	// creo la tabella salesPerDepartment sul db destinazione
    // -------------------------------------------------------------------------------
	$stmt = "   CREATE TABLE IF NOT EXISTS mtx.salesPerDepartment (
					`store` varchar(4) NOT NULL DEFAULT '',
					`ddate` date NOT NULL,
					`department` varchar(100) NOT NULL DEFAULT '',
					`totaltaxableamount` decimal(11,2) NOT NULL DEFAULT '0.00',
					`rowCount` int(11) NOT NULL DEFAULT '0',
  					`quantity` int(11) NOT NULL DEFAULT '0',
					`customerCount` int(11) NOT NULL DEFAULT '0',
				PRIMARY KEY (`store`,`ddate`,`department`)
				) ENGINE=InnoDB DEFAULT CHARSET=latin1;";
	$h_query = $destinationDb->prepare( $stmt );
	$h_query->execute();

	// creo la tabella penetrationPerSubtotal sul db destinazione
	// -------------------------------------------------------------------------------
	$stmt = "	CREATE TABLE IF NOT EXISTS mtx.`penetrationPerSubtotal` (
				  `store` varchar(4) NOT NULL DEFAULT '',
				  `ddate` date NOT NULL,
				  `subtotal` varchar(100) NOT NULL DEFAULT '',
				  `customerCount` int(11) NOT NULL,
				  PRIMARY KEY (`store`,`ddate`,`subtotal`)
				) ENGINE=InnoDB DEFAULT CHARSET=latin1;";
	$h_query = $destinationDb->prepare( $stmt );
	$h_query->execute();

	// creo la tabella customers sul db destinazione
	// -------------------------------------------------------------------------------
	$stmt = "	CREATE TABLE IF NOT EXISTS mtx.`customers` (
					`store` varchar(4) NOT NULL DEFAULT '',
					`ddate` date NOT NULL,
					`customerCount` int(11) NOT NULL DEFAULT '0',
					PRIMARY KEY (`store`,`ddate`)
				) ENGINE=InnoDB DEFAULT CHARSET=latin1;";
	$h_query = $destinationDb->prepare( $stmt );
	$h_query->execute();


    // preparo la query di controllo esistenza dati (serve a evitare doppi caricamenti)
	// -------------------------------------------------------------------------------
	$stmt = "select ifnull(count(*),0) `recordCount` from mtx.sales where ddate = :ddate and store in ('0012','0016','0018')";
	$h_count_sales_query = $destinationDb->prepare( $stmt );

	// preparo la query di creazione salesPerDepartment (record vuoti)
	// -------------------------------------------------------------------------------
	$stmt = "	insert ignore into mtx.salesPerDepartment
				select a.store, :ddate ddate, b.department, 0 totaltaxableamount, 0 rowCount, 0 quantity, 0 customerCount from 
					(select codice store from archivi.negozi where societa in ('02','05') and codice like '00%' and data_inizio <= :ddate and (data_fine >= :ddate or data_fine is null)) as a join
					(select distinct nuovoReparto department from mtx.sottoreparto order by 1) as b;";
	$h_create_salesPerDepartment = $destinationDb->prepare( $stmt );

	// preparo la query di cancellazione salesPerDepartment
	// -------------------------------------------------------------------------------
	$stmt = "delete from mtx.salesPerDepartment where ddate = :ddate and store in ('0012','0016','0018')";
	$h_delete_salesPerDepartment = $destinationDb->prepare( $stmt );

	// preparo la query di inserimento salesPerDepartment
	// -------------------------------------------------------------------------------
	$stmt = "	insert into mtx.salesPerDepartment
				select s.store, s.ddate, r.nuovoReparto department, ifnull(sum(s.totaltaxableamount),0) totaltaxableamount, ifnull(sum(s.rowCount),0) rowCount, ifnull(sum(s.quantity),0) quantity, count(distinct s.reg, s.trans) customerCount 
				from mtx.sales as s join mtx.sottoreparto as r on s.articledepartment = r.idsottoreparto 
				where s.ddate = :ddate and store in ('0012','0016','0018')
				group by 1,2,3";
	$h_insert_salesPerDepartment = $destinationDb->prepare( $stmt );

	// preparo la query di creazione penetrationPerSubtotal (record vuoti)
	// -------------------------------------------------------------------------------
	$stmt = "	insert ignore into mtx.penetrationPerSubtotal
				select a.store, :ddate ddate, b.subtotal, 0 customerCount 
				from (
				    select codice store 
				    from archivi.negozi 
				    where societa in ('02','05') and codice like '00%' and data_inizio <= :ddate and (data_fine >= :ddate or data_fine is null)
				    ) as a join (select distinct subtotali subtotal from mtx.sottoreparto order by 1) as b";
	$h_create_penetrationPerDepartment = $destinationDb->prepare( $stmt );

	// preparo la query di cancellazione penetrationPerSubtotal
	// -------------------------------------------------------------------------------
	$stmt = "delete from mtx.penetrationPerSubtotal where ddate = :ddate and store in ('0012','0016','0018')";
	$h_delete_penetrationPerDepartment = $destinationDb->prepare( $stmt );

	// preparo la query di inserimento penetrationPerSubtotal
	// -------------------------------------------------------------------------------
	$stmt = "	insert into mtx.penetrationPerSubtotal
				select s.store, s.ddate, r.subtotali subtotal, count(distinct s.reg, s.trans) customerCount 
				from mtx.sales as s join mtx.sottoreparto as r on s.articledepartment = r.idsottoreparto 
				where ddate = :ddate and store in ('0012','0016','0018')
				group by 1,2,3";
	$h_insert_penetrationPerDepartment = $destinationDb->prepare( $stmt );

	// preparo la query di creazione customers (record vuoti)
	// -------------------------------------------------------------------------------
	$stmt = "	insert ignore into mtx.customers
				select codice store, :ddate ddate, 0 customerCount
				from archivi.negozi 
				where societa in ('02','05') and codice like '00%' and data_inizio <= :ddate and (data_fine >= :ddate or data_fine is null)";
	$h_create_customers = $destinationDb->prepare( $stmt );

	// preparo la query di cancellazione customers
	// -------------------------------------------------------------------------------
	$stmt = "delete from mtx.customers where ddate = :ddate and store in ('0012','0016','0018')";
	$h_delete_customers = $destinationDb->prepare( $stmt );

	// preparo la query di inserimento customers
	// -------------------------------------------------------------------------------
	$stmt = "	insert into mtx.customers
				select s.store, s.ddate, count(distinct s.reg, s.trans) customerCount from mtx.sales as s where s.ddate = :ddate  and store in ('0012','0016','0018') group by 1,2;";
	$h_insert_customers = $destinationDb->prepare( $stmt );

	// eseguo il caricamento dei dati
    // -------------------------------------------------------------------------------
    $data = clone $dataInizio;
    while ($data <= $dataFine) {
	    $inizioCaricamento = (new DateTime())->setTimezone($timeZone);
	    echo "Inizio caricamento giornata del " . $data->format('Y-m-d') . ' : ' . $inizioCaricamento->format('H:m:s') . "\n";

	    $h_count_sales_query->execute([':ddate' => $data->format('Y-m-d')]);
	    $value = $h_count_sales_query->fetch(PDO::FETCH_ASSOC);
	    $salesRecordCount = $value['recordCount'] * 1;

	    // creo e poi calcolo i consolidati per macro reparto
	    $h_delete_salesPerDepartment->execute([':ddate' => $data->format('Y-m-d')]);
	    $h_insert_salesPerDepartment->execute([':ddate' => $data->format('Y-m-d')]);
	    $h_create_salesPerDepartment->execute([':ddate' => $data->format('Y-m-d')]); // <- creo i reparti che non si sono movimentati nella giornata

	    // creo e poi calcolo la penetrazione per subtotale
	    $h_delete_penetrationPerDepartment->execute([':ddate' => $data->format('Y-m-d')]);
	    $h_insert_penetrationPerDepartment->execute([':ddate' => $data->format('Y-m-d')]);
	    $h_create_penetrationPerDepartment->execute([':ddate' => $data->format('Y-m-d')]); // <- creo i subtotali che non si sono movimentati nella giornata

	    // creo e poi calcolo i clienti
	    $h_delete_customers->execute([':ddate' => $data->format('Y-m-d')]);
	    $h_insert_customers->execute([':ddate' => $data->format('Y-m-d')]);
	    $h_create_customers->execute([':ddate' => $data->format('Y-m-d')]); // <- creo i clienti che non si sono movimentati nella giornata


	    $fineCaricamento = (new DateTime())->setTimezone($timeZone);
	    echo "Fine caricamento giornata del " . $data->format('Y-m-d') . "\n";

	    $data->add(new DateInterval('P1D'));
    }
} catch (PDOException $e) {
    echo "Errore: " . $e->getMessage();
    die();
}
