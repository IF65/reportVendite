<?php

ini_set('memory_limit','8192M');

require __DIR__ . "/../vendor/autoload.php";

use GetOpt\GetOpt as Getopt;
use GetOpt\Option;

$timeZone = new DateTimeZone('Europe/Rome');

$ieri = (new DateTime('now', $timeZone))->sub(new DateInterval('P1D'))->format('Y-m-d');

$options = new GetOpt([
	Option::create('h', 'help', GetOpt::NO_ARGUMENT)
		->setDescription('Mostra questo help')->setDefaultValue(0),
	Option::create('d', 'debug', GetOpt::NO_ARGUMENT )
		->setDescription('Esegue la procedura in modalità di debug.')->setDefaultValue(0),
	Option::create('i', 'inizio', GetOpt::REQUIRED_ARGUMENT )
		->setDescription("Data inizio caricamento. (Default $ieri).")
		->setDefaultValue($ieri),
	Option::create('f', 'fine', GetOpt::REQUIRED_ARGUMENT )
		->setDescription('Data fine caricamento. (Se mancante viene presa come data di fine la data d\'inizio).'),
	Option::create('s', 'sede', GetOpt::REQUIRED_ARGUMENT )
		->setDescription('Sede da caricare.'),
	Option::create(null, 'deleteAll', GetOpt::NO_ARGUMENT )
		->setDescription('Elimina tutti i dati eventualmente presenti prima di caricare.')->setDefaultValue(0),
	Option::create(null, 'keepSales', GetOpt::NO_ARGUMENT )
		->setDescription('Elimina tutti i dati eventualmente presenti prima di caricare tranne la tabella sales.')->setDefaultValue(0),
]);

// process arguments and catch user errors
try {
	try {
		$options->process();
	} catch (Missing $exception) {
		// catch missing exceptions if help is requested
		if (!$options->getOption('help')) {
			throw $exception;
		}
	}
} catch (ArgumentException $exception) {
	file_put_contents('php://stderr', $exception->getMessage() . PHP_EOL);
	echo PHP_EOL . $options->getHelpText();
	exit;
}
$debug = $options->getOption('d');

$deletaAll = $options->getOption('deleteAll');
$keepSales = $options->getOption('keepSales');
if (! $deletaAll) {
	$keepSales = 0;
}

$help = $options->getOption('h');
if ($help) {
	echo PHP_EOL . $options->getHelpText();
	exit;
}

$dataInizio = $options->getOption('i');
if (!preg_match('/^\d{4}\-\d{2}\-\d{2}$/', $dataInizio)) {
	$dataInizio = '';
}

$dataFine = $options->getOption('f');
if (!preg_match('/^\d{4}\-\d{2}\-\d{2}$/', $dataFine)) {
	$dataFine = $dataInizio;
}

$sede = $options->getOption('s');
if (!preg_match('/^(00|01|02|04|05|31|36|60)\d{2}$/', $sede)) {
	$sede = '';
}

// identificazione server
// -------------------------------------------------------------------------------
$sourceHostname = '10.11.14.128';
$sourceUser = 'root';
$sourcePassword = 'mela';

$destinationHostname = '10.11.14.248';
$destinationUser = 'root';
$destinationPassword = 'mela';

if ($debug) {
	$destinationHostname = '10.11.14.177';
}

if ($sede != '') {
	try {
		// apro i datasource
		// -------------------------------------------------------------------------------
		$sourceDb = new PDO("mysql:host=$sourceHostname", $sourceUser, $sourcePassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);
		$destinationDb = new PDO("mysql:host=$destinationHostname", $destinationUser, $destinationPassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);

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
		$h_query = $destinationDb->prepare($stmt);
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
		$h_query = $destinationDb->prepare($stmt);
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
		$h_query = $destinationDb->prepare($stmt);
		$h_query->execute();

		// creo la tabella customers sul db destinazione
		// -------------------------------------------------------------------------------
		$stmt = "	CREATE TABLE IF NOT EXISTS mtx.`customers` (
					`store` varchar(4) NOT NULL DEFAULT '',
					`ddate` date NOT NULL,
					`customerCount` int(11) NOT NULL DEFAULT '0',
					PRIMARY KEY (`store`,`ddate`)
				) ENGINE=InnoDB DEFAULT CHARSET=latin1;";
		$h_query = $destinationDb->prepare($stmt);
		$h_query->execute();

		// creo la tabella customers sul db destinazione
		// -------------------------------------------------------------------------------
		$stmt = "	CREATE TABLE IF NOT EXISTS mtx.`control` (
					  `store` varchar(4) NOT NULL DEFAULT '',
					  `ddate` date NOT NULL,
					  `totalamount` decimal(11,2) NOT NULL DEFAULT '0.00',
					  `totalhours` decimal(11,2) NOT NULL DEFAULT '0.00',
					  `customercount` int(10) unsigned NOT NULL DEFAULT '0',
					  `closed` tinyint(3) unsigned NOT NULL DEFAULT '0',
					  `salesamount` decimal(11,2) NOT NULL DEFAULT '0.00',
					  `departmentamount` decimal(11,2) NOT NULL DEFAULT '0.00',
					  PRIMARY KEY (`store`,`ddate`)
					) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4";
		$h_query = $destinationDb->prepare($stmt);
		$h_query->execute();

		// inserisco i reparti mancanti nella giornata negozio per il caricamento ore
		// -------------------------------------------------------------------------------
		$stmt = "insert ignore into mtx.hours select :store store, :ddate ddate, id code, 0 amount, 0 hours, department  from departments;";
		$h_insert_missing_departments = $destinationDb->prepare($stmt);

		// carico la corrispondenza barcode => [codice articolo, reparto]
		// -------------------------------------------------------------------------------
		$stmt = "   select b.`BAR13-BAR2` barcode, b.`CODCIN-BAR2` codice, a.`IDSOTTOREPARTO` reparto 
                from archivi.barartx2 as b join dimensioni.articolo as a on b.`CODCIN-BAR2`=a.`CODICE_ARTICOLO`;";
		$h_query = $sourceDb->prepare($stmt);
		$h_query->execute();
		$result = $h_query->fetchAll(PDO::FETCH_ASSOC);
		$articoli = [];
		foreach ($result as $articolo) {
			$articoli[$articolo['barcode']] = ['codice' => $articolo['codice'], 'reparto' => $articolo['reparto']];
		}
		unset($result);

		// preparo le query di inserimento e successivocaricamento
		// -------------------------------------------------------------------------------
		$stmt = "   insert ignore into mtx.sales 
                    (store, ddate, reg, trans, department, barcode, articledepartment, articlecode, weight, rowCount, quantity, totalamount, totaltaxableamount, fidelityCard) 
                 	values
                    (:store, :ddate, :reg, :trans, :department, :barcode, :articledepartment, :articlecode, :weight, :rowCount, :quantity, :totalamount, :totaltaxableamount, :fidelityCard)";
		$h_insert_sales = $destinationDb->prepare($stmt);

		$stmt = "	select codice store from archivi.negozi where codice = :store 
					union all 
					select codiceTcPos store from archivi.negozi where codice = :store and codiceTcPos <> ''";
		$h_load_store_list = $sourceDb->prepare($stmt);
		$h_load_store_list->execute([':store' => $sede]);
		$value = $h_load_store_list->fetchAll(PDO::FETCH_COLUMN);
		$list = '\'' . implode('\',\'', $value) . '\'';

		$stmt = "   select case when store<>:store then :store else store end store, ddate, reg, trans, userno department, barcode, '' articledepartment, '' articlecode, 
                           0 weight, count(*) rowCount, sum(quantita) quantity, sum(totalamount) totalamount, 
                           sum(case when totalamount<0 then totaltaxableamount*-1 else totaltaxableamount end) totaltaxableamount, '' fidelityCard 
	                from mtx.idc 
	                where ddate = :data and binary recordtype = 'S' and recordcode1 = 1 and store in ($list)
	                group by 1,2,3,4,5,6
	                having totalamount <> 0";
		$h_load_idc = $sourceDb->prepare($stmt);

		// preparo la query di controllo esistenza dati (serve a evitare doppi caricamenti)
		// -------------------------------------------------------------------------------
		$stmt = "select ifnull(count(*),0) `recordCount` from mtx.sales where ddate = :ddate and store = :store";
		$h_count_sales = $destinationDb->prepare($stmt);

		// preparo la query di cancellazione record dalla tabella sales (serve in caso di opzione deleteAll)
		// -------------------------------------------------------------------------------
		$stmt = "delete from mtx.sales where ddate = :ddate and store = :store";
		$h_delete_sales = $destinationDb->prepare($stmt);

		// preparo la query di creazione record di salesPerDepartment (solo record record vuoti)
		// -------------------------------------------------------------------------------
		$stmt = "	insert ignore into mtx.salesPerDepartment
				select :store store, :ddate ddate, a.department, 0 totaltaxableamount, 0 rowCount, 0 quantity, 0 customerCount from 
					(select distinct nuovoReparto department from mtx.sottoreparto order by 1) as a;";
		$h_create_salesPerDepartment = $destinationDb->prepare($stmt);

		// preparo la query di cancellazione salesPerDepartment
		// -------------------------------------------------------------------------------
		$stmt = "delete from mtx.salesPerDepartment where ddate = :ddate and store = :store";
		$h_delete_salesPerDepartment = $destinationDb->prepare($stmt);

		// preparo la query di inserimento salesPerDepartment
		// -------------------------------------------------------------------------------
		$stmt = "	insert into mtx.salesPerDepartment
				select s.store, s.ddate, ifnull(r.nuovoReparto,'SCATOLAME') department, ifnull(sum(s.totaltaxableamount),0) totaltaxableamount, ifnull(sum(s.rowCount),0) rowCount, ifnull(sum(s.quantity),0) quantity, count(distinct s.reg, s.trans) customerCount 
				from mtx.sales as s left join mtx.sottoreparto as r on s.articledepartment = r.idsottoreparto 
				where s.ddate = :ddate and s.store = :store
				group by 1,2,3";
		$h_insert_salesPerDepartment = $destinationDb->prepare($stmt);

		// preparo la query di creazione record di penetrationPerSubtotal (solo record vuoti)
		// -------------------------------------------------------------------------------
		$stmt = "	insert ignore into mtx.penetrationPerSubtotal
				select :store store, :ddate ddate, a.subtotal, 0 customerCount 
				from (select distinct subtotali subtotal from mtx.sottoreparto order by 1) as a";
		$h_create_penetrationPerSubtotal = $destinationDb->prepare($stmt);

		// preparo la query di cancellazione penetrationPerSubtotal
		// -------------------------------------------------------------------------------
		$stmt = "delete from mtx.penetrationPerSubtotal where ddate = :ddate and store = :store";
		$h_delete_penetrationPerSubtotal = $destinationDb->prepare($stmt);

		// preparo la query di inserimento penetrationPerSubtotal
		// -------------------------------------------------------------------------------
		$stmt = "	insert into mtx.penetrationPerSubtotal
				select s.store, s.ddate, r.subtotali subtotal, count(distinct s.reg, s.trans) customerCount 
				from mtx.sales as s join mtx.sottoreparto as r on s.articledepartment = r.idsottoreparto 
				where ddate = :ddate and store = :store
				group by 1,2,3";
		$h_insert_penetrationPerSubtotal = $destinationDb->prepare($stmt);

		// preparo la query di creazione record di customers (solo record vuoti)
		// -------------------------------------------------------------------------------
		$stmt = "	insert ignore into mtx.customers
					select :store store, :ddate ddate, 0 customerCount";
		$h_create_customers = $destinationDb->prepare($stmt);

		// preparo la query di cancellazione customers
		// -------------------------------------------------------------------------------
		$stmt = "	delete from mtx.customers where ddate = :ddate and store = :store";
		$h_delete_customers = $destinationDb->prepare($stmt);

		// preparo la query di inserimento customers
		// -------------------------------------------------------------------------------
		$stmt = "	insert into mtx.customers
					select s.store, s.ddate, count(distinct s.reg, s.trans) customerCount 
					from mtx.sales as s where s.ddate = :ddate and store = :store group by 1,2;";
		$h_insert_customers = $destinationDb->prepare($stmt);

		// preparo la query di creazione record vuoto di control
		// -------------------------------------------------------------------------------
		$stmt = "	insert ignore into mtx.control (store, ddate, totalamount, totalhours, customercount, closed, salesamount, departmentamount) 
					values (:store, :ddate, 0.00, 0.0, 0, 0, 0.00, 0.00)";
		$h_create_record_control = $destinationDb->prepare($stmt);

		// preparo la query di update di control
		// -------------------------------------------------------------------------------
		$stmt = "	update mtx.control as c join 
    				(select ddate, store, sum(totaltaxableamount) totalamount from mtx.sales where ddate = :ddate and store = :store group by 1,2) as s on c.store=s.store and c.ddate=s.ddate join 
	    			(select ddate, store, sum(totaltaxableamount) totalamount from mtx.salesPerDepartment where ddate = :ddate and store = :store group by 1,2) as d on c.store=d.store 
					set c.salesamount = s.totalamount, c.departmentamount = d.totalamount 
					where c.ddate = :ddate and c.store = :store";
		$h_update_control = $destinationDb->prepare($stmt);

		// eseguo il caricamento dei dati
		// -------------------------------------------------------------------------------
		$data = DateTime::createFromFormat('Y-m-d', $dataInizio, $timeZone); ;
		while ($data <= DateTime::createFromFormat('Y-m-d', $dataFine, $timeZone)) {
			$inizioCaricamento = (new DateTime())->setTimezone($timeZone);
			echo "Negozio $sede inizio caricamento giornata del " . $data->format('Y-m-d') . ' : ' . $inizioCaricamento->format('H:i:s') . "\n";

			$h_insert_missing_departments->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);

			if ($deletaAll && ! $keepSales) {
				$h_delete_sales->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			}

			$h_count_sales->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			$value = $h_count_sales->fetch(PDO::FETCH_ASSOC);
			$salesRecordCount = $value['recordCount'] * 1;

			if (! $keepSales) {
				if (!$salesRecordCount) {
					$h_load_idc->execute(['data' => $data->format('Y-m-d'), ':store' => $sede]);
					$sales = $h_load_idc->fetchAll(PDO::FETCH_ASSOC);

					foreach ($sales as $sale) {
						$quantity = $sale['quantity'];
						$weight = $sale['weight'];
						$barcode = $sale['barcode'];
						if (preg_match('/^(2\d{6})00000.$/', $barcode, $matches)) {
							$barcode = $matches[1];
							$weight = $quantity;
							$quantity = $sale['rowCount'];
						}

						/**
						 * nella tabella barartx il barcode è un decimal quindi gli zero iniziali vengono persi.
						 * se dopo la prima ricerca l'articolo non è stato trovato e il barcode ha uno o più zeri
						 * iniziali li elimino e ripeto la ricerca.
						 */
						$articleCode = (key_exists($barcode, $articoli)) ? $articoli[$barcode]['codice'] : '';
						if ($articleCode == '') {
							if (preg_match('/^0+(\d*)$/', $barcode, $matches)) {
								$articleCode = (key_exists($matches[1], $articoli)) ? $articoli[$matches[1]]['codice'] : '';
							}
						}
						$articledepartment = (key_exists($barcode, $articoli)) ? $articoli[$barcode]['reparto'] : '';
						if ($articledepartment == '') {
							if ($sale['department'] > 9 && $sale['department'] < 100) {
								$articledepartment = '0001';
							} elseif ($sale['department'] < 10) {
								$articledepartment = str_pad($sale['department'] * 100, 4, "0", STR_PAD_LEFT);
							} else {
								$articledepartment = str_pad($sale['department'], 4, "0", STR_PAD_LEFT);
							}
						}
						$totaltaxableamount = $sale['totaltaxableamount'];
						if ($quantity < 0 && $totaltaxableamount > 0) {
							$totaltaxableamount = $totaltaxableamount * -1;
						}

						$h_insert_sales->execute([
							'store' => $sale['store'],
							'ddate' => $sale['ddate'],
							'reg' => $sale['reg'],
							'trans' => $sale['trans'],
							'department' => $sale['department'],
							'barcode' => $barcode,
							'articledepartment' => $articledepartment,
							'articlecode' => $articleCode,
							'weight' => $weight,
							'rowCount' => $sale['rowCount'],
							'quantity' => $quantity,
							'totalamount' => $sale['totalamount'],
							'totaltaxableamount' => $totaltaxableamount,
							'fidelityCard' => $sale['fidelityCard']
						]);
					}
				}
			}

			// creo e poi calcolo i consolidati per macro reparto
			$h_delete_salesPerDepartment->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			$h_insert_salesPerDepartment->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			$h_create_salesPerDepartment->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]); // <- creo i reparti che non si sono movimentati nella giornata

			// creo e poi calcolo la penetrazione per subtotale
			$h_delete_penetrationPerSubtotal->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			$h_insert_penetrationPerSubtotal->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			$h_create_penetrationPerSubtotal->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]); // <- creo i subtotali che non si sono movimentati nella giornata

			// creo e poi calcolo i clienti
			$h_delete_customers->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			$h_insert_customers->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			$h_create_customers->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]); // <- creo i clienti che non si sono movimentati nella giornata

			// aggiorno la tabella di controllo
			$h_create_record_control->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);
			$h_update_control->execute([':ddate' => $data->format('Y-m-d'), ':store' => $sede]);

			$fineCaricamento = (new DateTime())->setTimezone($timeZone);
			echo "Negozio $sede fine caricamento giornata del " . $data->format('Y-m-d') . ' : ' . $inizioCaricamento->format('H:i:s') . "\n";

			$data->add(new DateInterval('P1D'));
		}
	} catch (PDOException $e) {
		echo "Errore: " . $e->getMessage();
		die();
	}
} else {
	echo PHP_EOL . $options->getHelpText();
}