<?php

ini_set('memory_limit','8192M');

require(realpath(__DIR__ . '/..') . '/vendor/autoload.php');

use GuzzleHttp\Client;

// costanti
// -----------------------------------------------------------
$hostname = "10.11.14.248";
$dbname = "mtx";
$user = "root";
$password = "mela";

$timeZone = new DateTimeZone('Europe/Rome');
$currentDate = new DateTime('now', $timeZone);
$startingDate = (clone $currentDate)->sub(new DateInterval('P14D'));
//$startingDate = DateTime::createFromFormat('Y-m-d', '2020-01-01', $timeZone);

// inizio
// -----------------------------------------------------------

try {
	// apro il datasource
	// ------------------------------------------------------------------------------†-
	$db = new PDO("mysql:host=$hostname", $user, $password, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);

	// creazione tabelle
	// -------------------------------------------------------------------------------
	$stmt = "	CREATE TABLE IF NOT EXISTS mtx.`hours` (
				  `store` varchar(4) NOT NULL DEFAULT '',
				  `ddate` date NOT NULL,
				  `code` tinyint(3) unsigned NOT NULL,
				  `amount` decimal(11,2) NOT NULL,
				  `hours` decimal(11,2) NOT NULL,
				  `department` varchar(100) NOT NULL DEFAULT '',
				  PRIMARY KEY (`store`,`ddate`,`department`),
				  KEY `store` (`store`,`ddate`,`code`)
				) ENGINE=InnoDB DEFAULT CHARSET=latin1;";
	$h_create = $db->prepare($stmt);
	$h_create->execute();

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
	$h_create = $db->prepare($stmt);
	$h_create->execute();

	$stmt = "	CREATE TABLE IF NOT EXISTS mtx.`departments` (
				  `id` int(11) NOT NULL,
				  `department` varchar(100) NOT NULL DEFAULT '',
				  PRIMARY KEY (`department`),
				  KEY `id` (`id`)
				) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
	$h_create = $db->prepare($stmt);
	$h_create->execute();

	$stmt = "select id code, department from mtx.departments order by 1";
	$h_load_departments = $db->prepare($stmt);
	$h_load_departments->execute();
	$result = $h_load_departments->fetchAll(PDO::FETCH_ASSOC);
	$departments = [];
	foreach($result as $department) {
		$departments[$department['code']] = $department['department'];
	}

	$stmt = "	insert into mtx.hours 
					(`store`,`ddate`,`code`,`amount`,`hours`,`department`)
				values
					(:store,:ddate,:code,:amount,:hours, :department)
				on duplicate key update 
				    `code`=:code,
				    `amount`=:amount,
				    `hours`=:hours,
					`department`=:department";
	$h_insert_hours = $db->prepare($stmt);

	$stmt = "	insert into mtx.hours
				select s.store, s.ddate, s.code, 0 amount, 0 hours, d.`department` department 
				from (select distinct h.store, h.ddate, d.id code from mtx.hours h join mtx.departments as d where h.store = :store and h.ddate = :ddate order by 1,2,3) as s left join 
				      mtx.hours as h on s.store=h.store and s.ddate=h.ddate and s.code=h.code left join mtx.departments as d on s.code=d.`id` 
				where h.code is null";
	$stmt = "insert ignore into mtx.hours select :store store, :ddate ddate, id code, 0 amount, 0 hours, department  from departments;";
	$h_insert_missing_departments = $db->prepare($stmt);

	$stmt = "	insert into mtx.control 
					(`store`,`ddate`,`totalamount`,`totalhours`,`customercount`,`closed`)
				values
					(:store,:ddate,:totalamount,:totalhours,:customercount,:closed)
				on duplicate key update 
				    `totalamount`=:totalamount,
				    `totalhours`=:totalhours,
				    `customercount`=:customercount,
					`closed`=:closed";
	$h_insert_control = $db->prepare($stmt);

	$stmt = "	update mtx.control as c join (select store, ddate, sum(hours) hours from mtx.hours group by 1,2) as h on c.store=h.store and c.ddate=h.ddate 
				set c.totalhours = h.hours where h.ddate >= :ddate";
	$h_update_control = $db->prepare($stmt);

	$client = new Client([
		'base_uri' => 'http://10.11.14.74/',
		'headers' => ['Content-Type' => 'application/json'],
		'timeout' => 60.0,
	]);

	$response = $client->post('quadrature',
		['json' =>
			[
				'dataFine' => $currentDate->format(DATE_ATOM),
				'dataInizio' => $startingDate->format(DATE_ATOM),
				'function' => 'caricaOreLavorate'
			]
		]
	);

	if ($response->getStatusCode() == 200) {
		$reparti = json_decode($response->getBody()->getContents(), true);
		if (isset($reparti)) {
			foreach ($reparti as $reparto) {
				$h_insert_missing_departments->execute([
					':store' => $reparto['store'],
					':ddate' => (new DateTime($reparto['ddate'], $timeZone))->format('Y-m-d')
				]);

				$department = 'ALTRO';
				if (key_exists($reparto['code'], $departments)) {
					$department = $departments[$reparto['code']];
				}

				$h_insert_hours->execute([
					':store' => $reparto['store'],
					':ddate' => (new DateTime($reparto['ddate'], $timeZone))->format('Y-m-d'),
					':code' => $reparto['code'],
					':amount' => $reparto['amount'],
					':hours' => $reparto['hours'],
					':department' => $department
				]);

				$h_insert_control->execute([
					':store' => $reparto['store'],
					':ddate' => (new DateTime($reparto['ddate'], $timeZone))->format('Y-m-d'),
					':totalamount' => $reparto['totalamount'],
					':totalhours' => 0,
					':customercount' => $reparto['customercount'],
					':closed' => $reparto['closed']
				]);
			}
			$h_update_control->execute([':ddate' => ($startingDate->format('Y-m-d'))]);
		}
	} else {
		echo "\n";
	}
} catch (PDOException $e) {
	echo "Errore: " . $e->getMessage();
	die();
}