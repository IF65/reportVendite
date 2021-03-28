<?php

ini_set('memory_limit','8192M');

require 'vendor/autoload.php';

use GuzzleHttp\Client;

// costanti
// -----------------------------------------------------------
$hostname = "10.11.14.248";
$dbname = "mtx";
$user = "root";
$password = "mela";

$timeZone = new DateTimeZone('Europe/Rome');
$currentDate = new DateTime('now', $timeZone);
$startingDate = (clone $currentDate)->sub(new DateInterval('P7D'));

// inizio
// -----------------------------------------------------------

try {
	// apro il datasource
	// -------------------------------------------------------------------------------
	$db = new PDO("mysql:host=$hostname", $user, $password, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);

	// creazione tabelle
	// -------------------------------------------------------------------------------
	$stmt = "	CREATE TABLE IF NOT EXISTS mtx.hours (
				  `store` varchar(4) NOT NULL DEFAULT '',
				  `ddate` date NOT NULL,
				  `code` tinyint(3) unsigned NOT NULL,
				  `amount` decimal(11,2) NOT NULL,
				  `hours` decimal(11,2) NOT NULL,
				  PRIMARY KEY (`store`,`ddate`,`code`)
				) ENGINE=InnoDB DEFAULT CHARSET=latin1;";
	$h_create = $db->prepare($stmt);
	$h_create->execute();

	$stmt = "	CREATE TABLE IF NOT EXISTS mtx.`control` (
				  `store` varchar(4) NOT NULL DEFAULT '',
				  `ddate` date NOT NULL,
				  `totalamount` decimal(11,2) NOT NULL,
				  `totalhours` decimal(11,2) NOT NULL,
				  `customercount` int(10) unsigned NOT NULL,
				  `closed` tinyint(3) unsigned NOT NULL,
				  PRIMARY KEY (`store`,`ddate`)
				) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4";
	$h_create = $db->prepare($stmt);
	$h_create->execute();

	$stmt = "	insert into mtx.hours 
					(`store`,`ddate`,`code`,`amount`,`hours`)
				values
					(:store,:ddate,:code,:amount,:hours)
				on duplicate key update 
				    `code`=:code,
				    `amount`=:amount,
				    `hours`=:hours";
	$h_insert_hours = $db->prepare($stmt);

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
		$dati = file_get_contents('/Users/if65/Desktop/totali.json');
		$reparti = json_decode($dati, true);
		if (isset($reparti)) {
			foreach ($reparti as $reparto) {
				$h_insert_hours->execute([
					':store' => $reparto['store'],
					':ddate' => (new DateTime($reparto['ddate'], $timeZone))->format('Y-m-d'),
					':code' => $reparto['code'],
					':amount' => $reparto['amount'],
					':hours' => $reparto['hours']
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