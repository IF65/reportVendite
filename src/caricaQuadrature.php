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
$startingDate = (clone $currentDate)->sub(new DateInterval('P830D'));

// inizio
// -----------------------------------------------------------

try {
	// apro il datasource
	// -------------------------------------------------------------------------------
	$db = new PDO("mysql:host=$hostname", $user, $password, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);

	// creazione tabelle
	// -------------------------------------------------------------------------------
	$stmt = "	CREATE TABLE IF NOT EXISTS controllo.`quadrature` (
				  `store` varchar(4) NOT NULL DEFAULT '',
				  `ddate` date NOT NULL,
				  `totalamount` decimal(11,2) NOT NULL DEFAULT '0.00',
				  `customercount` int(10) unsigned NOT NULL DEFAULT '0',
				  `closed` tinyint(3) unsigned NOT NULL DEFAULT '0',
				  PRIMARY KEY (`store`,`ddate`)
				) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4";
	$h_create = $db->prepare($stmt);
	$h_create->execute();

	$stmt = "	insert into controllo.quadrature 
					(`store`,`ddate`,`totalamount`,`customercount`,`closed`)
				values
					(:store,:ddate,:totalamount,:customercount,:closed)
				on duplicate key update 
				    `totalamount`=:totalamount,
				    `customercount`=:customercount,
					`closed`=:closed";
	$h_insert_quadrature = $db->prepare($stmt);

	$client = new Client([
		'base_uri' => 'http://10.11.14.74/',
		'headers' => ['Content-Type' => 'application/json'],
		'timeout' => 60.0,
	]);

	$response = $client->post('dwh',
		['json' =>
			[
				'dataFine' => $currentDate->format(DATE_ATOM),
				'dataInizio' => $startingDate->format(DATE_ATOM),
				'function' => 'caricaQuadrature'
			]
		]
	);

	if ($response->getStatusCode() == 200) {
		$jsonResponse = $response->getBody()->getContents();
		$righe = json_decode($jsonResponse, true);
		if (isset($righe)) {
			foreach ($righe as $riga) {
				$h_insert_quadrature->execute([
					':store' => $riga['store'],
					':ddate' => (new DateTime($riga['ddate'], $timeZone))->format('Y-m-d'),
					':totalamount' => $riga['totalamount'],
					':customercount' => $riga['customercount'],
					':closed' => $riga['closed']
				]);
			}
		}
	} else {
		echo "\n";
	}
} catch (PDOException $e) {
	echo "Errore: " . $e->getMessage();
	die();
}