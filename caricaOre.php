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

// inizio
// -----------------------------------------------------------

try {
	// apro il datasource
	// -------------------------------------------------------------------------------
	$db = new PDO("mysql:host=$hostname", $user, $password, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);

	// creo la tabella sales sul db destinazione
	// -------------------------------------------------------------------------------
	$stmt = "	CREATE TABLE IF NOT EXISTS mtx.ore (
					`sede` varchar(4) NOT NULL DEFAULT '',
					`data` date NOT NULL,
					`codiceReparto` tinyint(3) unsigned NOT NULL,
					`vendutoTotale` decimal(11,2) NOT NULL,
					`vendutoReparto` decimal(11,2) NOT NULL,
					`oreReparto` decimal(11,2) NOT NULL,
				PRIMARY KEY (`sede`,`data`,`codiceReparto`)
				) ENGINE=InnoDB DEFAULT CHARSET=latin1;";

	$h_create = $db->prepare($stmt);
	$h_create->execute();

	$stmt = "	insert into mtx.ore 
					(`sede`,`data`,`codiceReparto`,`vendutoTotale`,`vendutoReparto`,`oreReparto`)
				values
					(:sede,:data,:codiceReparto,:vendutoTotale,:vendutoReparto,:oreReparto)
				on duplicate key update 
				    `codiceReparto`=:codiceReparto,
				    `vendutoTotale`=:vendutoTotale,
				    `vendutoReparto`=:vendutoReparto,
				    `oreReparto`=:oreReparto";
	$h_insert = $db->prepare($stmt);

	$client = new Client([
		'base_uri' => 'http://10.11.14.74/',
		'headers' => ['Content-Type' => 'application/json'],
		'timeout' => 60.0,
	]);

	$response = $client->post('quadrature',
		['json' =>
			[
				'dataFine' => $currentDate->format(DATE_ATOM),
				'dataInizio' => $currentDate->sub(new DateInterval('P7D'))->format(DATE_ATOM),
				'function' => 'caricaOreLavorate'
			]
		]
	);

	if ($response->getStatusCode() == 200) {
		$reparti = json_decode($response->getBody()->getContents(), true);
		if (isset($reparti)) {
			foreach ($reparti as $reparto) {
				$h_insert->execute([
					':sede' => $reparto['sede'],
					':data' => (new DateTime($reparto['data'], $timeZone))->format('Y-m-d'),
					':codiceReparto' => $reparto['codiceReparto'],
					':vendutoTotale' => $reparto['vendutoTotale'],
					':vendutoReparto' => $reparto['vendutoReparto'],
					':oreReparto' => $reparto['oreReparto']
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