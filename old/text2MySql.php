<?php

ini_set('memory_limit','8192M');

require 'vendor/autoload.php';

$timeZone = new DateTimeZone('Europe/Rome');

// cartelle di caricamento
// -------------------------------------------------------------------------------
$loadDir = '/Users/if65/Desktop/consolidatiReparti/daTrasformare/';
$mySqlDir = '/Users/if65/Desktop/consolidatiReparti/daCaricare/';

// identificazione server
// -------------------------------------------------------------------------------
$sourceHostname = '10.11.14.128';
$sourceUser = 'root';
$sourcePassword = 'mela';


try {
	// apro il datasource
	// -------------------------------------------------------------------------------
	$sourceDb = new PDO( "mysql:host=$sourceHostname", $sourceUser, $sourcePassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION] );

		// carico la corrispondenza barcode => [codice articolo, reparto]
	// -------------------------------------------------------------------------------
	$stmt = "   select b.`BAR13-BAR2` barcode, b.`CODCIN-BAR2` codice, a.`IDSOTTOREPARTO` reparto 
                from archivi.barartx2 as b join dimensioni.articolo as a on b.`CODCIN-BAR2`=a.`CODICE_ARTICOLO`;";

	$h_query = $sourceDb->prepare( $stmt );
	$h_query->execute();
	$result = $h_query->fetchAll( PDO::FETCH_ASSOC );
	$articles = [];
	foreach ($result as $article) {
		$articles[$article['barcode']] = ['codice' => $article['codice'], 'reparto' => $article['reparto']];
	}
	unset($result);

	// arricchisco i file di testo per prepararlo al caricamento
	// -------------------------------------------------------------------------------
	$fileList = scandir($loadDir);
	foreach ($fileList as $fileName) {
		$result = [];

		if (preg_match('/^[^\.].*\.txt$/', $fileName)) {
			$file = file_get_contents($loadDir . $fileName);

			$rows = explode("\r\n", $file);
			foreach($rows as $row) {
				if(preg_match('/^(\d{4})\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t([^\t]*)\t$/', $row, $matches)) {
					$store = $matches[1];
					$ddate = $matches[2];
					$reg = $matches[3];
					$trans = $matches[4];
					$department = $matches[5];
					$barcode = $matches[6];
					$articledepartment = $matches[7];
					$articlecode = $matches[8];
					$weight = $matches[9] * 1;
					$rowCount =$matches[10] * 1;
					$quantity = $matches[11] * 1;
					$totalamount = $matches[12] * 1;
					$totaltaxableamount = $matches[13] * 1;
					$fidelityCard = '';

					$barcodeToSearch = $barcode;
					if (preg_match('/^(\d{7})00000\d$/', $barcode, $matches)) {
						$barcodeToSearch = $matches[1];
						$weight = $quantity;
						if ($quantity > 0) { $quantity = 1; } else { $quantity = -1; };
					}

					$articlecode = (key_exists($barcodeToSearch, $articles)) ? $articles[$barcodeToSearch]['codice'] : '';
					$articledepartment = (key_exists($barcodeToSearch, $articles)) ? $articles[$barcodeToSearch]['reparto'] : '';
					if ($articledepartment == '') {
						if ($department > 9 && $department < 100) {
							$articledepartment = '0001';
						} elseif ($department < 10) {
							$articledepartment = str_pad($department * 100, 4, "0", STR_PAD_LEFT);
						} else {
							$articledepartment = str_pad($department, 4, "0", STR_PAD_LEFT);
						}
					}

					if ($quantity < 0 && $totaltaxableamount > 0) {
						$totaltaxableamount = $totaltaxableamount * -1;
					}

					$result[] = sprintf("%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s",
						$store,
						$ddate,
						$reg,
						$trans,
						$department,
						$barcode,
						$articledepartment,
						$articlecode,
						$weight,
						$rowCount,
						$quantity,
						$totalamount,
						$totaltaxableamount,
						$fidelityCard
					);
				}
			}
			file_put_contents($mySqlDir . $fileName, implode("\r\n",$result));
		}
	}
} catch (PDOException $e) {
	echo "Errore: " . $e->getMessage();
	die();
}
