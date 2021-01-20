<?php

ini_set('memory_limit','8192M');

require 'vendor/autoload.php';

$timeZone = new DateTimeZone('Europe/Rome');

$dataInizio = new DateTime('2020-09-16', $timeZone);
$dataFine = new DateTime('2020-10-15', $timeZone);

// identificazione server
// -------------------------------------------------------------------------------
$sourceHostname = '10.11.14.128';
$sourceUser = 'root';
$sourcePassword = 'mela';

$destinationHostname = '10.11.14.177';
$destinationUser = 'root';
$destinationPassword = 'mela';

try {
    // apro i datasource
    // -------------------------------------------------------------------------------
    $sourceDb = new PDO( "mysql:host=$sourceHostname", $sourceUser, $sourcePassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION] );
    $destinationDb = new PDO( "mysql:host=$destinationHostname", $destinationUser, $destinationPassword, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION] );

    // creo (se non c'è) la tabella sales sul db destinazione
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

    // carico la corrispondenza barcode => [codice articolo, reparto]
    // -------------------------------------------------------------------------------
    $stmt = "   select b.`BAR13-BAR2` barcode, b.`CODCIN-BAR2` codice, a.`IDSOTTOREPARTO` reparto 
                from archivi.barartx2 as b join dimensioni.articolo as a on b.`CODCIN-BAR2`=a.`CODICE_ARTICOLO`;";

    $h_query = $sourceDb->prepare( $stmt );
    $h_query->execute();
    $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
    $articoli = [];
    foreach ($result as $articolo) {
        $articoli[$articolo['barcode']] = ['codice' => $articolo['codice'], 'reparto' => $articolo['reparto']];
    }
    unset($result);

    // preparo le  query di inserimento/caricamento
    // -------------------------------------------------------------------------------
    $stmt = "   insert ignore into mtx.sales 
                    (store, ddate, reg, trans, department, barcode, articledepartment, articlecode, weight, rowCount, quantity, totalamount, totaltaxableamount, fidelityCard) 
                 values
                    (:store, :ddate, :reg, :trans, :department, :barcode, :articledepartment, :articlecode, :weight, :rowCount, :quantity, :totalamount, :totaltaxableamount, :fidelityCard)";
    $h_insert_query = $destinationDb->prepare( $stmt );

    $stmt = "   select store, ddate, reg, trans, userno department, barcode, '' articledepartment, '' articlecode, 
                           0 weight, count(*) rowCount, sum(quantita) quantity, sum(totalamount) totalamount, 
                           sum(totaltaxableamount) totaltaxableamount, '' fidelityCard 
                    from mtx.idc 
                    where ddate = :data and binary recordtype = 'S' and recordcode1 = 1 and store = '0101'
                    group by 1,2,3,4,5,6
                    having totalamount <> 0";
    $h_load_query = $sourceDb->prepare( $stmt );

    // eseguo il caricamento dei dati
    // -------------------------------------------------------------------------------
    $data = clone $dataInizio;
    while ($data <= $dataFine) {
        $inizioCaricamento = (new DateTime())->setTimezone($timeZone);
        echo "Inizio caricamento giornata del " . $data->format( 'Y-m-d' ) . ' : ' . $inizioCaricamento->format( 'H:m:s' ) .  "\n";

        $h_load_query->execute( ['data' => $data->format( 'Y-m-d' )] );
        $sales = $h_load_query->fetchAll( PDO::FETCH_ASSOC );

        foreach ($sales as $sale) {
            $quantity = $sale['quantity'];
            $weight = $sale['weight'];
            $barcode = $sale['barcode'];
            if (preg_match( '/^(2\d{6})00000.$/', $barcode, $matches )) {
                $barcode = $matches[1];
                $weight = $quantity;
                $quantity = $sale['rowCount'];
            }

            $articleCode = (key_exists( $barcode, $articoli )) ? $articoli[$barcode]['codice'] : '';
            $articledepartment = (key_exists( $barcode, $articoli )) ? $articoli[$barcode]['reparto'] : '';
            if ($articledepartment == '') {
                if ($sale['department'] > 9 && $sale['department'] < 100) {
                    $articledepartment = '0001';
                } elseif ($sale['department'] < 10) {
                    $articledepartment = str_pad( $sale['department'] * 100, 4, "0", STR_PAD_LEFT );
                } else {
                    $articledepartment = str_pad( $sale['department'], 4, "0", STR_PAD_LEFT );
                }
            }
            $totaltaxableamount = $sale['totaltaxableamount'];
            if ($quantity < 0 && $totaltaxableamount > 0) {
                $totaltaxableamount = $totaltaxableamount * -1;
            }

            $h_insert_query->execute( [
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
            ] );
        }

        $fineCaricamento = (new DateTime())->setTimezone($timeZone);
        echo "Fine caricamento giornata del " . $data->format( 'Y-m-d' ) . ' : ' . $fineCaricamento->format( 'H:m:s' ) .  "\n";

        $interval = $inizioCaricamento->diff($fineCaricamento);
        echo "Tempo impiegato : " . $interval->format('%i min') ."\n\n";

        $data->add(new DateInterval('P1D'));
    }
} catch (PDOException $e) {
    echo "Errore: " . $e->getMessage();
    die();
}
