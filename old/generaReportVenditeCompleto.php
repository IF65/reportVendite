<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use GuzzleHttp\Client;


$timeZone = new DateTimeZone('Europe/Rome');


// inizio e fine settimana corrente
$dataCorrenteAC = new DateTime('2021-01-16', $timeZone);
$dataInizioAC = new DateTime('2021-01-11', $timeZone);
$dataFineAC = new DateTime('2021-01-17', $timeZone);
$dataCorrenteAP = new DateTime('2020-01-18', $timeZone);
$dataInizioAP = new DateTime('2020-01-13', $timeZone);
$dataFineAP = new DateTime('2020-01-19', $timeZone);

$hostname = 'localhost';
$user = 'root';
$password = 'mela';

$repartoIndefinito = 'NON DEFINITO';

$plusOneDay = new DateInterval('P1D');

try {

    if (true) {
        $db = new PDO( "mysql:host=$hostname", $user, $password, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION] );

        // recupero elenco negozi aperti nell'intervallo date prescelto
        // -------------------------------------------------------------------------------
        $stmt = "   select n.codice code, n.negozio_descrizione description 
                    from archivi.negozi as n 
                    where n.`data_inizio` <= :dataInizio and (n.`data_fine`>=:dataFine or n.`data_fine`is null) and 
                          n.`societa` in ('02','05') and n.`codice` not like '00%' 
                    order by 1;";
        $h_query = $db->prepare( $stmt );
        $h_query->execute( [':dataInizio' => $dataInizioAC->format( 'Y-m-d' ), ':dataFine' => $dataFineAC->format( 'Y-m-d' )] );
        $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
        $shopList = [];
        foreach ($result as $shop) {
            $shopList[$shop['code']] = $shop['description'];
        }

        // recupero elenco reparti (nuovi)
        // -------------------------------------------------------------------------------
        $stmt = "   select distinct `nuovoReparto`
                from mtx.sottoreparto 
                order by sortOrder";
        $h_query = $db->prepare( $stmt );
        $h_query->execute();
        $departmentList = $h_query->fetchAll( PDO::FETCH_COLUMN );
        $departmentCount = sizeof( $departmentList );

        // recupero corrispondeza sottoreparto/reparto (nuovo)
        // -------------------------------------------------------------------------------
        $stmt = "   select `idsottoreparto`, `nuovoReparto`
                    from mtx.sottoreparto 
                    order by 1,2";
        $h_query = $db->prepare( $stmt );
        $h_query->execute();
        $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
        $departmentTransCode = [];
        foreach ($result as $department) {
            $departmentTransCode[$department['idsottoreparto']] = $department['nuovoReparto'];
        }

        // recupero composizione subtotali
        // -------------------------------------------------------------------------------
        $stmt = "   select distinct  `subtotali`,`nuovoReparto`
                    from mtx.sottoreparto 
                    order by 1";
        $h_query = $db->prepare( $stmt );
        $h_query->execute();
        $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
        $subtotalTransCode = [];
        $subtotalDescription = [];
        foreach ($result as $subtotal) {
            $subtotalTransCode[$subtotal['nuovoReparto']] = $subtotal['subtotali'];
            if (!in_array( $subtotal['subtotali'], $subtotalDescription )) {
                $subtotalDescription[] = $subtotal['subtotali'];
            }
        }
        $subtotalCount = sizeof( $subtotalDescription );

        // recupero ore lavorate per reparto
        // -------------------------------------------------------------------------------
        $client = new Client([
            'base_uri' => 'http://10.11.14.74/',
            'headers' => ['Content-Type' => 'application/json'],
            'timeout'  => 200,
        ]);

        $response = $client->post('/eDatacollect',
            ['json' =>
                [
                    'dataInizioAC' => $dataInizioAC->format( 'Y-m-d' ),
                    'dataFineAC' => $dataFineAC->format( 'Y-m-d' ),
                    'dataInizioAP' => $dataInizioAP->format( 'Y-m-d' ),
                    'dataFineAP' => $dataFineAP->format( 'Y-m-d' ),
                    'function' => 'recuperaOreReparto'
                ]
            ]
        );

        $hourDepartment = [];
        if ($response->getStatusCode() == 200) {
            $hourDepartment = json_decode($response->getBody()->getContents(), true);
        }

        // creazione file excel
        // -------------------------------------------------------------------------------
        $style = new Style();

        // creazione del workbook
        $workBook = new Spreadsheet();
        $workBook->getDefaultStyle()->getFont()->setName( 'Arial' );
        $workBook->getDefaultStyle()->getFont()->setSize( 12 );
        $workBook->getDefaultStyle()->getAlignment()->setVertical( Alignment::VERTICAL_CENTER );
        $workBook->getProperties()
            ->setCreator( "IF65 S.p.A. (Gruppo Italmark)" )
            ->setLastModifiedBy( "IF65 S.p.A." )
            ->setTitle( "report VENDITE" )
            ->setSubject( "report incassi" )
            ->setDescription( "report incassi" )
            ->setKeywords( "office 2007 openxml php" )
            ->setCategory( "IF65 Docs" );

        $integerFormat = '###,###,##0;[Red][<0]-###,###,##0;###,###,##0';
        $currencyFormat = '###,###,##0.00;[Red][<0]-###,###,##0.00;###,###,##0.00';
        $percentageFormat = '0.00%;[Red][<0]-0.00%;0.00%';

        $styleBorderArray = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];

        $debugList = ['0101' => 'PIAVE', '0102' => 'MARCHETTI']; //$shopList
        $firstSheet = true;
        foreach ($debugList as $sedeSelezionata => $sedeSelezionataDescrizione) {
            echo "$sedeSelezionata\n";
            // recupero numero di scontrini emessi dell'anno in corso e del precedente
            // -------------------------------------------------------------------------------
            $stmt = "   select s.store, s.ddate, count(distinct reg, trans) transCount
                    from mtx.sales as s join mtx.sottoreparto as r on s.`articledepartment`=r.`idsottoreparto` 
                    where s.ddate >= :dataInizio and s.ddate <= :dataFine and s.store = :sede
                    group by 1,2
                    order by 1,2;";

            $h_query = $db->prepare( $stmt );

            $h_query->execute( [':dataInizio' => $dataInizioAC->format( 'Y-m-d' ), ':dataFine' => $dataFineAC->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
            $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
            $transazioniAC = [];
            foreach ($result as $row) {
                foreach ($departmentList as $department) {
                    $index = $row['store'] . $row['ddate'] . $department;
                    $transazioniAC[$index] = $row['transCount'];
                }
            }

            $h_query->execute( [':dataInizio' => $dataInizioAP->format( 'Y-m-d' ), ':dataFine' => $dataFineAP->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
            $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
            $transazioniAP = [];
            foreach ($result as $row) {
                foreach ($departmentList as $department) {
                    $index = $row['store'] . $row['ddate'] . $department;
                    $transazioniAP[$index] = $row['transCount'];
                }
            }

            // recupero penetrazione dell'anno in corso e del precedente per reparto
            // -------------------------------------------------------------------------------
            $stmt = "   select s.store, s.ddate, r.nuovoReparto, concat(s.store, s.ddate, r.nuovoReparto) `index`,count(distinct reg, trans) transCount 
                    from mtx.sales as s join mtx.sottoreparto as r on s.`articledepartment`=r.`idsottoreparto` 
                    where s.ddate >= :dataInizio and s.ddate <= :dataFine and s.store = :sede
                    group by 1,2,3
                    order by 3,1,2;";

            $h_query = $db->prepare( $stmt );

            $h_query->execute( [':dataInizio' => $dataInizioAC->format( 'Y-m-d' ), ':dataFine' => $dataFineAC->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
            $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
            $penetrazioneAC = [];
            foreach ($result as $row) {
                $penetrazioneAC[$row['index']] = $row['transCount'];
            }

            $h_query->execute( [':dataInizio' => $dataInizioAP->format( 'Y-m-d' ), ':dataFine' => $dataFineAP->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
            $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
            $penetrazioneAP = [];
            foreach ($result as $row) {
                $penetrazioneAP[$row['index']] = $row['transCount'];
            }

            // recupero penetrazione dell'anno in corso e del precedente per subtotale
            // -------------------------------------------------------------------------------
            $stmt = "   select s.store, s.ddate, r.subtotali, count(distinct reg, trans) transCount 
                    from mtx.sales as s join mtx.sottoreparto as r on s.`articledepartment`=r.`idsottoreparto` 
                    where s.ddate >= :dataInizio and s.ddate <= :dataFine and s.store = :sede
                    group by 1,2,3
                    order by 3,1,2;";

            $h_query = $db->prepare( $stmt );

            $h_query->execute( [':dataInizio' => $dataInizioAC->format( 'Y-m-d' ), ':dataFine' => $dataFineAC->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
            $penetrazioneSubtotaleAC = $h_query->fetchAll( PDO::FETCH_ASSOC );

            $h_query->execute( [':dataInizio' => $dataInizioAP->format( 'Y-m-d' ), ':dataFine' => $dataFineAP->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
            $penetrazioneSubtotaleAP = $h_query->fetchAll( PDO::FETCH_ASSOC );

            // recupero dati del periodo corrente dell'anno in corso e del precedente
            // -------------------------------------------------------------------------------
            $stmt = "   select s.store, s.ddate, r.nuovoReparto department, sum(s.rowcount) rowcount, sum(s.quantity) quantity, sum(s.totaltaxableamount) totaltaxableamount
                    from mtx.sales as s join mtx.sottoreparto as r on s.`articledepartment`=r.`idsottoreparto` 
                    where s.ddate >= :dataInizio and s.ddate <= :dataFine and s.store = :sede
                    group by 1,2,3
                    order by 1,2,3;";

            $h_query = $db->prepare( $stmt );

            $h_query->execute( [':dataInizio' => $dataInizioAC->format( 'Y-m-d' ), ':dataFine' => $dataFineAC->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
            $periodoAC = $h_query->fetchAll( PDO::FETCH_ASSOC );

            $h_query->execute( [':dataInizio' => $dataInizioAP->format( 'Y-m-d' ), ':dataFine' => $dataFineAP->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
            $periodoAP = $h_query->fetchAll( PDO::FETCH_ASSOC );


            // arricchimento dati
            // -------------------------------------------------------------------------------
            $tempPeriodoAC = [];
            foreach ($periodoAC as $giorno) {
                $index = $giorno['store'] . $giorno['ddate'] . $giorno['department'];

                $giorno['transazioni'] = (key_exists( $index, $transazioniAC )) ? $transazioniAC[$index] : 0;
                $giorno['penetrazione'] = (key_exists( $index, $penetrazioneAC )) ? $penetrazioneAC[$index] : 0;
                $giorno['ore'] = (key_exists( $index, $hourDepartment )) ? $hourDepartment[$index] : 0;

                $tempPeriodoAC[$index] = $giorno;
            }

            $tempPeriodoAP = [];
            foreach ($periodoAP as $giorno) {
                $index = $giorno['store'] . $giorno['ddate'] . $giorno['department'];

                $giorno['transazioni'] = (key_exists( $index, $transazioniAP )) ? $transazioniAP[$index] : 0;
                $giorno['penetrazione'] = (key_exists( $index, $penetrazioneAP )) ? $penetrazioneAP[$index] : 0;
                $giorno['ore'] = (key_exists( $index, $hourDepartment )) ? $hourDepartment[$index] : 0;

                $tempPeriodoAP[$index] = $giorno;
            }
            $periodoAC = $tempPeriodoAC;
            $periodoAP = $tempPeriodoAP;

            // creazione del primo sheet
            // -------------------------------------------------------------------------------
            if ($firstSheet) {
                $sheet = $workBook->setActiveSheetIndex( 0 ); // la numerazione dei worksheet parte da 0
                $firstSheet = false;
            } else {
                $sheet =$workBook->createSheet();
            }
            $sheet->setTitle( "$sedeSelezionata");

            $sheet->getDefaultRowDimension()->setRowHeight( 32 );
            $sheet->getDefaultColumnDimension()->setWidth( 12 );

            // riclassificazione dati e creazione primo sheet
            // -------------------------------------------------------------------------------
            $periodoRiclassificatoAC = [];
            $periodoRiclassificatoAP = [];
            $penetrazioneSubtotaleRiclassificataAC = [];
            $penetrazioneSubtotaleRiclassificataAP = [];

            //riclassificazioneDati( $dataCorrenteAC, $dataCorrenteAC, $dataCorrenteAP, $dataCorrenteAP );
	        riclassificazioneDati( $dataInizioAC, $dataFineAC, $dataInizioAP, $dataFineAP );

            $originY = 1;
            $originX = 1;
            $spaziatura = "          ";
            $message = $spaziatura . "SEDE: $sedeSelezionata - $sedeSelezionataDescrizione$spaziatura";
            $message .= $spaziatura . "GIORNATA CORRENTE: " . $dataCorrenteAC->format( 'Y-m-d' );
            $message .= $spaziatura . "GIORNATA PRECEDENTE: " . $dataCorrenteAP->format( 'Y-m-d' );
            writeTitleRow( $message );
            writeDepartmentRows();
            writeSubtotalRows();
            writeTotalRow();

            $sheet->getStyle( RXY($originX + 1, $originY + 1, $originX + 3, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 4, $originY + 1, $originX + 6, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 7, $originY + 1, $originX + 9, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 10, $originY + 1, $originX + 12, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 13, $originY + 1, $originX + 15, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 16, $originY + 1, $originX + 17, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 19, $originY + 1, $originX + 21, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 22, $originY + 1, $originX + 24, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 25, $originY + 1, $originX + 27, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 28, $originY + 1, $originX + 30, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );

            // riclassificazione dati e creazione secondo sheet
            // -------------------------------------------------------------------------------
            $periodoRiclassificatoAC = [];
            $periodoRiclassificatoAP = [];
            $penetrazioneSubtotaleRiclassificataAC = [];
            $penetrazioneSubtotaleRiclassificataAP = [];

            riclassificazioneDati( $dataInizioAC, $dataFineAC, $dataInizioAP, $dataFineAP );

            $originY = 19;
            $sheet->setBreak(XY($originX, $originY - 1), Worksheet::BREAK_ROW);

            $spaziatura = "          ";
            $message = $spaziatura . "SEDE: $sedeSelezionata - $sedeSelezionataDescrizione$spaziatura";
            $message .= $spaziatura . "SETTIMANA CORRENTE: " . $dataInizioAC->format( 'Y-m-d' ) . ' => ' . $dataFineAC->format( 'Y-m-d' );
            $message .= $spaziatura . "SETTIMANA PRECEDENTE: " . $dataInizioAP->format( 'Y-m-d' ) . ' => ' . $dataFineAP->format( 'Y-m-d' );
            writeTitleRow( $message );
            writeDepartmentRows();
            writeSubtotalRows();
            writeTotalRow();

            $test = RXY($originX , $originY + 1, $originX , $originY + $subtotalCount + $departmentCount + 1);

            $sheet->getStyle( RXY($originX , $originY + 1, $originX , $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 1, $originY + 1, $originX + 3, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 4, $originY + 1, $originX + 6, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 7, $originY + 1, $originX + 9, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 10, $originY + 1, $originX + 12, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 13, $originY + 1, $originX + 15, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 16, $originY + 1, $originX + 17, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 19, $originY + 1, $originX + 21, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 22, $originY + 1, $originX + 24, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 25, $originY + 1, $originX + 27, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );
            $sheet->getStyle( RXY($originX + 28, $originY + 1, $originX + 30, $originY + $subtotalCount + $departmentCount + 1))->applyFromArray( $styleBorderArray );

            // formattazioni varie
            $sheet->getColumnDimensionByColumn( 1 )->setWidth( 30.0 );

            $sheet->getColumnDimensionByColumn( 3 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 6 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 9 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 12 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 15 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 18 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 21 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 24 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 27 )->setVisible(False);
            $sheet->getColumnDimensionByColumn( 30 )->setVisible(False);

            $sheet->setSelectedCell('A1');

            $sheet->getPageSetup()->setFitToWidth(1);
            $sheet->getPageSetup()->setFitToHeight(0);
            $sheet->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
            $sheet->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A4);
        }

    }

    $writer = new Xlsx( $workBook );
    $writer->save( '/Users/if65/Desktop/if65.xlsx' );

} catch (PDOException $e) {
    echo "Errore: " . $e->getMessage();
    die();
}


function riclassificazioneDati(DateTime $dataInizioPeriodoAC, DateTime $dataFinePeriodoAC, DateTime $dataInizioPeriodoAP, DateTime $dataFinePeriodoAP) {
    global $periodoAC, $periodoAP;
    global $periodoRiclassificatoAC, $periodoRiclassificatoAP;
    global $penetrazioneSubtotaleAC, $penetrazioneSubtotaleAP;
    global $penetrazioneSubtotaleRiclassificataAC, $penetrazioneSubtotaleRiclassificataAP;

    $dataInizioPeriodo = $dataInizioPeriodoAC;
    $dataFinePeriodo = $dataFinePeriodoAC;
    $periodoRiclassificatoAC = [];
    foreach ($periodoAC as $giorno) {
        $data = new DateTime($giorno['ddate']);
        if ($data <= $dataFinePeriodo and $data >= $dataInizioPeriodo) {
            if (key_exists($giorno['department'], $periodoRiclassificatoAC)) {
                $rowcount = $periodoRiclassificatoAC[$giorno['department']]['rowcount'] + $giorno['rowcount'];
                $periodoRiclassificatoAC[$giorno['department']]['rowcount'] = $rowcount;

                $quantity = $periodoRiclassificatoAC[$giorno['department']]['quantity'] + $giorno['quantity'];
                $periodoRiclassificatoAC[$giorno['department']]['quantity'] = $quantity;

                $totaltaxableamount = $periodoRiclassificatoAC[$giorno['department']]['totaltaxableamount'] + $giorno['totaltaxableamount'];
                $periodoRiclassificatoAC[$giorno['department']]['totaltaxableamount'] = $totaltaxableamount;

                $transazioni = $periodoRiclassificatoAC[$giorno['department']]['transazioni'] + $giorno['transazioni'];
                $periodoRiclassificatoAC[$giorno['department']]['transazioni'] = $transazioni;

                $penetrazione = $periodoRiclassificatoAC[$giorno['department']]['penetrazione'] + $giorno['penetrazione'];
                $periodoRiclassificatoAC[$giorno['department']]['penetrazione'] = $penetrazione;

                $ore = $periodoRiclassificatoAC[$giorno['department']]['ore'] + $giorno['ore'];
                $periodoRiclassificatoAC[$giorno['department']]['ore'] = $ore;
            } else {
                $periodoRiclassificatoAC[$giorno['department']] = [
                    'store' => $giorno['store'],
                    'ddate' => $giorno['ddate'],
                    'department' => $giorno['department'],
                    'rowcount' => $giorno['rowcount'],
                    'quantity' => $giorno['quantity'],
                    'totaltaxableamount' => $giorno['totaltaxableamount'],
                    'transazioni' => $giorno['transazioni'],
                    'penetrazione' => $giorno['penetrazione'],
                    'ore' => $giorno['ore']
                ];
            }
        }
    }

    $penetrazioneSubtotaleRiclassificataAC = [];
    foreach($penetrazioneSubtotaleAC as $giorno) {
        $data = new DateTime($giorno['ddate']);
        if ($data <= $dataFinePeriodo and $data >= $dataInizioPeriodo) {
            if (key_exists($giorno['subtotali'], $penetrazioneSubtotaleRiclassificataAC)) {
                $penetrazione = $penetrazioneSubtotaleRiclassificataAC[$giorno['subtotali']] + $giorno['transCount'];
                $penetrazioneSubtotaleRiclassificataAC[$giorno['subtotali']] = $penetrazione;
            } else {
                $penetrazioneSubtotaleRiclassificataAC[$giorno['subtotali']] = $giorno['transCount'];
            }
        }
    }

    $dataInizioPeriodo = $dataInizioPeriodoAP;
    $dataFinePeriodo = $dataFinePeriodoAP;
    $periodoRiclassificatoAP = [];
    foreach ($periodoAP as $giorno) {
        $data = new DateTime($giorno['ddate']);
        if ($data <= $dataFinePeriodo and $data >= $dataInizioPeriodo) {
            if (key_exists($giorno['department'], $periodoRiclassificatoAP)) {
                $rowcount = $periodoRiclassificatoAP[$giorno['department']]['rowcount'] + $giorno['rowcount'];
                $periodoRiclassificatoAP[$giorno['department']]['rowcount'] = $rowcount;

                $quantity = $periodoRiclassificatoAP[$giorno['department']]['quantity'] + $giorno['quantity'];
                $periodoRiclassificatoAP[$giorno['department']]['quantity'] = $quantity;

                $totaltaxableamount = $periodoRiclassificatoAP[$giorno['department']]['totaltaxableamount'] + $giorno['totaltaxableamount'];
                $periodoRiclassificatoAP[$giorno['department']]['totaltaxableamount'] = $totaltaxableamount;

                $transazioni = $periodoRiclassificatoAP[$giorno['department']]['transazioni'] + $giorno['transazioni'];
                $periodoRiclassificatoAP[$giorno['department']]['transazioni'] = $transazioni;

                $penetrazione = $periodoRiclassificatoAP[$giorno['department']]['penetrazione'] + $giorno['penetrazione'];
                $periodoRiclassificatoAP[$giorno['department']]['penetrazione'] = $penetrazione;

                $ore = $periodoRiclassificatoAP[$giorno['department']]['ore'] + $giorno['ore'];
                $periodoRiclassificatoAP[$giorno['department']]['ore'] = $ore;
            } else {
                $periodoRiclassificatoAP[$giorno['department']] = [
                    'store' => $giorno['store'],
                    'ddate' => $giorno['ddate'],
                    'department' => $giorno['department'],
                    'rowcount' => $giorno['rowcount'],
                    'quantity' => $giorno['quantity'],
                    'totaltaxableamount' => $giorno['totaltaxableamount'],
                    'transazioni' => $giorno['transazioni'],
                    'penetrazione' => $giorno['penetrazione'],
                    'ore' => $giorno['ore']
                ];
            }
        }
    }

    $penetrazioneSubtotaleRiclassificataAP = [];
    foreach($penetrazioneSubtotaleAP as $giorno) {
        $data = new DateTime($giorno['ddate']);
        if ($data <= $dataFinePeriodo and $data >= $dataInizioPeriodo) {
            if (key_exists($giorno['subtotali'], $penetrazioneSubtotaleRiclassificataAP)) {
                $penetrazione = $penetrazioneSubtotaleRiclassificataAP[$giorno['subtotali']] + $giorno['transCount'];
                $penetrazioneSubtotaleRiclassificataAP[$giorno['subtotali']] = $penetrazione;
            } else {
                $penetrazioneSubtotaleRiclassificataAP[$giorno['subtotali']] = $giorno['transCount'];
            }
        }
    }
}

function writeTitleRow(String $message = '') {
    global $sheet;
    global $originX, $originY;
    global $currentRow;
    global $styleBorderArray;

    $currentRow = $originY + 1;
    $currentColumn = $originX;
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, 'Reparto', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Vendite A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Vendite A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '% su A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Mix % Vendite A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Mix % Vendite A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta % Mix', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Clienti A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Clienti A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta Clienti %', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Penetrazione % A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Penetrazione % A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta % Penetrazione', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Scontrino Medio A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Scontrino Medio A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta Scontrino Medio %', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Pezzi A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Pezzi A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta Pezzi %', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Prezzo Medio A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Prezzo Medio A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta Prezzo Medio %', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Pezzi per Cliente A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Pezzi per Cliente A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta Pezzi per Cliente %', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Ore A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Ore A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta Ore %', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Procapite A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Procapite A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta Procapite %', DataType::TYPE_STRING );

    $titleStyle = [
        'font' => [
            'bold' => true,
            'color' => [
                'argb' => 'FF000000',
            ]
        ],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'startColor' => [
                'argb' => 'FFFFD700',
            ]
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER
        ]
    ];

    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->applyFromArray($titleStyle);
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow))->applyFromArray( $styleBorderArray );
    $sheet->getRowDimension( $currentRow )->setRowHeight( 48 );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->getAlignment()->setWrapText( true );
    //$sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->getFill()->setFillType( Fill::FILL_SOLID )->getStartColor()->setARGB( 'FFDCDCDC' );
    $currentRow++;

    // riquadro di intestazione
    $sheet->getStyle( RXY($originX, $originY, $currentColumn, $originY) )->applyFromArray( $titleStyle );
    $sheet->getStyle( RXY($originX, $originY, $currentColumn, $originY))->applyFromArray( $styleBorderArray );
    //$sheet->getStyle( RXY($originX, $originY, $currentColumn, $originY) )->getFont()->getColor()->setARGB(Color::COLOR_RED);
    $sheet->mergeCells(RXY($originX, $originY, $currentColumn, $originY));
    $sheet->setCellValueExplicitByColumnAndRow($originX, $originY, $message, DataType::TYPE_STRING);
    $sheet->getRowDimension( $originY )->setRowHeight( 48 );
}

function writeDepartmentRows()
{
    global $sheet;
    global $departmentCount;
    global $departmentList;
    global $currentRow;
    global $subtotalCount;
    global $integerFormat, $currencyFormat, $percentageFormat;
    global $originX, $originY;
    global $periodoRiclassificatoAC, $periodoRiclassificatoAP;

    foreach ($departmentList as $department) {
        $indexAC = $department;
        $indexAP = $department;

        $incassoAC = key_exists( $indexAC, $periodoRiclassificatoAC ) ? $periodoRiclassificatoAC[$indexAC]['totaltaxableamount'] : 0;
        $incassoAP = key_exists( $indexAP, $periodoRiclassificatoAP ) ? $periodoRiclassificatoAP[$indexAP]['totaltaxableamount'] : 0;
        $penetrazioneAC = key_exists( $indexAC, $periodoRiclassificatoAC ) ? $periodoRiclassificatoAC[$indexAC]['penetrazione'] : 0;
        $penetrazioneAP = key_exists( $indexAP, $periodoRiclassificatoAP ) ? $periodoRiclassificatoAP[$indexAP]['penetrazione'] : 0;
        $quantityAC = key_exists( $indexAC, $periodoRiclassificatoAC ) ? $periodoRiclassificatoAC[$indexAC]['quantity'] : 0;
        $quantityAP = key_exists( $indexAP, $periodoRiclassificatoAP ) ? $periodoRiclassificatoAP[$indexAP]['quantity'] : 0;
        $oreAC = key_exists( $indexAC, $periodoRiclassificatoAC ) ? $periodoRiclassificatoAC[$indexAP]['ore'] : 0;
        $oreAP = key_exists( $indexAP, $periodoRiclassificatoAP ) ? $periodoRiclassificatoAP[$indexAP]['ore'] : 0;

        // Reparto
        $currentColumn = $originX;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $department, DataType::TYPE_STRING );

        // Vendite A.C.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $incassoAC, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $incassoAC_XY = XY( $currentColumn, $currentRow );
        $incassoAC_TOT_XY = XY( $currentColumn, $originY + $departmentCount + $subtotalCount + 1 );

        // Vendite A.P.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $incassoAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $incassoAP_XY = XY( $currentColumn, $currentRow );
        $incassoAP_TOT_XY = XY( $currentColumn, $originY + $departmentCount + $subtotalCount + 1 );

        // % su A.P.
        $currentColumn++;
        $formula = "=IF($incassoAP_XY<>0, ($incassoAC_XY-$incassoAP_XY)/$incassoAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Mix % Vendite A.C.
        $currentColumn++;
        $formula = "=IF($incassoAC_TOT_XY<>0, $incassoAC_XY/$incassoAC_TOT_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $mixIncassoAC_XY = XY( $currentColumn, $currentRow );

        // Mix % Vendite A.P.
        $currentColumn++;
        $formula = "=IF($incassoAP_TOT_XY<>0, $incassoAP_XY/$incassoAP_TOT_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $mixIncassoAP_XY = XY( $currentColumn, $currentRow );

        // Delta % Mix
        $currentColumn++;
        $formula = "=$mixIncassoAC_XY - $mixIncassoAP_XY";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Clienti A.C.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $penetrazioneAC, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $penetrazioneAC_XY = XY( $currentColumn, $currentRow );
        $clientiAC_XY = XY( $currentColumn, $originY + $departmentCount + $subtotalCount + 1);

        // Clienti A.P.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $penetrazioneAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $penetrazioneAP_XY = XY( $currentColumn, $currentRow );
        $clientiAP_XY = XY( $currentColumn, $originY + $departmentCount + $subtotalCount + 1);

        // Delta Clienti %
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, ($penetrazioneAC_XY - $penetrazioneAP_XY)/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Penetrazione % A.C.
        $currentColumn++;
        $formula = "=IF($clientiAC_XY<>0, $penetrazioneAC_XY/$clientiAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $penetrazioneAC_P_XY = XY( $currentColumn, $currentRow );

        // Penetrazione % A.P.
        $currentColumn++;
        $formula = "=IF($clientiAP_XY<>0, $penetrazioneAP_XY/$clientiAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $penetrazioneAP_P_XY = XY( $currentColumn, $currentRow );

        // Delta % Penetrazione
        $currentColumn++;
        $formula = "=$penetrazioneAC_P_XY - $penetrazioneAP_P_XY";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Scontrino Medio A.C.
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, $incassoAC_XY/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $scontrinoMedioAC_XY = XY( $currentColumn, $currentRow );

        // Scontrino Medio A.P.
        $currentColumn++;
        $formula = "=IF($penetrazioneAP_XY<>0, $incassoAP_XY/$penetrazioneAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $scontrinoMedioAP_XY = XY( $currentColumn, $currentRow );

        // Delta Scontrino Medio %
        $currentColumn++;
        $formula = "=IF($scontrinoMedioAC_XY<>0, ($scontrinoMedioAC_XY - $scontrinoMedioAP_XY)/$scontrinoMedioAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Pezzi A.C.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $quantityAC, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $pezziAC_XY = XY( $currentColumn, $currentRow );

        // Pezzi A.P.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $quantityAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $pezziAP_XY = XY( $currentColumn, $currentRow );

        // Delta Pezzi %
        $currentColumn++;
        $formula = "=IF($pezziAC_XY<>0, ($pezziAC_XY - $pezziAP_XY)/$pezziAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Prezzo Medio A.C.
        $currentColumn++;
        $formula = "=IF($pezziAC_XY<>0, $incassoAC_XY/$pezziAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $prezzoMedioAC_XY = XY( $currentColumn, $currentRow );

        // Prezzo Medio A.P.
        $currentColumn++;
        $formula = "=IF($pezziAP_XY<>0, $incassoAP_XY/$pezziAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $prezzoMedioAP_XY = XY( $currentColumn, $currentRow );

        // Delta Prezzo Medio %
        $currentColumn++;
        $formula = "=IF($prezzoMedioAC_XY<>0, ($prezzoMedioAC_XY - $prezzoMedioAP_XY)/$prezzoMedioAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Pezzi per cliente A.C.
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, $pezziAC_XY/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $pezziClienteAC_XY = XY( $currentColumn, $currentRow );

        // Pezzi per cliente A.P.
        $currentColumn++;
        $formula = "=IF($penetrazioneAP_XY<>0, $pezziAP_XY/$penetrazioneAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $pezziClienteAP_XY = XY( $currentColumn, $currentRow );

        // Delta Pezzi per Cliente %
        $currentColumn++;
        $formula = "=IF($pezziClienteAC_XY<>0, ($pezziClienteAC_XY - $pezziClienteAP_XY)/$pezziClienteAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Ore A.C.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $oreAC, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $oreAC_XY = XY( $currentColumn, $currentRow );

        // Ore A.P.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $oreAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $oreAP_XY = XY( $currentColumn, $currentRow );

        // Delta Ore %
        $currentColumn++;
        $formula = "=IF($oreAC_XY<>0, ($oreAC_XY - $oreAP_XY)/$oreAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Procapite A.C.
        $currentColumn++;
        $formula = "=IF($oreAC_XY<>0, $incassoAC_XY/$oreAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $procapiteAC_XY = XY( $currentColumn, $currentRow );

        // Procapite A.P.
        $currentColumn++;
        $formula = "=IF($oreAP_XY<>0, $incassoAP_XY/$oreAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $procapiteAP_XY = XY( $currentColumn, $currentRow );

        // Delta Procapite %
        $currentColumn++;
        $formula = "=IF($procapiteAC_XY<>0, ($procapiteAC_XY - $procapiteAP_XY)/$procapiteAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        $sheet->getRowDimension( $currentRow )->setRowHeight( 32 );

        $currentRow++;
    }
}

function writeSubtotalRows() {
    global $currentRow;
    global $sheet;
    global $originX, $originY;
    global $departmentCount;
    global $subtotalTransCode;
    global $subtotalDescription;
    global $subtotalCount;
    global $integerFormat,$currencyFormat,$percentageFormat;
    global $penetrazioneSubtotaleRiclassificataAC, $penetrazioneSubtotaleRiclassificataAP;

    $subtotalRows = [];
    for ($i = $currentRow - $departmentCount; $i < $currentRow; $i++) {
        $reparto = $sheet->getCellByColumnAndRow($originX, $i)->getValue();
        if (! key_exists($subtotalTransCode[$reparto], $subtotalRows)) {
            $subtotalRows[$subtotalTransCode[$reparto]] = [$i];
        } else {
            $rows = $subtotalRows[$subtotalTransCode[$reparto]];
            $rows[] = $i;
            $subtotalRows[$subtotalTransCode[$reparto]] = $rows;
        }
    }
    foreach ($subtotalDescription as $description) {
        $currentColumn = $originX;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, 'TOTALE ' . $description, DataType::TYPE_STRING );

        $penetrazioneAC = key_exists( $description, $penetrazioneSubtotaleRiclassificataAC ) ? $penetrazioneSubtotaleRiclassificataAC[$description] : 0;
        $penetrazioneAP = key_exists( $description, $penetrazioneSubtotaleRiclassificataAP ) ? $penetrazioneSubtotaleRiclassificataAP[$description] : 0;
        $quantity = 0;

        // Vendite A.C.
        $currentColumn++;
        $formula = "=SUM(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $incassoAC_XY = XY( $currentColumn, $currentRow );
        $incassoAC_TOT_XY = XY($currentColumn, $originY + $departmentCount + $subtotalCount + 1);

        // Vendite A.P.
        $currentColumn++;
        $formula = "=SUM(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $incassoAP_XY = XY( $currentColumn, $currentRow );
        $incassoAP_TOT_XY = XY($currentColumn, $originY + $departmentCount + $subtotalCount + 1);

        // % su A.P.
        $currentColumn++;
        $formula = "=IF($incassoAP_XY<>0, ($incassoAC_XY-$incassoAP_XY)/$incassoAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Mix % Vendite A.C.
        $currentColumn++;
        $formula = "=IF($incassoAC_TOT_XY<>0, $incassoAC_XY/$incassoAC_TOT_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $mixIncassoAC_XY = XY( $currentColumn, $currentRow );

        // Mix % Vendite A.P.
        $currentColumn++;
        $formula = "=IF($incassoAP_TOT_XY<>0, $incassoAP_XY/$incassoAP_TOT_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $mixIncassoAP_XY = XY( $currentColumn, $currentRow );

        // Delta % Mix Incasso
        $currentColumn++;
        $formula = "=$mixIncassoAC_XY - $mixIncassoAP_XY";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Clienti A.C.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $penetrazioneAC, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $penetrazioneAC_XY = XY( $currentColumn, $currentRow );
        $clientiAC_XY = XY( $currentColumn, $originY + $departmentCount + $subtotalCount + 1);

        // Clienti A.P.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $penetrazioneAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $penetrazioneAP_XY = XY( $currentColumn, $currentRow );
        $clientiAP_XY = XY( $currentColumn, $originY + $departmentCount + $subtotalCount + 1);

        // Delta Clienti %
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, ($penetrazioneAC_XY - $penetrazioneAP_XY)/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Penetrazione % A.C.
        $currentColumn++;
        $formula = "=IF($clientiAC_XY<>0, $penetrazioneAC_XY/$clientiAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $penetrazioneAC_P_XY = XY( $currentColumn, $currentRow );

        // Penetrazione % A.P.
        $currentColumn++;
        $formula = "=IF($clientiAP_XY<>0, $penetrazioneAP_XY/$clientiAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $penetrazioneAP_P_XY = XY( $currentColumn, $currentRow );

        // Delta % Penetrazione
        $currentColumn++;
        $formula = "=$penetrazioneAC_P_XY - $penetrazioneAP_P_XY";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Scontrino Medio A.C.
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, $incassoAC_XY/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $scontrinoMedioAC_XY = XY( $currentColumn, $currentRow );

        // Scontrino Medio A.P.
        $currentColumn++;
        $formula = "=IF($penetrazioneAP_XY<>0, $incassoAP_XY/$penetrazioneAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $scontrinoMedioAP_XY = XY( $currentColumn, $currentRow );

        // Delta Scontrino Medio %
        $currentColumn++;
        $formula = "=IF($scontrinoMedioAC_XY<>0, ($scontrinoMedioAC_XY - $scontrinoMedioAP_XY)/$scontrinoMedioAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Pezzi A.C.
        $currentColumn++;
        $formula = "=SUM(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $pezziAC_XY = XY( $currentColumn, $currentRow );

        // Pezzi A.P.
        $currentColumn++;
        $formula = "=SUM(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $pezziAP_XY = XY( $currentColumn, $currentRow );

        // Delta Pezzi %
        $currentColumn++;
        $formula = "=IF($pezziAC_XY<>0, ($pezziAC_XY - $pezziAP_XY)/$pezziAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Prezzo Medio A.C.
        $currentColumn++;
        $formula = "=IF($pezziAC_XY<>0, $incassoAC_XY/$pezziAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $prezzoMedioAC_XY = XY( $currentColumn, $currentRow );

        // Prezzo Medio A.P.
        $currentColumn++;
        $formula = "=IF($pezziAP_XY<>0, $incassoAP_XY/$pezziAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $prezzoMedioAP_XY = XY( $currentColumn, $currentRow );

        // Delta Prezzo Medio %
        $currentColumn++;
        $formula = "=IF($prezzoMedioAC_XY<>0, ($prezzoMedioAC_XY - $prezzoMedioAP_XY)/$prezzoMedioAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Pezzi per cliente A.C.
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, $pezziAC_XY/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $pezziClienteAC_XY = XY( $currentColumn, $currentRow );

        // Pezzi per cliente A.P.
        $currentColumn++;
        $formula = "=IF($penetrazioneAP_XY<>0, $pezziAP_XY/$penetrazioneAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $pezziClienteAP_XY = XY( $currentColumn, $currentRow );

        // Delta Pezzi per Cliente %
        $currentColumn++;
        $formula = "=IF($pezziClienteAC_XY<>0, ($pezziClienteAC_XY - $pezziClienteAP_XY)/$pezziClienteAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Ore A.C.
        $currentColumn++;
        $formula = "=SUM(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $oreAC_XY = XY( $currentColumn, $currentRow );

        // Ore A.P.
        $currentColumn++;
        $formula = "=SUM(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $oreAP_XY = XY( $currentColumn, $currentRow );

        // Delta Ore %
        $currentColumn++;
        $formula = "=IF($oreAC_XY<>0, ($oreAC_XY - $oreAP_XY)/$oreAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Procapite A.C.
        $currentColumn++;
        $formula = "=IF($oreAC_XY<>0, $incassoAC_XY/$oreAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $procapiteAC_XY = XY( $currentColumn, $currentRow );

        // Procapite A.P.
        $currentColumn++;
        $formula = "=IF($oreAP_XY<>0, $incassoAP_XY/$oreAP_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
        $procapiteAP_XY = XY( $currentColumn, $currentRow );

        // Delta Procapite %
        $currentColumn++;
        $formula = "=IF($procapiteAC_XY<>0, ($procapiteAC_XY - $procapiteAP_XY)/$procapiteAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        $sheet->getRowDimension( $currentRow )->setRowHeight( 32 );

        $sheet->getStyle( RXY( 1, $currentRow, $currentColumn, $currentRow ) )->applyFromArray( ['font' => ['bold' => true, 'italic' => true]] );

        $currentRow++;
    }
}

function writeTotalRow() {
    global $currentRow;
    global $sheet;
    global $originX;
    global $departmentCount;
    global $integerFormat,$currencyFormat,$percentageFormat;
    global $subtotalCount;
    global $styleBorderArray;
    global $clientiAC;
    global $periodoRiclassificatoAC, $periodoRiclassificatoAP;

    $clientiAC = $periodoRiclassificatoAC[array_keys($periodoRiclassificatoAC)[0]]['transazioni'];
    $clientiAP = $periodoRiclassificatoAP[array_keys($periodoRiclassificatoAP)[0]]['transazioni'];


    $currentColumn = $originX;

    // Reparto
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, 'TOTALE', DataType::TYPE_STRING );

    // Vendite A.C.
    $currentColumn++;
    $formula = "=SUM(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $incassoAC_XY = XY( $currentColumn, $currentRow );

    // Vendite A.P.
    $currentColumn++;
    $formula = "=SUM(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount ) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $incassoAP_XY = XY( $currentColumn, $currentRow );

    // % su A.P.
    $currentColumn++;
    $formula = "=IF($incassoAP_XY<>0, ($incassoAC_XY-$incassoAP_XY)/$incassoAP_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

    // Mix % Vendite A.C.
    $currentColumn++;

    // Mix % Vendite A.P.
    $currentColumn++;

    // Delta % Mix
    $currentColumn++;

    // Clienti A.C.
    $currentColumn++;
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $clientiAC, DataType::TYPE_NUMERIC );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
    $clientiAC_XY = XY( $currentColumn, $currentRow );

    // Clienti A.P.
    $currentColumn++;
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $clientiAP, DataType::TYPE_NUMERIC );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
    $clientiAP_XY = XY( $currentColumn, $currentRow );

    // Delta Clienti %
    $currentColumn++;
    $formula = "=IF($clientiAC_XY<>0, ($clientiAC_XY - $clientiAP_XY)/$clientiAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

    // Penetrazione % A.C.
    $currentColumn++;

    // Penetrazione % A.P.
    $currentColumn++;

    // Delta % Penetrazione
    $currentColumn++;

    // Scontrino Medio A.C.
    $currentColumn++;
    $formula = "=IF($clientiAC_XY<>0, $incassoAC_XY/$clientiAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $scontrinoMedioAC_XY = XY( $currentColumn, $currentRow );

    // Scontrino Medio A.P.
    $currentColumn++;
    $formula = "=IF($clientiAP_XY<>0, $incassoAP_XY/$clientiAP_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $scontrinoMedioAP_XY = XY( $currentColumn, $currentRow );

    // Delta Scontrino Medio %
    $currentColumn++;
    $formula = "=IF($scontrinoMedioAC_XY<>0, ($scontrinoMedioAC_XY - $scontrinoMedioAP_XY)/$scontrinoMedioAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

    // Pezzi A.C.
    $currentColumn++;
    $formula = "=SUM(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount ) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
    $pezziAC_XY = XY( $currentColumn, $currentRow );

    // Pezzi A.P.
    $currentColumn++;
    $formula = "=SUM(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount ) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
    $pezziAP_XY = XY( $currentColumn, $currentRow );

    // Delta Pezzi %
    $currentColumn++;
    $formula = "=IF($pezziAC_XY<>0, ($pezziAC_XY - $pezziAP_XY)/$pezziAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

    // Prezzo Medio A.C.
    $currentColumn++;
    $formula = "=IF($pezziAC_XY<>0, $incassoAC_XY/$pezziAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $prezzoMedioAC_XY = XY( $currentColumn, $currentRow );

    // Prezzo Medio A.P.
    $currentColumn++;
    $formula = "=IF($pezziAP_XY<>0, $incassoAP_XY/$pezziAP_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $prezzoMedioAP_XY = XY( $currentColumn, $currentRow );

    // Delta Prezzo Medio %
    $currentColumn++;
    $formula = "=IF($prezzoMedioAC_XY<>0, ($prezzoMedioAC_XY - $prezzoMedioAP_XY)/$prezzoMedioAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

    // Pezzi per cliente A.C.
    $currentColumn++;
    $formula = "=IF($clientiAC_XY<>0, $pezziAC_XY/$clientiAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $pezziClienteAC_XY = XY( $currentColumn, $currentRow );

    // Pezzi per cliente A.P.
    $currentColumn++;
    $formula = "=IF($clientiAP_XY<>0, $pezziAP_XY/$clientiAP_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $pezziClienteAP_XY = XY( $currentColumn, $currentRow );

    // Delta Pezzi per Cliente %
    $currentColumn++;
    $formula = "=IF($pezziClienteAC_XY<>0, ($pezziClienteAC_XY - $pezziClienteAP_XY)/$pezziClienteAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

    // Ore A.C.
    $currentColumn++;
    $formula = "=SUM(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount ) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $oreAC_XY = XY( $currentColumn, $currentRow );

    // Ore A.P.
    $currentColumn++;
    $formula = "=SUM(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount ) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $oreAP_XY = XY( $currentColumn, $currentRow );

    // Delta Ore %
    $currentColumn++;
    $formula = "=IF($oreAC_XY<>0, ($oreAC_XY - $oreAP_XY)/$oreAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

    // Procapite A.C.
    $currentColumn++;
    $formula = "=IF($oreAC_XY<>0, $incassoAC_XY/$oreAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $procapiteAC_XY = XY( $currentColumn, $currentRow );

    // Procapite A.P.
    $currentColumn++;
    $formula = "=IF($oreAP_XY<>0, $incassoAP_XY/$oreAP_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );
    $procapiteAP_XY = XY( $currentColumn, $currentRow );

    // Delta Procapite %
    $currentColumn++;
    $formula = "=IF($procapiteAC_XY<>0, ($procapiteAC_XY - $procapiteAP_XY)/$procapiteAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->applyFromArray(['font' => ['bold' => true]] );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow))->applyFromArray( $styleBorderArray );
    $sheet->getRowDimension( $currentRow )->setRowHeight( 48 );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->getAlignment()->setWrapText( true );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->getFill()->setFillType( Fill::FILL_SOLID )->getStartColor()->setARGB( 'FFFFD700' );

}

function XY(int $x, int $y): string {
    return Coordinate::stringFromColumnIndex( $x ) . $y;
}

function RXY(int $x1, int $y1, int $x2, int $y2): string {
    return Coordinate::stringFromColumnIndex( $x1 ) . $y1. ':' . Coordinate::stringFromColumnIndex( $x2 ) . $y2;
}
