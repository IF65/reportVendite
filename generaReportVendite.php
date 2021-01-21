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
use PhpOffice\PhpSpreadsheet\Shared\Date;
// sede selezionata
$sedeSelezionata = '0102';

$timeZone = new DateTimeZone('Europe/Rome');

// inizio e fine settimana corrente
$dataCorrente = new DateTime('2021-01-16', $timeZone);

$dataInizioAC = new DateTime('2021-01-11', $timeZone);
$dataFineAC = new DateTime('2021-01-17', $timeZone);
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

        $sedeSelezionataDescrizione = $shopList[$sedeSelezionata];

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
        foreach($result as $row) {
            foreach ($departmentList as $department) {
                $index = $row['store'] . $row['ddate'] . $department;
                $transazioniAC[$index] = $row['transCount'];
            }
        }

        $h_query->execute( [':dataInizio' => $dataInizioAP->format( 'Y-m-d' ), ':dataFine' => $dataFineAP->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
        $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
        $transazioniAP = [];
        foreach($result as $row) {
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
        foreach($result as $row) {
            $penetrazioneAC[$row['index']] = $row['transCount'];
        }

        $h_query->execute( [':dataInizio' => $dataInizioAP->format( 'Y-m-d' ), ':dataFine' => $dataFineAP->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
        $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
        $penetrazioneAP = [];
        foreach($result as $row) {
            $penetrazioneAP[$row['index']] = $row['transCount'];
        }

        // recupero penetrazione dell'anno in corso e del precedente per subtotale
        // -------------------------------------------------------------------------------
        $stmt = "   select s.store, s.ddate, r.subtotali, concat(s.store, s.ddate, r.subtotali) `index`,count(distinct reg, trans) transCount 
                    from mtx.sales as s join mtx.sottoreparto as r on s.`articledepartment`=r.`idsottoreparto` 
                    where s.ddate >= :dataInizio and s.ddate <= :dataFine and s.store = :sede
                    group by 1,2,3
                    order by 3,1,2;";

        $h_query = $db->prepare( $stmt );

        $h_query->execute( [':dataInizio' => $dataInizioAC->format( 'Y-m-d' ), ':dataFine' => $dataFineAC->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
        $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
        $penetrazioneSubtotaleAC = [];
        foreach($result as $row) {
            $penetrazioneSubtotaleAC[$row['index']] = $row['transCount'];
        }

        $h_query->execute( [':dataInizio' => $dataInizioAP->format( 'Y-m-d' ), ':dataFine' => $dataFineAP->format( 'Y-m-d' ), ':sede' => $sedeSelezionata] );
        $result = $h_query->fetchAll( PDO::FETCH_ASSOC );
        $penetrazioneSubtotaleAP = [];
        foreach($result as $row) {
            $penetrazioneSubtotaleAP[$row['index']] = $row['transCount'];
        }

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


        // riclassificazione dati
        // -------------------------------------------------------------------------------
        $periodoRiclassificatoAC = [];
        foreach ($periodoAC as $giorno) {
            $index = $giorno['store'] . $giorno['ddate'] . $giorno['department'];

            $giorno['transazioni'] = (key_exists($index, $transazioniAC)) ? $transazioniAC[$index] : 0;
            $giorno['penetrazione'] = (key_exists($index, $penetrazioneAC)) ? $penetrazioneAC[$index] : 0;

            $periodoRiclassificatoAC[$index] = $giorno;
        }

        $periodoRiclassificatoAP = [];
        foreach ($periodoAP as $giorno) {
            $index = $giorno['store'] . $giorno['ddate'] . $giorno['department'];

            $giorno['transazioni'] = (key_exists($index, $transazioniAP)) ? $transazioniAP[$index] : 0;
            $giorno['penetrazione'] = (key_exists($index, $penetrazioneAP)) ? $penetrazioneAP[$index] : 0;

            $periodoRiclassificatoAP[$index] = $giorno;
        }
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

    // dichiarazione variabili
    // -------------------------------------------------------------------------------
    $originY = 1;
    $originX = 1;
    $currentRow = $originY;

    // creazione del primo sheet
    // -------------------------------------------------------------------------------
    $sheet = $workBook->setActiveSheetIndex( 0 ); // la numerazione dei worksheet parte da 0
    $sheet->setTitle( $sedeSelezionata );
    $sheet->getDefaultRowDimension()->setRowHeight( 32 );
    $sheet->getDefaultColumnDimension()->setWidth( 12 );

    // Giornata corrente
    // -------------------------------------------------------------------------------
    writeTitleRow();
    writeDepartmentRows();
    writeSubtotalRows();
    writeTotalRow();

    // formattazioni varie
    $sheet->getColumnDimensionByColumn(1)->setWidth(30.0);

    $writer = new Xlsx($workBook);
    $writer->save('/Users/if65/Desktop/if65.xlsx');

} catch (PDOException $e) {
    echo "Errore: " . $e->getMessage();
    die();
}

function writeTitleRow() {
    global $sheet;
    global $originX, $originY;
    global $currentRow;
    global $styleBorderArray;

    $currentRow = $originY;
    $currentColumn = $originX;
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, 'Reparto', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Vendite A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Vendite A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '% su A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Mix % Vendite A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Mix % Vendite A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta % Mix', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Clienti A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Penetrazione A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Penetrazione % A.C.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Clienti A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Penetrazione A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Penetrazione % A.P.', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Delta % Penetrazione', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Scontrino Medio', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Pezzi', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Prezzo Medio', DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Pezzi per Cliente', DataType::TYPE_STRING );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->applyFromArray(['font' => ['bold' => true], 'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]] );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow))->applyFromArray( $styleBorderArray );
    $sheet->getRowDimension( $currentRow )->setRowHeight( 48 );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->getAlignment()->setWrapText( true );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->getFill()->setFillType( Fill::FILL_SOLID )->getStartColor()->setARGB( 'FFDCDCDC' );
    $currentRow++;
}

function writeDepartmentRows()
{
    global $sheet;
    global $sedeSelezionata;
    global $dataCorrente;
    global $departmentCount;
    global $departmentList;
    global $currentRow;
    global $subtotalCount;
    global $integerFormat, $currencyFormat, $percentageFormat;
    global $originX, $originY;
    global $periodoRiclassificatoAC, $periodoRiclassificatoAP;

    foreach ($departmentList as $department) {
        $indexAC = $sedeSelezionata . $dataCorrente->format( 'Y-m-d' ) . $department;
        $indexAP = $sedeSelezionata . (clone $dataCorrente)->sub( new DateInterval( 'P364D' ) )->format( 'Y-m-d' ) . $department;

        $incassoAC = key_exists( $indexAC, $periodoRiclassificatoAC ) ? $periodoRiclassificatoAC[$indexAC]['totaltaxableamount'] : 0;
        $incassoAP = key_exists( $indexAP, $periodoRiclassificatoAP ) ? $periodoRiclassificatoAP[$indexAP]['totaltaxableamount'] : 0;
        $clientiAC = key_exists( $indexAC, $periodoRiclassificatoAC ) ? $periodoRiclassificatoAC[$indexAC]['transazioni'] : 0;
        $clientiAP = key_exists( $indexAP, $periodoRiclassificatoAP ) ? $periodoRiclassificatoAP[$indexAP]['transazioni'] : 0;
        $penetrazioneAC = key_exists( $indexAC, $periodoRiclassificatoAC ) ? $periodoRiclassificatoAC[$indexAC]['penetrazione'] : 0;
        $penetrazioneAP = key_exists( $indexAP, $periodoRiclassificatoAP ) ? $periodoRiclassificatoAP[$indexAP]['penetrazione'] : 0;
        $quantity = key_exists( $indexAC, $periodoRiclassificatoAC ) ? $periodoRiclassificatoAC[$indexAC]['quantity'] : 0;

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

        // Delta % Mix Incasso
        $currentColumn++;
        $formula = "=$mixIncassoAC_XY - $mixIncassoAP_XY";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );

        // Clienti A.C.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $clientiAC, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $clientiAC_XY = XY( $currentColumn, $currentRow );

        // Penetrazione A.C.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $penetrazioneAC, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $penetrazioneAC_XY = XY( $currentColumn, $currentRow );

        // Penetrazione % A.C.
        $currentColumn++;
        $formula = "=IF($clientiAC_XY<>0, $penetrazioneAC_XY/$clientiAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $penetrazioneAC_P_XY = XY( $currentColumn, $currentRow );

        // Clienti A.P.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $clientiAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $clientiAP_XY = XY( $currentColumn, $currentRow );

        // Penetrazione A.P.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $penetrazioneAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $penetrazioneAP_XY = XY( $currentColumn, $currentRow );

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

        // Scontrino Medio
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, $incassoAC_XY/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

        // Pezzi
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $quantity, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $pezzi_XY = XY( $currentColumn, $currentRow );

        // Prezzo Medio
        $currentColumn++;
        $formula = "=IF($pezzi_XY<>0, $incassoAC_XY/$pezzi_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

        // Pezzi per cliente
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, $pezzi_XY/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

        $sheet->getRowDimension( $currentRow )->setRowHeight( 32 );

        $currentRow++;
    }
}

function writeSubtotalRows() {
    global $currentRow;
    global $sedeSelezionata, $dataCorrente;
    global $sheet;
    global $originX, $originY;
    global $departmentCount;
    global $subtotalTransCode;
    global $subtotalDescription;
    global $subtotalCount;
    global $integerFormat,$currencyFormat,$percentageFormat;
    global $penetrazioneSubtotaleAC, $penetrazioneSubtotaleAP;

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

        $indexAC = $sedeSelezionata . $dataCorrente->format( 'Y-m-d' ) . $description;
        $indexAP = $sedeSelezionata . (clone $dataCorrente)->sub( new DateInterval( 'P364D' ) )->format( 'Y-m-d' ) . $description;
        $penetrazioneAC = key_exists( $indexAC, $penetrazioneSubtotaleAC ) ? $penetrazioneSubtotaleAC[$indexAC] : 0;
        $penetrazioneAP = key_exists( $indexAP, $penetrazioneSubtotaleAP ) ? $penetrazioneSubtotaleAP[$indexAP] : 0;
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
        $formula = "=MAX(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $clientiAC_XY = XY( $currentColumn, $currentRow );

        // Penetrazione A.C.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $penetrazioneAC, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $penetrazioneAC_XY = XY( $currentColumn, $currentRow );

        // Penetrazione % A.C.
        $currentColumn++;
        $formula = "=IF($clientiAC_XY<>0, $penetrazioneAC_XY/$clientiAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $percentageFormat );
        $penetrazioneAC_P_XY = XY( $currentColumn, $currentRow );

        // Clienti A.P.
        $currentColumn++;
        $formula = "=MAX(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $clientiAP_XY = XY( $currentColumn, $currentRow );
        $clientiAP_TOT_XY = XY($currentColumn, $originY + $departmentCount + $subtotalCount + 1);

        // Penetrazione A.P.
        $currentColumn++;
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $penetrazioneAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $penetrazioneAP_XY = XY( $currentColumn, $currentRow );

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

        // Scontrino Medio
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, $incassoAC_XY/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

        // Pezzi
        $currentColumn++;
        $formula = "=SUM(";
        foreach ($subtotalRows[$description] as $row) {
            $formula .= (XY( $currentColumn, $row ) . ',');
        }
        $formula = preg_replace( '/.$/', ')', $formula );
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
        $pezzi_XY = XY( $currentColumn, $currentRow );

        // Prezzo Medio
        $currentColumn++;
        $formula = "=IF($pezzi_XY<>0, $incassoAC_XY/$pezzi_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

        // Pezzi per cliente
        $currentColumn++;
        $formula = "=IF($penetrazioneAC_XY<>0, $pezzi_XY/$penetrazioneAC_XY,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

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
    $formula = "=MAX(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
    $clientiAC_XY = XY( $currentColumn, $currentRow );

    // Penetrazione A.C.
    $currentColumn++;

    // Penetrazione % A.C.
    $currentColumn++;

    // Clienti A.P.
    $currentColumn++;
    $formula = "=MAX(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
    $clientiAP_XY = XY( $currentColumn, $currentRow );

    // Penetrazione A.P.
    $currentColumn++;

    // Penetrazione % A.P.
    $currentColumn++;

    // Delta % Penetrazione
    $currentColumn++;

    // Scontrino Medio
    $currentColumn++;
    $formula = "=IF($clientiAC_XY<>0, $incassoAC_XY/$clientiAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

    // Pezzi
    $currentColumn++;
    $formula = "=SUM(" . RXY( $currentColumn, $currentRow - $departmentCount - $subtotalCount, $currentColumn, $currentRow - 1 - $subtotalCount) . ")";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $integerFormat );
    $pezzi_XY = XY( $currentColumn, $currentRow );

    // Prezzo Medio
    $currentColumn++;
    $formula = "=IF($pezzi_XY<>0, $incassoAC_XY/$pezzi_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

    // Pezzi per cliente
    $currentColumn++;
    $formula = "=IF($clientiAC_XY<>0, $pezzi_XY/$clientiAC_XY,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow( $currentColumn, $currentRow )->getNumberFormat()->setFormatCode( $currencyFormat );

    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->applyFromArray(['font' => ['bold' => true]] );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow))->applyFromArray( $styleBorderArray );
    $sheet->getRowDimension( $currentRow )->setRowHeight( 48 );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->getAlignment()->setWrapText( true );
    $sheet->getStyle( RXY($originX, $currentRow, $currentColumn, $currentRow) )->getFill()->setFillType( Fill::FILL_SOLID )->getStartColor()->setARGB( 'FFDCDCDC' );

}

function XY(int $x, int $y): string {
    return Coordinate::stringFromColumnIndex( $x ) . $y;
}

function RXY(int $x1, int $y1, int $x2, int $y2): string {
    return Coordinate::stringFromColumnIndex( $x1 ) . $y1. ':' . Coordinate::stringFromColumnIndex( $x2 ) . $y2;
}