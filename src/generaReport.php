<?php

require '../vendor/autoload.php';

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
use PhpOffice\PhpSpreadsheet\Calculation;

$timeZone = new DateTimeZone('Europe/Rome');

/**
 * date di inizio e fine settimana corrente (da parametrizzare su linea di comando)
 */
$messaggio = "Settimana 4";
$dataCorrenteAC = new DateTime('2021-05-30', $timeZone);
$dataInizioAC = new DateTime('2021-05-24', $timeZone);
$dataFineAC = new DateTime('2021-05-30', $timeZone);
$dataCorrenteAP = new DateTime('2019-05-31', $timeZone);
$dataInizioAP = new DateTime('2019-05-25', $timeZone);
$dataFineAP = new DateTime('2019-05-31', $timeZone);

/**
 * parametri per l'accesso all'host
*/
$hostname = '10.11.14.248';
$user = 'root';
$password = 'mela';


$db = new PDO("mysql:host=$hostname", $user, $password, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);

/**
 * recupero elenco negozi aperti nell'intervallo date prescelto e creo gli array assiociativi
 * necessari ai raggruppamenti per area e tipo negozio.
*/
$stmt = "   select n.codice code, n.negozio_descrizione description, n.area, n.tipoFood type 
            from archivi.negozi as n 
            where n.`data_inizio` <= :dataFine and (n.`data_fine` >= :dataInizio or n.`data_fine`is null) and 
                  n.`societa` in ('02','05') and n.`codice` not like '00%' 
            order by 1;";
$h_query = $db->prepare($stmt);
$h_query->execute([':dataInizio' => $dataInizioAC->format('Y-m-d'), ':dataFine' => $dataFineAC->format('Y-m-d')]);
$result = $h_query->fetchAll(PDO::FETCH_ASSOC);
$shopList = [];
foreach ($result as $shop) {
	$shopList[$shop['code']] = $shop['description'];
}
$shop2Area = [];
foreach ($result as $shop) {
	$shop2Area[$shop['code']] = $shop['area'];
}
$shop2Type = [];
foreach ($result as $shop) {
	$shop2Type[$shop['code']] = $shop['type'];
}

/**
 * recupero elenco reparti (nuovi)
 */
$stmt = "   select distinct 
                nuovoReparto department,
                subtotali subtotal
   			from mtx.sottoreparto
   			order by sortOrder";
$h_query = $db->prepare($stmt);
$h_query->execute();
$result = $h_query->fetchAll(PDO::FETCH_ASSOC);
$departmentList = [];
foreach ($result as $department) {
	$departmentList[] = $department['department'];
}
$department2Subtotal=[];
foreach ($result as $department) {
	$department2Subtotal[$department['department']] = $department['subtotal'];
}
$departmentCount = sizeof($departmentList);

/**
 * recupero elenco subtotali
 */
$stmt = "   select distinct 
                subtotali subtotal
   			from mtx.sottoreparto
   			order by sortOrderSubTotale";
$h_query = $db->prepare($stmt);
$h_query->execute();
$subtotalList = $h_query->fetchAll(PDO::FETCH_COLUMN);
$subtotalCount = sizeof($subtotalList);

/**
 * calcolo i dati per alimentare il report utilizzando il periodo prescelto e il
 * tipo report prescelto.
 */
$stmt="	select
			d.store, 
			d.department,
			r.sortOrder,
			r.subtotal,
			r.sortOrderSubtotal,
			n.area,
			n.tipoFood,
			n.negozio_descrizione,
			ifnull(sum(case when d.ddate >= :startCY and d.ddate <= :stopCY then d.totaltaxableamount else 0 end),0) totaltaxableamountCY,
			ifnull(sum(case when d.ddate >= :startLY and d.ddate <= :stopLY then d.totaltaxableamount else 0 end),0) totaltaxableamountLY,
			ifnull(sum(case when d.ddate >= :startCY and d.ddate <= :stopCY then d.rowCount else 0 end),0) rowCountCY,
			ifnull(sum(case when d.ddate >= :startLY and d.ddate <= :stopLY then d.rowCount else 0 end),0) rowCountLY,
			ifnull(sum(case when d.ddate >= :startCY and d.ddate <= :stopCY then d.quantity else 0 end),0) quantityCY,
			ifnull(sum(case when d.ddate >= :startLY and d.ddate <= :stopLY then d.quantity else 0 end),0) quantityLY,
			ifnull(sum(case when d.ddate >= :startCY and d.ddate <= :stopCY then d.customerCount else 0 end),0) customerCountPerDepartmentCY,
			ifnull(sum(case when d.ddate >= :startLY and d.ddate <= :stopLY then d.customerCount else 0 end),0) customerCountPerDepartmentLY,
			ifnull(sum(case when d.ddate >= :startCY and d.ddate <= :stopCY then p.customerCount else 0 end),0) customerCountPerSubtotalCY,
			ifnull(sum(case when d.ddate >= :startLY and d.ddate <= :stopLY then p.customerCount else 0 end),0) customerCountPerSubtotalLY,
			ifnull(sum(case when d.ddate >= :startCY and d.ddate <= :stopCY then c.customerCount else 0 end),0) customerCY,
			ifnull(sum(case when d.ddate >= :startLY and d.ddate <= :stopLY then c.customerCount else 0 end),0) customerLY
		from mtx.salesPerDepartment as d join (select distinct nuovoReparto department, sortOrder, subtotali subtotal, sortOrderSubTotale sortOrderSubtotal from mtx.sottoreparto ) as r on d.department = r.department join mtx.customers as c on c.`ddate`=d.`ddate` and c.`store`=d.`store` join mtx.penetrationPerSubtotal as p on p.`ddate`=d.`ddate` and p.`store`=d.`store` and p.subtotal=r.subtotal join archivi.negozi as n on d.store = n.codice 
		where ((d.ddate >= :startCY and d.ddate <= :stopCY) or (d.ddate >= :startLY and d.ddate <= :stopLY)) 
		group by 1,2
		order by store, sortOrderSubtotal, sortOrder;";
$h_query = $db->prepare($stmt);
$h_query->execute([
	':startCY' => $dataInizioAC->format( 'Y-m-d' ),
	':stopCY' => $dataFineAC->format( 'Y-m-d' ),
	':startLY' => $dataInizioAP->format( 'Y-m-d' ),
	':stopLY' => $dataFineAP->format( 'Y-m-d' )
]);
$result = $h_query->fetchAll(PDO::FETCH_ASSOC);

/**
 * per scrivere le pagine trasferisco i dati in un array associativo di pagina. Per esempio nel caso si tratti
 * della stampa per negozio trasferisco nell'array di pagina un negozio alla volta e poi passo il risultato alla
 * procedura di scrittura. In questo modo la procedura di scrittura rimane unica anche nel caso si stampino report
 * di area, di tipo negozio oppure dei totali.
 */

/** negozi */
$spreadsheet = new Spreadsheet();
$spreadsheet->getDefaultStyle()->getFont()->setName( 'Calibri' );
$spreadsheet->getDefaultStyle()->getFont()->setSize( 12 );
$spreadsheet->getDefaultStyle()->getAlignment()->setVertical( Alignment::VERTICAL_CENTER );

$spreadsheet->removeSheetByIndex(0);
foreach ($shopList as $shopCode => $shopDescription) {
	$sheet = [];
	foreach ($result as $row) {
		if ($row['store'] == $shopCode) {
			$sheet[] = $row;
		}
	}

	$worksheet = new Worksheet($spreadsheet, (string)$shopCode);
	$worksheet->getDefaultRowDimension()->setRowHeight( 24 );
	$worksheet->getDefaultColumnDimension()->setWidth( 12 );

	fillSheet($worksheet, $sheet);
	$spreadsheet->addSheet($worksheet);
}

$writer = new Xlsx($spreadsheet);
$writer->save('/Users/if65/Desktop/if65_0.xlsx');

function fillSheet(Worksheet &$worksheet, array $sheet)
{
	/*$text = json_encode($sheet, JSON_PRETTY_PRINT);
	file_put_contents('/Users/if65/Desktop/dati.json',$text);*/

	$subtotals = [];
	foreach ($sheet as $row) {
		if (! key_exists($row['subtotal'], $subtotals)) {
			$subtotals[$row['subtotal']] = ['departmentCount' => 1];
		} else {
			$subtotals[$row['subtotal']]['departmentCount'] += 1;
		}
	}

	/** determino le righe di inizio e fine dei dati di ognuno dei subtotali */
	$y = 5;
	foreach($subtotals as $subtotal => $void) {
		$subtotals[$subtotal]['start'] = $y;
		$subtotals[$subtotal]['stop'] = $y + $subtotals[$subtotal]['departmentCount'] - 1;
		$y = $subtotals[$subtotal]['stop'] + 4;
	}

	/** riga su cui vanno scritti i totali */
	$total_Y = $y - 1;

	/**
	 * TOTALI
	 * inserisco qui alcune posizioni di celle appartenenti ai totali perché serveno nelle formule
	 * dei reparti e dei sottototali.
	 */
	$incassoTotaleAC_XY = XY(2, $total_Y);
	$incassoTotaleAP_XY = XY(3, $total_Y);
	$clientiTotaleAC_XY = XY(8, $total_Y);
	$clientiTotaleAP_XY = XY(9, $total_Y);
	$pezziTotaleAC_XY = XY(17, $total_Y);
	$pezziTotaleAP_XY = XY(18, $total_Y);

	$customerCY = 0;
	$customerLY = 0;

	$y = 1;
	$worksheet->getRowDimension($y)->setRowHeight(48);

	$y = 2;
	$worksheet->getRowDimension($y)->setRowHeight(48);
	$worksheet->setCellValueExplicitByColumnAndRow(2, $y, 'Venduto', DataType::TYPE_STRING);
	$worksheet->mergeCellsByColumnAndRow(2, $y, 4, $y);
	$worksheet->setCellValueExplicitByColumnAndRow(5, $y, 'Mix Venduto', DataType::TYPE_STRING);
	$worksheet->mergeCellsByColumnAndRow(5, $y, 7, $y);
	$worksheet->setCellValueExplicitByColumnAndRow(8, $y, 'Clienti', DataType::TYPE_STRING);
	$worksheet->mergeCellsByColumnAndRow(8, $y, 10, $y);
	$worksheet->setCellValueExplicitByColumnAndRow(11, $y, 'Penetrazione', DataType::TYPE_STRING);
	$worksheet->mergeCellsByColumnAndRow(11, $y, 13, $y);
	$worksheet->setCellValueExplicitByColumnAndRow(14, $y, 'Scontrino Medio', DataType::TYPE_STRING);
	$worksheet->mergeCellsByColumnAndRow(14, $y, 16, $y);
	$worksheet->setCellValueExplicitByColumnAndRow(17, $y, 'Pezzi', DataType::TYPE_STRING);
	$worksheet->mergeCellsByColumnAndRow(17, $y, 19, $y);
	$worksheet->setCellValueExplicitByColumnAndRow(20, $y, 'Prezzo Medio', DataType::TYPE_STRING);
	$worksheet->mergeCellsByColumnAndRow(20, $y, 22, $y);
	$worksheet->setCellValueExplicitByColumnAndRow(23, $y, 'Pezzi per Cliente', DataType::TYPE_STRING);
	$worksheet->mergeCellsByColumnAndRow(23, $y, 25, $y);

	$y = 3;
	//$worksheet->getRowDimension($y)->setRowHeight(12);
	foreach ($subtotals as $subtotal => $subtotalPosition) {
		$y = $subtotalPosition['start'] - 1;
		$worksheet->setCellValueExplicitByColumnAndRow(1, $y, $subtotal, DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(2, $y, 'A.C. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(3, $y, 'A.P. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(4, $y, 'VAR.%', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(5, $y, 'A.C. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(6, $y, 'A.P. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(7, $y, 'VAR.%', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(8, $y, 'A.C. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(9, $y, 'A.P. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(10, $y, 'VAR.%', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(11, $y, 'A.C. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(12, $y, 'A.P. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(13, $y, 'VAR.%', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(14, $y, 'A.C. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(15, $y, 'A.P. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(16, $y, 'VAR.%', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(17, $y, 'A.C. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(18, $y, 'A.P. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(19, $y, 'VAR.%', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(20, $y, 'A.C. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(21, $y, 'A.P. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(22, $y, 'VAR.%', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(23, $y, 'A.C. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(24, $y, 'A.P. €', DataType::TYPE_STRING);
		$worksheet->setCellValueExplicitByColumnAndRow(25, $y, 'VAR.%', DataType::TYPE_STRING);

		$y = $subtotalPosition['start'];
		foreach ($sheet as $row) {
			/** costanti di totale */
			$customerCY = $row['customerCY'];
			$customerLY = $row['customerLY'];

			if ($row['subtotal'] == $subtotal) {
				/** costanti di sottototale */
				$customerCountPerSubtotalCY = $row['customerCountPerSubtotalCY'];
				$customerCountPerSubtotalLY = $row['customerCountPerSubtotalLY'];

				//$worksheet->setCellValueExplicitByColumnAndRow(1, $y, $row['department'], DataType::TYPE_STRING);

				/** incasso */
				$incassoAC_XY = XY(2, $y);
				$worksheet->setCellValueExplicit($incassoAC_XY, $row['totaltaxableamountCY'], DataType::TYPE_NUMERIC);

				$incassoAP_XY = XY(3, $y);
				$worksheet->setCellValueExplicit($incassoAP_XY, $row['totaltaxableamountLY'], DataType::TYPE_NUMERIC);

				$incassoDelta_XY = XY(4, $y);
				$formula = "=IF($incassoAP_XY<>0, ($incassoAC_XY-$incassoAP_XY)/$incassoAP_XY,0)";
				$worksheet->setCellValueExplicit($incassoDelta_XY, $formula, DataType::TYPE_FORMULA);

				/** mix vendite */
				$mixAC_XY = XY(5, $y);
				$formula = "=IF($incassoTotaleAC_XY<>0, $incassoAC_XY/$incassoTotaleAC_XY,0)";
				$worksheet->setCellValueExplicit($mixAC_XY, $formula, DataType::TYPE_FORMULA);

				$mixAP_XY = XY(6, $y);
				$formula = "=IF($incassoTotaleAP_XY<>0, $incassoAP_XY/$incassoTotaleAP_XY,0)";
				$worksheet->setCellValueExplicit($mixAP_XY, $formula, DataType::TYPE_FORMULA);

				$mixDelta_XY = XY(7, $y);
				$formula = "=$mixAC_XY - $mixAP_XY";
				$worksheet->setCellValueExplicit($mixDelta_XY, $formula, DataType::TYPE_FORMULA);

				/** clienti */
				$clientiAC_XY = XY(8, $y);
				$worksheet->setCellValueExplicit($clientiAC_XY, $row['customerCountPerDepartmentCY'], DataType::TYPE_NUMERIC);

				$clientiAP_XY = XY(9, $y);
				$worksheet->setCellValueExplicit($clientiAP_XY, $row['customerCountPerDepartmentLY'], DataType::TYPE_NUMERIC);

				$clientiDelta_XY = XY(10, $y);
				$formula = "=IF($clientiAP_XY<>0, ($clientiAC_XY-$clientiAP_XY)/$clientiAP_XY,0)";
				$worksheet->setCellValueExplicit($clientiDelta_XY, $formula, DataType::TYPE_FORMULA);

				/** penetrazione */
				$penetrazioneAC_XY = XY(11, $y);
				$formula = "=IF($clientiTotaleAC_XY<>0, $clientiAC_XY/$clientiTotaleAC_XY,0)";
				$worksheet->setCellValueExplicit($penetrazioneAC_XY, $formula, DataType::TYPE_FORMULA);

				$penetrazioneAP_XY = XY(12, $y);
				$formula = "=IF($clientiTotaleAP_XY<>0, $clientiAP_XY/$clientiTotaleAP_XY,0)";
				$worksheet->setCellValueExplicit($penetrazioneAP_XY, $formula, DataType::TYPE_FORMULA);

				$penetrazioneDelta_XY = XY(13, $y);
				$formula = "=$penetrazioneAC_XY - $penetrazioneAP_XY";
				$worksheet->setCellValueExplicit($penetrazioneDelta_XY, $formula, DataType::TYPE_FORMULA);

				/** scontrino medio */
				$scontrinoMedioAC_XY = XY(14, $y);
				$formula = "=IF($clientiAC_XY<>0, $incassoAC_XY/$clientiAC_XY,0)";
				$worksheet->setCellValueExplicit($scontrinoMedioAC_XY, $formula, DataType::TYPE_FORMULA);

				$scontrinoMedioAP_XY = XY(15, $y);
				$formula = "=IF($clientiAP_XY<>0, $incassoAP_XY/$clientiAP_XY,0)";
				$worksheet->setCellValueExplicit($scontrinoMedioAP_XY, $formula, DataType::TYPE_FORMULA);

				$scontrinoMedioDelta_XY = XY(16, $y);
				$formula = "=IF($scontrinoMedioAP_XY<>0, ($scontrinoMedioAC_XY - $scontrinoMedioAP_XY)/$scontrinoMedioAP_XY,0)";
				$worksheet->setCellValueExplicit($scontrinoMedioDelta_XY, $formula, DataType::TYPE_FORMULA);

				/** pezzi */
				$pezziAC_XY = XY(17, $y);
				$worksheet->setCellValueExplicit($pezziAC_XY, $row['quantityCY'], DataType::TYPE_NUMERIC);

				$pezziAP_XY = XY(18, $y);
				$worksheet->setCellValueExplicit($pezziAP_XY, $row['quantityLY'], DataType::TYPE_NUMERIC);

				$pezziDelta_XY = XY(19, $y);
				$formula = "=IF($pezziAP_XY<>0, ($pezziAC_XY - $pezziAP_XY)/$pezziAP_XY,0)";
				$worksheet->setCellValueExplicit($pezziDelta_XY, $formula, DataType::TYPE_FORMULA);

				/** prezzo medio */
				$prezzoMedioAC_XY = XY(20, $y);
				$formula = "=IF($pezziAC_XY<>0, $incassoAC_XY/$pezziAC_XY,0)";
				$worksheet->setCellValueExplicit($prezzoMedioAC_XY, $formula, DataType::TYPE_FORMULA);

				$prezzoMedioAP_XY = XY(21, $y);
				$formula = "=IF($pezziAP_XY<>0, $incassoAP_XY/$pezziAP_XY,0)";
				$worksheet->setCellValueExplicit($prezzoMedioAP_XY, $formula, DataType::TYPE_FORMULA);

				$prezzoMedioDelta_XY = XY(22, $y);
				$formula = "=IF($prezzoMedioAP_XY<>0, ($prezzoMedioAC_XY - $prezzoMedioAP_XY)/$prezzoMedioAP_XY,0)";
				$worksheet->setCellValueExplicit($prezzoMedioDelta_XY, $formula, DataType::TYPE_FORMULA);

				/** pezzi per cliente */
				$pezziPerClienteAC_XY = XY(23, $y);
				$formula = "=IF($clientiAC_XY<>0, $pezziAC_XY/$clientiAC_XY,0)";
				$worksheet->setCellValueExplicit($pezziPerClienteAC_XY, $formula, DataType::TYPE_FORMULA);

				$pezziPerClienteAP_XY = XY(24, $y);
				$formula = "=IF($clientiAP_XY<>0, $pezziAP_XY/$clientiAP_XY,0)";
				$worksheet->setCellValueExplicit($pezziPerClienteAP_XY, $formula, DataType::TYPE_FORMULA);

				$pezziPerClienteDelta_XY = XY(25, $y);
				$formula = "=IF($pezziPerClienteAP_XY<>0, ($pezziPerClienteAC_XY - $pezziPerClienteAP_XY)/$pezziPerClienteAP_XY,0)";
				$worksheet->setCellValueExplicit($pezziPerClienteDelta_XY, $formula, DataType::TYPE_FORMULA);

				$worksheet->getDefaultRowDimension()->setRowHeight(24);
				$y++;
			}
		}

		/**
		 * SOTTOTOTALI
		 */
		if(true) {
			$incassoSubtotaleLabel_XY = XY(1, $subtotalPosition['stop'] + 1);
			$worksheet->setCellValueExplicit($incassoSubtotaleLabel_XY, 'SOTTOTOTALI', DataType::TYPE_STRING);

			/** incasso */
			$incassoSubtotaleAC_XY = XY(2, $subtotalPosition['stop'] + 1);
			$incassoSubtotaleAC_RXY = RXY(2, $subtotalPosition['start'], 2, $subtotalPosition['stop']);
			$formula = "=SUM($incassoSubtotaleAC_RXY)";
			$worksheet->setCellValueExplicit($incassoSubtotaleAC_XY, $formula, DataType::TYPE_FORMULA);

			$incassoSubtotaleAP_XY = XY(3, $subtotalPosition['stop'] + 1);
			$incassoSubtotaleAP_RXY = RXY(3, $subtotalPosition['start'], 3, $subtotalPosition['stop']);
			$formula = "=SUM($incassoSubtotaleAP_RXY)";
			$worksheet->setCellValueExplicit($incassoSubtotaleAP_XY, $formula, DataType::TYPE_FORMULA);

			$incassoSubtotaleDelta_XY = XY(4, $subtotalPosition['stop'] + 1);
			$formula = "=IF($incassoSubtotaleAP_XY<>0, ($incassoSubtotaleAC_XY-$incassoSubtotaleAP_XY)/$incassoSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($incassoSubtotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

			/** mix vendite */
			$mixSubtotaleAC_XY = XY(5, $subtotalPosition['stop'] + 1);
			$formula = "=IF($incassoTotaleAC_XY<>0, $incassoSubtotaleAC_XY/$incassoTotaleAC_XY,0)";
			$worksheet->setCellValueExplicit($mixSubtotaleAC_XY, $formula, DataType::TYPE_FORMULA);

			$mixSubtotaleAP_XY = XY(6, $subtotalPosition['stop'] + 1);
			$formula = "=IF($incassoTotaleAP_XY<>0, $incassoSubtotaleAP_XY/$incassoTotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($mixSubtotaleAP_XY, $formula, DataType::TYPE_FORMULA);

			$mixSubtotaleDelta_XY = XY(7, $subtotalPosition['stop'] + 1);
			$formula = "=$mixSubtotaleAC_XY - $mixSubtotaleAP_XY";
			$worksheet->setCellValueExplicit($mixSubtotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

			/** clienti */
			$clientiSubtotaleAC_XY = XY(8, $subtotalPosition['stop'] + 1);
			$worksheet->setCellValueExplicit($clientiSubtotaleAC_XY, $customerCountPerSubtotalCY, DataType::TYPE_NUMERIC);

			$clientiSubtotaleAP_XY = XY(9, $subtotalPosition['stop'] + 1);
			$worksheet->setCellValueExplicit($clientiSubtotaleAP_XY, $customerCountPerSubtotalLY, DataType::TYPE_NUMERIC);

			$clientiSubtotaleDelta_XY = XY(10, $subtotalPosition['stop'] + 1);
			$formula = "=IF($clientiSubtotaleAP_XY<>0, ($clientiSubtotaleAC_XY-$clientiSubtotaleAP_XY)/$clientiSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($clientiSubtotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

			/** penetrazione */
			$penetrazioneSubtotaleAC_XY = XY(11, $subtotalPosition['stop'] + 1);
			$formula = "=IF($clientiTotaleAC_XY<>0, $clientiSubtotaleAC_XY/$clientiTotaleAC_XY,0)";
			$worksheet->setCellValueExplicit($penetrazioneSubtotaleAC_XY, $formula, DataType::TYPE_FORMULA);

			$penetrazioneSubtotaleAP_XY = XY(12, $subtotalPosition['stop'] + 1);
			$formula = "=IF($clientiTotaleAP_XY<>0, $clientiSubtotaleAP_XY/$clientiTotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($penetrazioneSubtotaleAP_XY, $formula, DataType::TYPE_FORMULA);

			$penetrazioneSubtotaleDelta_XY = XY(13, $subtotalPosition['stop'] + 1);
			$formula = "=$penetrazioneSubtotaleAC_XY - $penetrazioneSubtotaleAP_XY";
			$worksheet->setCellValueExplicit($penetrazioneSubtotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

			/** scontrino medio */
			$scontrinoMedioSubtotaleAC_XY = XY(14, $subtotalPosition['stop'] + 1);
			$formula = "=IF($clientiSubtotaleAC_XY<>0, $incassoSubtotaleAC_XY/$clientiSubtotaleAC_XY,0)";
			$worksheet->setCellValueExplicit($scontrinoMedioSubtotaleAC_XY, $formula, DataType::TYPE_FORMULA);

			$scontrinoMedioSubtotaleAP_XY = XY(15, $subtotalPosition['stop'] + 1);
			$formula = "=IF($clientiSubtotaleAP_XY<>0, $incassoSubtotaleAP_XY/$clientiSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($scontrinoMedioSubtotaleAP_XY, $formula, DataType::TYPE_FORMULA);

			$scontrinoMedioSubtotaleDelta_XY = XY(16, $subtotalPosition['stop'] + 1);
			$formula = "=IF($scontrinoMedioSubtotaleAP_XY<>0, ($scontrinoMedioSubtotaleAC_XY - $scontrinoMedioSubtotaleAP_XY)/$scontrinoMedioSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($scontrinoMedioSubtotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

			/** pezzi */
			$pezziSubtotaleAC_XY = XY(17, $subtotalPosition['stop'] + 1);
			$pezziSubtotaleAC_RXY = RXY(17, $subtotalPosition['start'], 17, $subtotalPosition['stop']);
			$formula = "=SUM($pezziSubtotaleAC_RXY)";
			$worksheet->setCellValueExplicit($pezziSubtotaleAC_XY, $formula, DataType::TYPE_FORMULA);

			$pezziSubtotaleAP_XY = XY(18, $subtotalPosition['stop'] + 1);
			$pezziSubtotaleAP_RXY = RXY(18, $subtotalPosition['start'], 18, $subtotalPosition['stop']);
			$formula = "=SUM($pezziSubtotaleAP_RXY)";
			$worksheet->setCellValueExplicit($pezziSubtotaleAP_XY, $formula, DataType::TYPE_FORMULA);

			$pezziSubtotaleDelta_XY = XY(19, $subtotalPosition['stop'] + 1);
			$formula = "=IF($pezziSubtotaleAP_XY<>0, ($pezziSubtotaleAC_XY-$pezziSubtotaleAP_XY)/$pezziSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($pezziSubtotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

			/** prezzo medio */
			$prezzoMedioSubtotaleAC_XY = XY(20, $subtotalPosition['stop'] + 1);
			$formula = "=IF($pezziSubtotaleAC_XY<>0, $incassoSubtotaleAC_XY/$pezziSubtotaleAC_XY,0)";
			$worksheet->setCellValueExplicit($prezzoMedioSubtotaleAC_XY, $formula, DataType::TYPE_FORMULA);

			$prezzoMedioSubtotaleAP_XY = XY(21, $subtotalPosition['stop'] + 1);
			$formula = "=IF($pezziSubtotaleAP_XY<>0, $incassoSubtotaleAP_XY/$pezziSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($prezzoMedioSubtotaleAP_XY, $formula, DataType::TYPE_FORMULA);

			$prezzoMedioSubtotaleDelta_XY = XY(22, $subtotalPosition['stop'] + 1);
			$formula = "=IF($prezzoMedioSubtotaleAP_XY<>0, ($prezzoMedioSubtotaleAC_XY - $prezzoMedioSubtotaleAP_XY)/$prezzoMedioSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($prezzoMedioSubtotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

			/** pezzi per cliente */
			$pezziPerClienteSubtotaleAC_XY = XY(23, $subtotalPosition['stop'] + 1);
			$formula = "=IF($clientiSubtotaleAC_XY<>0, $pezziSubtotaleAC_XY/$clientiSubtotaleAC_XY,0)";
			$worksheet->setCellValueExplicit($pezziPerClienteSubtotaleAC_XY, $formula, DataType::TYPE_FORMULA);

			$pezziPerClienteSubtotaleAP_XY = XY(24, $subtotalPosition['stop'] + 1);
			$formula = "=IF($clientiSubtotaleAP_XY<>0, $pezziSubtotaleAP_XY/$clientiSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($pezziPerClienteSubtotaleAP_XY, $formula, DataType::TYPE_FORMULA);

			$pezziPerClienteSubtotaleDelta_XY = XY(25, $subtotalPosition['stop'] + 1);
			$formula = "=IF($pezziPerClienteSubtotaleAP_XY<>0, ($pezziPerClienteSubtotaleAC_XY - $pezziPerClienteSubtotaleAP_XY)/$pezziPerClienteSubtotaleAP_XY,0)";
			$worksheet->setCellValueExplicit($pezziPerClienteSubtotaleDelta_XY, $formula, DataType::TYPE_FORMULA);
		}
	}

	/**
	 * TOTALI
	 */
	if (true) {
		$totaleLabel_XY = XY(1, $total_Y);
		$worksheet->setCellValueExplicit($totaleLabel_XY, 'TOTALI', DataType::TYPE_STRING);

		/** incasso */
		//$incassoTotaleAC_XY = XY(2, $total_Y);
		$subtotals_Y = [];
		foreach ($subtotals as $subtotal) {
			$subtotals_Y[] = XY(2, $subtotal['stop'] + 1);
		}
		$formula = "=" . implode('+',$subtotals_Y);
		$worksheet->setCellValueExplicit($incassoTotaleAC_XY, $formula, DataType::TYPE_FORMULA);

		//$incassoTotaleAP_XY = XY(3, $total_Y);
		$subtotals_Y = [];
		foreach ($subtotals as $subtotal) {
			$subtotals_Y[] = XY(3, $subtotal['stop'] + 1);
		}
		$formula = "=" . implode('+',$subtotals_Y);
		$worksheet->setCellValueExplicit($incassoTotaleAP_XY, $formula, DataType::TYPE_FORMULA);

		$incassoTotaleDelta_XY = XY(4, $total_Y);
		$formula = "=IF($incassoTotaleAP_XY<>0, ($incassoTotaleAC_XY-$incassoTotaleAP_XY)/$incassoTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($incassoTotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

		/** clienti */
		//$clientiTotaleAC_XY = XY(8, $total_Y);
		$worksheet->setCellValueExplicit($clientiTotaleAC_XY, $customerCY, DataType::TYPE_NUMERIC);

		//$clientiTotaleAP_XY = XY(9, $total_Y);
		$worksheet->setCellValueExplicit($clientiTotaleAP_XY, $customerLY, DataType::TYPE_NUMERIC);

		$clientiTotaleDelta_XY = XY(10, $total_Y);
		$formula = "=IF($clientiTotaleAP_XY<>0, ($clientiTotaleAC_XY-$clientiTotaleAP_XY)/$clientiTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($clientiTotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

		/** scontrino medio */
		$scontrinoMedioTotaleAC_XY = XY(14, $total_Y);
		$formula = "=IF($clientiTotaleAC_XY<>0, $incassoTotaleAC_XY/$clientiTotaleAC_XY,0)";
		$worksheet->setCellValueExplicit($scontrinoMedioTotaleAC_XY, $formula, DataType::TYPE_FORMULA);

		$scontrinoMedioTotaleAP_XY = XY(15, $total_Y);
		$formula = "=IF($clientiTotaleAP_XY<>0, $incassoTotaleAP_XY/$clientiTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($scontrinoMedioTotaleAP_XY, $formula, DataType::TYPE_FORMULA);

		$scontrinoMedioTotaleDelta_XY = XY(16, $total_Y);
		$formula = "=IF($scontrinoMedioTotaleAP_XY<>0, ($scontrinoMedioTotaleAC_XY - $scontrinoMedioTotaleAP_XY)/$scontrinoMedioTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($scontrinoMedioTotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

		/** pezzi */
		//$pezziTotaleAC_XY = XY(17, $total_Y);
		$subtotals_Y = [];
		foreach ($subtotals as $subtotal) {
			$subtotals_Y[] = XY(17, $subtotal['stop'] + 1);
		}
		$formula = "=" . implode('+',$subtotals_Y);
		$worksheet->setCellValueExplicit($pezziTotaleAC_XY, $formula, DataType::TYPE_FORMULA);

		//$pezziTotaleAP_XY = XY(18, $total_Y);
		$subtotals_Y = [];
		foreach ($subtotals as $subtotal) {
			$subtotals_Y[] = XY(18, $subtotal['stop'] + 1);
		}
		$formula = "=" . implode('+',$subtotals_Y);
		$worksheet->setCellValueExplicit($pezziTotaleAP_XY, $formula, DataType::TYPE_FORMULA);

		$pezziTotaleDelta_XY = XY(19, $total_Y);
		$formula = "=IF($pezziTotaleAP_XY<>0, ($pezziTotaleAC_XY-$pezziTotaleAP_XY)/$pezziTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($pezziTotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

		/** prezzo medio */
		$prezzoMedioTotaleAC_XY = XY(20, $total_Y);
		$formula = "=IF($pezziTotaleAC_XY<>0, $incassoTotaleAC_XY/$pezziTotaleAC_XY,0)";
		$worksheet->setCellValueExplicit($prezzoMedioTotaleAC_XY, $formula, DataType::TYPE_FORMULA);

		$prezzoMedioTotaleAP_XY = XY(21, $total_Y);
		$formula = "=IF($pezziTotaleAP_XY<>0, $incassoTotaleAP_XY/$pezziTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($prezzoMedioTotaleAP_XY, $formula, DataType::TYPE_FORMULA);

		$prezzoMedioTotaleDelta_XY = XY(22, $total_Y);
		$formula = "=IF($prezzoMedioTotaleAP_XY<>0, ($prezzoMedioTotaleAC_XY - $prezzoMedioTotaleAP_XY)/$prezzoMedioTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($prezzoMedioTotaleDelta_XY, $formula, DataType::TYPE_FORMULA);

		/** pezzi per scontrino */
		$pezziPerScontrinoTotaleAC_XY = XY(23, $total_Y);
		$formula = "=IF($clientiTotaleAC_XY<>0, $pezziTotaleAC_XY/$clientiTotaleAC_XY,0)";
		$worksheet->setCellValueExplicit($pezziPerScontrinoTotaleAC_XY, $formula, DataType::TYPE_FORMULA);

		$pezziPerScontrinoTotaleAP_XY = XY(24, $total_Y);
		$formula = "=IF($clientiTotaleAP_XY<>0, $pezziTotaleAP_XY/$clientiTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($pezziPerScontrinoTotaleAP_XY, $formula, DataType::TYPE_FORMULA);

		$pezziPerScontrinoTotaleDelta_XY = XY(25, $total_Y);
		$formula = "=IF($pezziPerScontrinoTotaleAP_XY<>0, ($pezziPerScontrinoTotaleAC_XY - $pezziPerScontrinoTotaleAP_XY)/$pezziPerScontrinoTotaleAP_XY,0)";
		$worksheet->setCellValueExplicit($pezziPerScontrinoTotaleDelta_XY, $formula, DataType::TYPE_FORMULA);
	}
}

function XY(int $x, int $y): string {
	return Coordinate::stringFromColumnIndex( $x ) . $y;
}

function RXY(int $x1, int $y1, int $x2, int $y2): string {
	return Coordinate::stringFromColumnIndex( $x1 ) . $y1. ':' . Coordinate::stringFromColumnIndex( $x2 ) . $y2;
}



