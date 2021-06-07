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
$dataCorrenteAC = new DateTime('2021-01-31', $timeZone);
$dataInizioAC = new DateTime('2021-01-25', $timeZone);
$dataFineAC = new DateTime('2021-01-31', $timeZone);
$dataCorrenteAP = new DateTime('2020-02-02', $timeZone);
$dataInizioAP = new DateTime('2020-01-27', $timeZone);
$dataFineAP = new DateTime('2020-02-02', $timeZone);

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
$workBook = new Spreadsheet();
foreach ($shopList as $shopCode => $shopDescription) {
	$sheet = [];
	foreach ($result as $row) {
		if ($row['store'] == $shopCode) {
			$sheet[] = $row;
		}
	}
	$
}

function fillSheet(Worksheet $worksheet) {
	$
}




