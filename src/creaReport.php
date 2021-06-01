<?php
ini_set('memory_limit','8192M');

require __DIR__ . "/Report/autoload.php";

use \Report\Report;


$report = new Report();

//$dati = $report->getWeekData(Report::$REPORT_SHOP, 14, 2021);
$dati = $report->getData('2021-04-19', '2021-04-25', '2020-04-20', '2020-04-20');

file_put_contents('/Users/if65/Desktop/dati.json', $dati);


echo "\n";