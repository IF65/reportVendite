<?php


namespace Report;

use \PDO;

class Report
{
	private $db = null;

	private $host = '10.11.14.248';
	private $user = 'root';
	private $password = 'mela';

	public static $REPORT_SHOP = 10;
	public static $REPORT_SHOP_TYPE = 20;
	public static $REPORT_AREA = 30;

	public function __construct() {
		$this->connect();
	}

	/**
	 * Connessione al database
	 */
	public function connect()
	{
		try {
			$this->db = new PDO("mysql:host=" . $this->host, $this->user, $this->password, [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION]);
		} catch (PDOException $e) {
			echo "Errore: " . $e->getMessage();
			die();
		}
	}

	/**
	 * @param string $host
	 */
	public function setHost(string $host): void
	{
		$this->host = $host;
	}

	/**
	 * @return string
	 */
	public function getUser(): string
	{
		return $this->user;
	}

	/**
	 * @param string $password
	 */
	public function setPassword(string $password): void
	{
		$this->password = $password;
	}

	/**
	 * @return null
	 */
	public function getDb()
	{
		return $this->db;
	}

	/**
	 * @return data:string
	 */
	public function getWeekData(int $reportType, int $week, int $year, int $step = -1): string
	{
		try {
			$stmt = " select min(ddate) firstDate, max(ddate) lastDate from mtx.calendar where week = :week and year = :year;";
			$h_query = $this->db->prepare($stmt);

			$h_query->execute([':week' => $week, ':year' => $year]);
			$result = ($h_query->fetchAll(PDO::FETCH_ASSOC))[0];
			$firstDateCY = $result['firstDate'];
			$lastDateCY = $result['lastDate'];

			$h_query->execute([':week' => $week, ':year' => $year + $step]);
			$result = ($h_query->fetchAll(PDO::FETCH_ASSOC))[0];
			$firstDateLY = $result['firstDate'];
			$lastDateLY = $result['lastDate'];

			if ($reportType == Report::$REPORT_SHOP) {
				$stmt = "select
						d.store, 
						d.department,
						r.sortOrder,
						r.subtotal,
						r.sortOrderSubtotal,
						n.area,
						n.tipoFood,
						n.negozio_descrizione description,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then d.totaltaxableamount else 0 end),0) totaltaxableamountCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then d.totaltaxableamount else 0 end),0) totaltaxableamountLY,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then d.rowCount else 0 end),0) rowCountCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then d.rowCount else 0 end),0) rowCountLY,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then d.quantity else 0 end),0) quantityCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then d.quantity else 0 end),0) quantityLY,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then d.customerCount else 0 end),0) customerCountPerDepartmentCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then d.customerCount else 0 end),0) customerCountPerDepartmentLY,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then p.customerCount else 0 end),0) customerCountPerSubtotalCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then p.customerCount else 0 end),0) customerCountPerSubtotalLY,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then c.customerCount else 0 end),0) customerCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then c.customerCount else 0 end),0) customerLY
					from mtx.salesPerDepartment as d join 
					    (select distinct nuovoReparto department, sortOrder, subtotali subtotal, sortOrderSubTotale sortOrderSubtotal from mtx.sottoreparto ) as r on d.department = r.department 
					    	join mtx.customers as c on c.`ddate`=d.`ddate` and c.`store`=d.`store` 
					    	join mtx.penetrationPerSubtotal as p on p.`ddate`=d.`ddate` and p.`store`=d.`store` and p.subtotal=r.subtotal 
					    	join archivi.negozi as n on d.store = n.codice 
					where (d.ddate >= :firstDateCY and d.ddate <= :lastDateCY) or (d.ddate >= :firstDateLY and d.ddate <= :lastDateLY) 
					group by 1,2;";
				$h_query = $this->db->prepare($stmt);
				$h_query->execute([':firstDateCY' => $firstDateCY, ':lastDateCY' => $lastDateCY, ':firstDateLY' => $firstDateLY, ':lastDateLY' => $lastDateLY]);
				$result = $h_query->fetchAll(PDO::FETCH_ASSOC);
			}



			return json_encode($result, true);

		} catch (PDOException $e) {
			echo "Errore: " . $e->getMessage();
			die();
		}
	}

	/**
	 * @param string $firstDateCY
	 * @param string $lastDateCY
	 * @param string $firstDateLY
	 * @param string $lastDateLY
	 * @return string
	 * consolida e dati in base alle date e rstituisce un json raggruppato per sede/sottototale/reparto.
	 */
	public function getData(string $firstDateCY, string $lastDateCY, string $firstDateLY, string $lastDateLY): string
	{
		try {
			$stmt = "select 
       					codice store
					from archivi.negozi 
					where societa in ('02','05') and codicePrincipale = '' and 
					      ((data_inizio <= :lastDateCY and ifnull(data_fine, '2099-12-31') >= :firstDateCY) or (data_inizio <= :lastDateLY and ifnull(data_fine, '2099-12-31') >= :firstDateLY))";
			$h_query = $this->db->prepare($stmt);
			$h_query->execute([':firstDateCY' => $firstDateCY, ':lastDateCY' => $lastDateCY, ':firstDateLY' => $firstDateLY, ':lastDateLY' => $lastDateLY]);
			$rows = $h_query->fetchAll(PDO::FETCH_ASSOC);
			$result = [];
			foreach($rows as $row) {
				if (!key_exists($row['store'], $result)) {
					$result[$row['store']] = [];
				}
			}

			$stmt = "select nuovoReparto department, sortOrder, subtotali subtotal, sortOrderSubTotale sortOrderSubtotal from mtx.sottoreparto;";
			$h_query = $this->db->prepare($stmt);
			$h_query->execute();
			$rows = $h_query->fetchAll(PDO::FETCH_ASSOC);
			foreach($result as $store => $value) {
				foreach ($rows as $row) {
					if (!key_exists($row['subtotal'], $result[$store])) {
						$result[$store][$row['subtotal']] = [];
					}
					if (!key_exists($row['department'], $result[$store][$row['subtotal']])) {
						$result[$store][$row['subtotal']][$row['department']] = [];
					}
				}
			}

			$stmt = "select
						d.store, 
       					r.subtotal,
						d.department,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then d.totaltaxableamount else 0 end),0) totaltaxableamountCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then d.totaltaxableamount else 0 end),0) totaltaxableamountLY,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then d.rowCount else 0 end),0) rowCountCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then d.rowCount else 0 end),0) rowCountLY,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then d.quantity else 0 end),0) quantityCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then d.quantity else 0 end),0) quantityLY,
						ifnull(sum(case when d.ddate >= :firstDateCY and d.ddate <= :lastDateCY then d.customerCount else 0 end),0) customerCountPerDepartmentCY,
						ifnull(sum(case when d.ddate >= :firstDateLY and d.ddate <= :lastDateLY then d.customerCount else 0 end),0) customerCountPerDepartmentLY
					from mtx.salesPerDepartment as d join 
					    (select distinct nuovoReparto department, sortOrder, subtotali subtotal, sortOrderSubTotale sortOrderSubtotal from mtx.sottoreparto ) as r on d.department = r.department
					where (d.ddate >= :firstDateCY and d.ddate <= :lastDateCY) or (d.ddate >= :firstDateLY and d.ddate <= :lastDateLY) 
					group by 1, 2, 3;";
			$h_query = $this->db->prepare($stmt);
			$h_query->execute([':firstDateCY' => $firstDateCY, ':lastDateCY' => $lastDateCY, ':firstDateLY' => $firstDateLY, ':lastDateLY' => $lastDateLY]);
			$rows = $h_query->fetchAll(PDO::FETCH_ASSOC);

			foreach($rows as $row) {
				$result[ $row['store'] ][ $row['subtotal'] ][ $row['department'] ] = [
					'totaltaxableamountCY' => $row['totaltaxableamountCY'],
					'totaltaxableamountLY' => $row['totaltaxableamountLY'],
					'rowCountCY' => $row['rowCountCY'],
					'rowCountLY' => $row['rowCountLY'],
					'quantityCY' => $row['quantityCY'],
					'quantityLY' => $row['quantityLY'],
					'customerCountPerDepartmentCY' => $row['customerCountPerDepartmentCY'],
					'customerCountPerDepartmentLY' => $row['customerCountPerDepartmentLY']
				];
			}

			return json_encode($result, true);

		} catch (PDOException $e) {
			echo "Errore: " . $e->getMessage();
			die();
		}
	}

	public function __destruct()
	{
		unset($this->db);
	}

}