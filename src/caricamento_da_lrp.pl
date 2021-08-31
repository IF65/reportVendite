#!/usr/bin/perl -w
use strict;
use DBI;
use File::HomeDir;
use DateTime;
use POSIX;

# data di partenza caricamento
#---------------------------------------------------------------------------------------
my $current_date = DateTime->today(time_zone=>'local');
my $starting_date = $current_date->add(days =>-420)->ymd('/');#$current_date->add(years => -2)->ymd('/');

# parametri di configurazione del database mysql
#---------------------------------------------------------------------------------------
my $hostname    = "localhost";
my $username    = "root";
my $password    = "mela";

my $database    = "controllo";

# parametri di configurazione del database oracle
#---------------------------------------------------------------------------------------
my $hostname_or = "10.11.14.109";
my $username_or = "test";
my $password_or = "user1";

# variabili globali
#---------------------------------------------------------------------------------------
my $dbh;
my $sth_select_or;
my $sth_ins_upd;
my $sth_drop;
my $sth_totali;
my $sth_indici;

# Connessione al db
#---------------------------------------------------------------------------------------
if (!&ConnessioneDB()) {die;}

if ($sth_select_or->execute()) {
	while ( my @row = $sth_select_or->fetchrow_array() ) {
		my $societa = substr($row[0],0,2);
		my $negozio = $row[0];
		my $data = $row[1];
		my $cassa = $row[2];
		my $transazione = $row[3];
		my $ora = floor($row[4]/60);
		my $minuto = $row[4] % 60;
		my $importo = $row[5];
		my $punti_emessi_prodotto = $row[6];
		my $punti_emessi_reparto = $row[7];
		my $punti_emessi_ticket = $row[8];
		my $punti_redenti = $row[9];
		my $carta = $row[10];

		if ($carta !~ /^046/) {$carta = ''};

		if (!$sth_ins_upd->execute($societa, $negozio, $data, $cassa, $transazione, $ora, $minuto, $importo, $punti_emessi_prodotto, $punti_emessi_reparto, $punti_emessi_ticket, $punti_redenti, $carta)) {
			print "Errore durante l'inserimento/modifica di un record su db! " .$dbh->errstr."\n";
			return 0;
		}
	}
}

# if ($sth_drop->execute()) {
# 	if ($sth_totali->execute()) {
# 		if ($sth_indici->execute()) {
# 		}
# 	}
# }

$sth_ins_upd->finish();
$sth_select_or->finish();
$sth_drop->finish();
$sth_totali->finish();
$sth_indici->finish();

sub ConnessioneDB{
	my $sth;

	$dbh = DBI->connect('dbi:Oracle:',qq{$username_or/$password_or@(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$hostname_or)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=orcl)))},"") or die;
	$sth_select_or = $dbh->prepare(
							qq{ select
								p.CODICE,
								to_char(to_date(t.ID_TEMPO,'J'),'YYYY-MM-DD'),
								lpad(t.NUMERO_CASSA,3,'0'),
								lpad(t.NUMERO_SCONTRINO,4,'0'),
								ID_MINUTO, 
								nvl(t.IMPORTO,0),
								nvl(t.PUNTI_EMESSI_PRODOTTO,0),
								nvl(t.PUNTI_EMESSI_REPARTO,0),
								nvl(t.PUNTI_EMESSI_TICKET,0),
								nvl(t.PUNTI_REDENTI,0),
								nvl(c.CARTA, '')
								from testate t join pdv p on t.ID_PDV = p.ID left outer join clienti c on t.ID_CLIENTE = c.ID
								where to_date(t.ID_TEMPO,'J')>=to_date('$starting_date','YYYY/MM/DD')
								order by 2,1 asc
							});

	$dbh = DBI->connect("DBI:mysql:$database:$hostname", $username, $password) or die;

	# creazione della tabella scontrini_lrp
	$sth = $dbh->prepare(
							qq{
								 CREATE TABLE IF NOT EXISTS `testate_lrp` (
								  `societa` varchar(2) NOT NULL DEFAULT '',
								  `negozio` varchar(4) NOT NULL DEFAULT '',
								  `data` date NOT NULL,
								  `cassa` varchar(4) NOT NULL DEFAULT '',
								  `transazione` varchar(4) NOT NULL DEFAULT '',
								  `ora` tinyint(4) NOT NULL,
								  `minuto` int(11) NOT NULL DEFAULT '0',
								  `totale` float NOT NULL DEFAULT '0',
								  `punti_emessi_prodotto` float NOT NULL DEFAULT '0',
								  `punti_emessi_reparto` float NOT NULL DEFAULT '0',
								  `punti_emessi_ticket` float NOT NULL DEFAULT '0',
								  `punti_redenti` float NOT NULL DEFAULT '0',
								  `carta` varchar(13) NOT NULL DEFAULT '',
								  PRIMARY KEY (`negozio`,`data`,`cassa`,`transazione`),
								  KEY `carta` (`carta`),
								  KEY `societa` (`societa`,`negozio`,`data`,`ora`)
								) ENGINE=InnoDB DEFAULT CHARSET=latin1;
							});

	if (!$sth->execute()) {
		print "Errore durante l'esecuzione di una query su db! " .$dbh->errstr."\n";
		return 0;
	}

	$sth->finish();

	$sth_ins_upd = $dbh->prepare(
							qq{
								insert into $database.testate_lrp(societa, negozio, data, cassa, transazione, ora, minuto, totale, punti_emessi_prodotto, punti_emessi_reparto, punti_emessi_ticket, punti_redenti, carta)
								values (?,?,?,?,?,?,?,?,?,?,?,?,?) on duplicate key update societa = values(societa), negozio = values(negozio), data = values(data), cassa = values(cassa),
								transazione = values(transazione), ora = values(ora), minuto = values(minuto), totale = values(totale), punti_emessi_prodotto = values(punti_emessi_prodotto), punti_emessi_reparto = values(punti_emessi_reparto),
								punti_emessi_ticket = values(punti_emessi_ticket), punti_redenti = values(punti_redenti), carta = values(carta)
							});

	$sth_drop = $dbh->prepare(qq{drop table if exists `totali_lrp`});
 	$sth_totali = $dbh->prepare(qq{create table `totali_lrp` as select negozio, data, count(*) as `clienti`, round(sum(totale),2) as `totale` from testate_lrp group by 1,2 order by 1,2});
 	$sth_indici = $dbh->prepare(qq{alter table `totali_lrp` add primary key(`negozio`,`data`)});

	return 1;
}
