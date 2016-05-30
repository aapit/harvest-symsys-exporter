<?php

require 'Harvest/Sheet.php';
require 'Harvest/Transformation.php';
date_default_timezone_set('Europe/Amsterdam');


$employeeNumberMap = array(
	100 => 'Mattijs Bliek',
	102 => 'Jelmer Boomsma',
	103 => 'Josephine Cambier',
	105 => 'Rolf Coppens',
	106 => 'Jeroen Disch',
	107 => 'Pieter-Jannick Dijkstra',
	108 => 'Ramiro Hammen',
	109 => 'Harmen Janssen',
	110 => 'Larix Kortbeek',
	111 => 'Koen Schaft',
	112 => 'Claudia van Schendel',
	113 => 'David Spreekmeester',
	117 => 'Clara Dujardin',
	118 => 'Roos Floris',
	119 => 'Justine Servais',
	120 => 'Jean Bohm'
);

$taskCodeMap = array(
	90 => 'Administratief',
	105 => 'Art direction',
	100 => 'Concept (FO/TO/COPY)',
	113 => 'Design interactief',
	115 => 'Design visueel',
	112 => 'Development Back end',
	111 => 'Development Front end',
	311 => 'DTP',
	92 => 'Hosting/server/devops',
	93 => 'Offerte inschatten',
	94 => 'Planning',
	15 => 'Project management',
	43 => 'Strategie en Consultancy',
	95 => 'Testing',
	96 => 'Tools (proces en workflow)'
);

$path = $argv[1];
$sheet = new HarvestSheet($path);

$numberOfContentRows = $sheet->getNumberOfContentRows();
$projectCodes = $sheet->getColumnValues('Project Code');

// Find employee numbers by first and last name columns.
$employeeNames = $sheet->getConcatColumnValues(array('First Name', 'Last Name'), ' ');
$employeeNumberRows = array();
foreach ($employeeNames as $employeeName) {
	$employeeNumberRows[] = array_search($employeeName, $employeeNumberMap);
}

// Find work codes by textual Task category
$taskDescriptions = $sheet->getColumnValues('Task');
$taskCodeRows = array();
foreach ($taskDescriptions as $taskDescription) {
	$taskCodeRows[] = array_search($taskDescription, $taskCodeMap);
}

//$sheet->splitColumn('Client', 'Client Code');

// Cap the descriptions, since Symsys only supports 50 chars.
$sheet->capColumn('Notes', 50);

// Convert hours to Dutch decimal separator.
$sheet->replaceColumnString('Hours', '.', ',');

// Now combine the column with a couple of dynamic values.
/*
$transformation = new HarvestTransformation($sheet);
$transformation->addCombinedColumn('Code', array(
	// First digit is always 1, for Grrr
	array_fill(0, $numberOfContentRows, '1'),
	$employeeNumberRows,
	$projectCodes,
	$taskCodeRows
));
*/
$sheet->insertColumn('Bedrijf', array_fill(0, $numberOfContentRows, '1'));
$sheet->insertColumn('Medewerker', $employeeNumberRows);
$sheet->insertColumn('Project', $projectCodes);
$sheet->insertColumn('Werkcode', $taskCodeRows);

$sheet->removeColumns(array(
	'Billable?', 'Invoiced?', 'Approved?', 'Employee?', 'Billable Rate',
	'Billable Amount', 'Cost Rate', 'Cost Amount', 'Currency',
	'First Name', 'Last Name', 'Project Code', 'Department',
	'Task', 'Client', 'Project'
));

// Output as CSV.
$sheet->output(HarvestSheet::OUTPUT_TYPE_CSV, false);