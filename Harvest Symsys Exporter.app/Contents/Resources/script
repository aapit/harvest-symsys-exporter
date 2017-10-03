<?php

require 'Harvest/Sheet.php';
require 'Harvest/Transformation.php';
date_default_timezone_set('Europe/Amsterdam');

const RESOURCE_EMPLOYEE_URL = 'https://docs.google.com/spreadsheets/d/1-SnXuHxVAlO4MYby4PUfOJtZIji6WKKdB2XoaGIrdVE/pub?output=csv';
const RESOURCE_TASKCODE_URL = 'https://docs.google.com/spreadsheets/d/1-SnXuHxVAlO4MYby4PUfOJtZIji6WKKdB2XoaGIrdVE/pub?gid=1835643858&single=true&output=csv';

function getCodeMap($resource) {
    $cacheBuster = time();
    $delimiter = strpos($resource, '?') === false
        ? '?'
        : '&'
    ;
    $csvString = file_get_contents($resource . $delimiter . $cacheBuster);

    $splitColumns = function($row) {
        // remove control characters
        $row = preg_replace('/[\x00-\x1F\x7F]/', '', $row);
        return explode(',', $row);
    };

    $map = array_map($splitColumns, explode("\n", $csvString));

    $output = array();

    foreach ($map as $row) {
        $output[$row[0]] = $row[1];
    }

    return $output;
}

$employeeNumberMap = getCodeMap(RESOURCE_EMPLOYEE_URL);
$taskCodeMap = getCodeMap(RESOURCE_TASKCODE_URL);

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
$sheet->replaceColumnString('Notes', ';', ',');
$sheet->replaceColumnString('Notes', "\n", ' | ');
$sheet->replaceColumnString('Notes', "\r", ' | ');

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
	'First Name', 'Last Name', 'Project Code', 'Department', 'Roles',
	'Task', 'Client', 'Project'
));

// Output as CSV.
$sheet->output(HarvestSheet::OUTPUT_TYPE_CSV, false);
