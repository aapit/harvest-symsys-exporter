<?php

//require 'PHPExcel/Classes/PHPExcel.php';
require 'vendor/autoload.php';

class HarvestSheet {
	const ERROR_NO_FILE =
		'Please provide a filename for the source, as first parameter.';
	const ERROR_NO_VALID_OUTPUT_TYPE =
		'Please provide a valid output type for the output.';
	const HEADER_ROW = 1;
	const FIRST_CONTENT_ROW = 2;
	const OUTPUT_TYPE_XLS = 'xls';
	const OUTPUT_TYPE_CSV = 'csv';
	const INPUT_TYPE_XLSX = 'xlsx';
	const DATE_FORMAT = 'dd-mm-yyyy';


	protected $_dateColumnLabels = array('Date');


	/**
 	 * @var PHPExcel
 	 */
	protected $_excelDoc;

	protected $_path;

	/**
 	 * Store the first blank column.
 	 */
	protected $_destColumn;


	public function __construct($path) {
		$this->_path = $path;
		$this->_excelDoc = $this->_openFile();

		$this->_formatDateColumns();
		$this->_initSheet();
	}

	public function output($type) {
		$path = 
			dirname($this->_path) . DIRECTORY_SEPARATOR
			. basename($this->_path, '.' . self::INPUT_TYPE_XLSX)
		;

		switch ($type) {
			case self::OUTPUT_TYPE_XLS:
				$objWriter = new PHPExcel_Writer_Excel5($this->_excelDoc);
				$extension = '.' . self::OUTPUT_TYPE_XLS;
				return $objWriter->save($path . $extension);
			case self::OUTPUT_TYPE_CSV:
				$objWriter = PHPExcel_IOFactory::createWriter($this->_excelDoc, 'CSV');
				$extension = '.' . self::OUTPUT_TYPE_CSV;
				return $objWriter->save($path . $extension);
		}

		throw new Exception(self::ERROR_NO_VALID_OUTPUT_TYPE);
	}

	public function getNumberOfContentRows() {
		return count($this->_getColumnContent('A'));
	}

	public function insertColumn($columnHeader, array $values) {
		$this->_setDestinationColumnName($columnHeader);
		$column = $this->_makeColumn($values);

		$this->_getSheet()->fromArray(
			$column,
			null,
			$this->_destColumn . self::FIRST_CONTENT_ROW
		);

		$this->_increaseDestColumn();
	}

	public function getColumnValues($columnHeader) {
		$columnIndex = $this->_getHeaderColumn($columnHeader);
		return $this->_getColumnContent($columnIndex);
	}

	public function getConcatColumnValues(array $columnHeaders, $separator = '') {
		$combinedValues = array_fill(0, $this->getNumberOfContentRows(), null);

		foreach ($columnHeaders as $columnHeader) {
			$columnIndex = $this->_getHeaderColumn($columnHeader);
			$values = $this->_getColumnContent($columnIndex);

			array_walk(
				$combinedValues,
				array($this, '_concatArrayValues'),
				array($separator, $values)
			);
		}

		return $combinedValues;
	}


	public function splitColumn($srcColumnHeader, $destHeader) {
		$srcColumnIndex = $this->_getHeaderColumn($srcColumnHeader);
		$this->_setDestinationColumnName($destHeader);

		$oldSrcContent = $this->_getColumnContent($srcColumnIndex);
		$newSrcContent = array_map(array($this, '_stripNumber'), $oldSrcContent);
		$newDestContent = array_map(array($this, '_extractNumber'), $oldSrcContent);
		$newSrcColumn = $this->_makeColumn($newSrcContent);
		$newDestColumn = $this->_makeColumn($newDestContent);

		$this->_getSheet()->fromArray(
			$newSrcColumn,
			null,
			$srcColumnIndex . self::FIRST_CONTENT_ROW	
		);

		$this->_getSheet()->fromArray(
			$newDestColumn,
			null,
			$this->_destColumn . self::FIRST_CONTENT_ROW
		);

		$this->_increaseDestColumn;
	}

	/**
 	 * Replace a given string by another in the values of a specific column.
 	 *
 	 * @param	String	$columnHeader	The text label this column has.
 	 * @param	String	$search			The string to search for.
 	 * @param	String	$replace		The string to replace it by.
 	 */
	public function replaceColumnString($columnHeader, $search, $replace) {
		$columnIndex = $this->_getHeaderColumn($columnHeader);
		$columnContent = $this->_getColumnContent($columnIndex);

		$cap = function(&$value, $key, array $searchAndReplace) {
			$value = str_replace(
				$searchAndReplace[0], $searchAndReplace[1], $value
			);
		};

		array_walk($columnContent, $cap, array($search, $replace));

		$this->_getSheet()->fromArray(
			$this->_makeColumn($columnContent),
			null,
			$columnIndex . self::FIRST_CONTENT_ROW
		);
	}

	/**
 	 * Limit the length of column values.
 	 *
 	 * @param	String	$columnHeader	The text label this column has.
 	 * @param	Int		$limit			The number of characters to limit to.
 	 */
	public function capColumn($columnHeader, $limit) {
		$columnIndex = $this->_getHeaderColumn($columnHeader);
		$columnContent = $this->_getColumnContent($columnIndex);

		$cap = function(&$value, $key, $limit) {
			$value = substr($value, 0, $limit);
		};

		array_walk($columnContent, $cap, $limit);

		$this->_getSheet()->fromArray(
			$this->_makeColumn($columnContent),
			null,
			$columnIndex . self::FIRST_CONTENT_ROW
		);
	}

	public function removeColumns(array $columnLabels) {
		foreach ($columnLabels as $label) {
			$column = $this->_getHeaderColumn($label);
			$this->_getSheet()->removeColumn($column, 1);
		}

		$this->_decreaseDestColumn(count($columnLabels));
	}

	protected function _increaseDestColumn() {
		$columnIndex = ord($this->_destColumn);
		$this->_destColumn = chr(++$columnIndex);
	}

	protected function _decreaseDestColumn($columns = 1) {
		$columnIndex = ord($this->_destColumn);
		$this->_destColumn = chr($columnIndex -= $columns);
	}

	/**
 	 * Concats different arrays of the same length.
 	 * Writes them back to the first argument, which is practical for array_walk().
 	 */
	protected function _concatArrayValues(&$combinedValue, $key, array $separatorAndToCombineValues) {
		list($separator, $toCombineValues) = $separatorAndToCombineValues;
		$combinedValue = implode(
			$separator,
			array($combinedValue, $toCombineValues[$key])
		);

		$combinedValue = trim($combinedValue, $separator);
	}

	protected function _initSheet() {
		//$this->_getSheet()->refreshColumnDimensions();
		//$this->_getSheet()->refreshRowDimensions();
		//$this->_getSheet()->updateCacheData();
		//PHPExcel_Calculation::getInstance()->clearCalculationCache();
		//$this->_destColumn = chr(ord($this->_getSheet()->getHighestDataColumn()));
		$this->_destColumn = chr(ord($this->_getSheet()->getHighestColumn()));
	}

	protected function _stripNumber($value) {
		$pattern = '/(\D+)( ?\(\d+\))/i';
		return trim(preg_replace($pattern, '$1', $value));
	}

	protected function _extractNumber($value) {
		if (strpos($value, '(') === false) {
			return null;
		}
		$pattern = '/[a-z \/\&]*(\s*\((\d+)\))?/i';
		return preg_replace($pattern, '$2', $value);
	}

	protected function _formatDateColumns() {
		$columns = array();

		foreach ($this->_dateColumnLabels as $label) {
			$columns[] = $this->_getHeaderColumn($label);
		}

		foreach ($columns as $column) {
			$this->_formatDateColumn($column);
		}
	}

	/**
 	 * @param String $column The column letter
 	 */
	protected function _formatDateColumn($column) {
		$this->_getSheet()
    		->getStyle(
				$this->_getColumnContentCoordinates($column)
			)
    		->getNumberFormat()
			->setFormatCode(self::DATE_FORMAT)
		;

		/*

		Mogelijkheden:

 		FORMAT_DATE_YYYYMMDD - deze werkt wel, maar geeft jaar slechts in 2 cijfers...
		FORMAT_DATE_DDMMYYYY - deze werkt wel, maar is misschien verwarrend
		FORMAT_DATE_DMYSLASH - deze lijkt ook niet te werken
		FORMAT_DATE_DMYMINUS - deze lijkt niet te werken

		meer op:
		http://www.cmsws.com/examples/applications/phpexcel/Documentation/API/PHPExcel_Style/PHPExcel_Style_NumberFormat.html#constFORMAT_DATE_DMYSLASH

		*/
	}


	/**
 	 * @param String $column	The column letter
 	 * @return String			The coordinates of the content part of the column,
 	 *							sans header. For instance: "A2:A599"
 	 */
	protected function _getColumnContentCoordinates($column) {
		return 
			$column . self::FIRST_CONTENT_ROW
			. ':'
			. $column . $this->_getSheet()->getHighestRow()
		;
	}

	/**
 	 * Retrieves the header column that corresponds with given label.
 	 * @return String The column letter
 	 */
	protected function _getHeaderColumn($label) {
		$headerRow = $this->_getHeaderRow();

		return array_search($label, $headerRow);
	}

	/**
 	 * Retrieves the cell values in the header row.
 	 * @return Array array(
 	 *					'A' => 'Column Name 1',
 	 *					'B' => 'Column Name 2'
 	 *				 )
 	 */
	protected function _getHeaderRow() {
		$headerCells = array();

		$row = $this->_getSheet()->getRowIterator(self::HEADER_ROW)->current();
		$cellIterator = $row->getCellIterator();
		$cellIterator->setIterateOnlyExistingCells(false);

		foreach ($cellIterator as $cell) {
    		$headerCells[$cell->getColumn()] = $cell->getValue();
		}

		return $headerCells;
	}

	protected function _getColumnContent($srcColumn) {
		$highestRow = $this->_getSheet()->getHighestRow();
		$from = $srcColumn . self::FIRST_CONTENT_ROW;
		$to = $srcColumn . $highestRow;

		$columnData = $this->_getSheet()
			->rangeToArray(
				$from . ':' . $to
			)
		;

		$flatten = function(&$value) {
			$value = $value[0];
		};
		array_walk($columnData, $flatten);

		return $columnData;
	}

	protected function _makeColumn(array $values) {
		return array_chunk($values, 1);
	}

	protected function _setDestinationColumnName($destName) {
		$this->_getSheet()->setCellValue($this->_destColumn . self::HEADER_ROW, $destName);
	}

	/**
 	 * @return PHPExcel
 	 */
	protected function _openFile() {
		if (empty($this->_path)) {
			throw new Exception(self::ERROR_NO_FILE);
		}

		$excelDoc = PHPExcel_IOFactory::load($this->_path);
		return $excelDoc;
	}

	protected function _getSheet() {
		return $this->_excelDoc->getActiveSheet();
	}
}
