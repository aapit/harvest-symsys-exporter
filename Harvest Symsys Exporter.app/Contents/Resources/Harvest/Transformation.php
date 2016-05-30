<?php

class HarvestTransformation {
	/** @var HarvestSheet $_sheet **/
	protected $_sheet;


	public function __construct(HarvestSheet $sheet) {
		$this->_sheet = $sheet;
	}

	public function addCombinedColumn($columnHeader, array $mergableArrays) {
		$numberOfContentRows = $this->_sheet->getNumberOfContentRows();
		$combinedCodes = array_fill(0, $numberOfContentRows, null);

		array_walk($combinedCodes, array($this, '_concatArrayValues'), $mergableArrays);
		$this->_sheet->insertColumn($columnHeader, $combinedCodes);
	}

	/**
 	 * Concats different arrays of the same length.
 	 * Writes them back to the first argument, which is practical for array_walk().
 	 */
	protected function _concatArrayValues(&$value, $key, array $mergableArrays) {
		$combinedValue = '';

		foreach($mergableArrays as $mergableArray) {
			$combinedValue .= $mergableArray[$key];
		}

		$value = $combinedValue;
	}
}
