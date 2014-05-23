<?php
// Error reporting
error_reporting ( E_ALL );

// : Includes
include_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';

// : End

/**
 * Object::RatesReadXLSData
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Barloworld Transport Solutions (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class RatesReadXLSData {
	// : Constants
	const DISALLOWED_EXTENSION = "The file %s has an unsupported extension.";
	const FILE_DOESNT_EXIST = "The file %s could not be found.";
	const FILE_COULD_NOT_CREATE = "The file %s could not be created.";
	const FILE_NOT_READABLE = "The file %s could not be read. Probable causes: Incorrect permission or in a path that PHP cannot read.";
	const FOPEN_COULD_NOT_OPEN = "The file %s could not be opened.";
	const FOPEN_NOT_ALLOWED_TO_OPEN = "The file %s cannot be opened, as fopen cannot open urls.";
	const DS = DIRECTORY_SEPARATOR;
	
	// : Variables
	protected $_fileName;
	protected $_data;
	protected $_errors = array ();
	
	// : Public functions
	// : Accessors
	
	/**
	 * RatesReadXLSData::getFileName()
	 *
	 * @return string: $this->_FileName
	 */
	public function getFileName() {
		return $this->_fileName;
	}
	
	/**
	 * RatesReadXLSData::setFileName($_setFile)
	 *
	 * @param string: $_setFile        	
	 */
	public function setFileName($_setFile) {
		$this->_fileName = $_setFile;
	}
	
	/**
	 * RatesReadXLSData::getData()
	 *
	 * @param string: $this->_data;        	
	 */
	public function getData() {
		return $this->_data;
	}
	
	/**
	 * RatesReadXLSData::getErrors
	 *
	 * @param string: $this->_errors        	
	 */
	public function getErrors() {
		return $this->_errors;
	}
	
	/**
	 * RatesReadXLSData::getPoints()
	 * Return array from supplied array with filtered data only
	 *
	 * @param array: $this->getData()        	
	 */
	public function getPoints() {
		$locations = $this->getData ();
		$routes = ( array ) array ();
		if (count ( $locations ["Points"] ) != 0) {
			foreach ( $locations ["Points"] as $key => $values ) {
				if (count ( $values ) != 0) {
					$x = 1;
					foreach ( $values as $keyvalue => $value ) {
						if ($keyvalue != "1") {
							switch ($key) {
								case "A" :
									if ($value != NULL || $value != "") {
										$routes ["LocationFrom"] [$x] = $value;
									}
									break;
								case "B" :
									if ($value != NULL || $value != "") {
										$routes ["LocationTo"] [$x] = $value;
									}
							}
							$x ++;
						}
					}
				}
			}
		}
		return $routes;
	}
	
	/**
	 * RatesReadXLSData::getCities()
	 * Return array from supplied array with filtered data only
	 *
	 * @param array: $this->getData()        	
	 */
	public function getCities() {
		$locations = $this->getData ();
		$cities = array ();
		$rates = $this->getRoutes ();
		$products = $this->getProducts ();
		foreach ( $locations ["Rates"] ["A"] as $key => $value ) {
			if ($key != "1") {
				foreach ( $products as $product ) {
					if ($rates [$value] [$product] != "") {
						$cities [] = $value . "kms Zone " . $product;
					}
				}
			}
		}
		return $cities;
	}
	
	/**
	 * RatesReadXLSData::getSettings()
	 * Return array from supplied array with script data only
	 *
	 * @param array: $this->getData()        	
	 * @param array: $_settings        	
	 */
	public function getSettings() {
		$_data = $this->getData ();
		$_settings = array ();
		foreach ( $_data ["Script"] as $_setting ) {
			$_settings [$_setting [1]] = $_setting [2];
		}
		return $_settings;
	}
	
	/**
	 * RatesReadXLSData::getProducts()
	 * Return array containing products using data from spreadsheet
	 *
	 * @param array: $this->getData()        	
	 */
	public function getProducts() {
		$rates = $this->getData ();
		if (count ( $rates ["Rates"] != 0 )) {
			$products = array ();
			foreach ( $rates ["Rates"] as $key => $values ) {
				if ((count ( $values ) != 0) and ($key != "A")) {
					foreach ( $values as $keyvalue => $value ) {
						if ($keyvalue == "1") {
							$products [] = $value;
						}
					}
				}
			}
			return $products;
		}
	}
	
	/**
	 * RatesReadXLSData::getRoutes()
	 * Return array containing every route and its rate using data from the spreadsheet
	 *
	 * @param array: $this->getData()        	
	 */
	public function getRoutes() {
		$locations = $this->getPoints ( "points" );
		$rates = $this->getData ();
		$routes = ( array ) array ();
		$products = $this->getProducts ();
		$test = array ();
		if (count ( $rates ["Rates"] ) != 0) {
			foreach ( $rates ["Rates"] as $keys => $values ) {
				if (count ( $values ) != 0) {
					foreach ( $values as $key => $value ) {
						if ($key != "1") {
							if ($keys == "A") {
								foreach ( $products as $product ) {
									$routes [$value] [$product] = "";
								}
							} else {
								$routes [$rates ["Rates"] ["A"] [$key]] [$values [1]] = $value;
							}
						}
					}
				}
			}
			return $routes;
		}
	}
	
	/**
	 * RatesReadXLSData::readExcelFile($excelFile, $sheetNames)
	 * Read a spreadsheet into memory containing F and V Contract data
	 * and arrange into a multidimensional array
	 *
	 * @param $excelFile, $sheetNames        	
	 */
	public function readExcelFile($excelFile, $sheetNames) {
		// Setup array containing column headings ranging from A TO ZZ
		$alphaA = range ( 'A', 'Z' );
		$alphaVar = range ( 'A', 'Z' );
		foreach ( $alphaA as $valueA ) {
			foreach ( $alphaA as $valueB ) {
				$alphaVar [] = $valueA . $valueB;
			}
		}
		
		foreach ( $sheetNames as $aSheet ) {
			// Create PHPExcel Reader Object
			$inputFileType = PHPExcel_IOFactory::identify ( $excelFile );
			$objReader = PHPExcel_IOFactory::createReader ( $inputFileType );
			$objReader->setLoadSheetsOnly ( $aSheet ); // Reader options
			$objPHPExcel = $objReader->load ( $excelFile ); // Load worksheet into memory
			$worksheet = $objPHPExcel->getActiveSheet ();
			// Read spreadsheet data and store into array
			foreach ( $worksheet->getRowIterator () as $row ) {
				$cellIterator = $row->getCellIterator ();
				$cellIterator->setIterateOnlyExistingCells ( false );
				foreach ( $cellIterator as $cell ) {
					$data [$aSheet] [$cell->getColumn ()] [$cell->getRow ()] = $cell->getValue ();
				}
			}
		}
		return $data;
	}
	
	// : End
	// : Magic
	
	/**
	 * RatesReadXLSData::__construct()
	 * Class constructor
	 */
	public function __construct($file, $sheetnames) {
		// Check if file exists
		if (! file_exists ( $file )) {
			$this->_errors [] = preg_replace ( '/%s/', $file, self::FILE_DOESNT_EXIST );
		} else {
			$this->setFileName ( $file );
			// Type cast an array to store the sheet names to pass to the read excel sheet function
			try {
				// Load spreadsheet into memory
				$this->setData ( $this->readExcelFile ( $this->getFileName (), $sheetnames ) );
			} catch ( Exception $e ) {
				$this->_errors [] = $e->getMessage ();
			}
		}
		if (count ( $this->_errors ) != 0) {
			return true;
		} else {
			return false;
		}
	}
	
	/**
	 * RatesReadXLSData::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End
	
	// : Private Functions
	
	/**
	 * RatesReadXLSData::setData($_var)
	 *
	 * @param string: $_var        	
	 */
	private function setData($_var) {
		$this->_data = $_var;
	}
	// : End
}