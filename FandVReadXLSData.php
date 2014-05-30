<?php
// Error reporting
error_reporting (E_ALL);

// : Includes

// : End

/**
 * Object::FandVReadXLSData
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Manline Group (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class FandVReadXLSData {
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
	protected $_errors = array();

	// : Public functions
	// : Accessors

	/**
	 * FandVReadXLSData::getFileName()
	 *
	 * @return string: $this->_FileName
	 */
	public function getFileName() {
		return $this->_fileName;
	}

	/**
	 * FandVReadXLSData::setFileName($_setFile)
	 *
	 * @param string: $_setFile
	 */
	public function setFileName($_setFile) {
		$this->_fileName = $_setFile;
	}
	
	/**
	 * FandVReadXLSData::getData()
	 *
	 * @param string: $this->_data;
	 */
	public function getData() {
		return $this->_data;
	}
	
	/**
	 * FandVReadXLSData::getErrors
	 *
	 * @param string: $this->_errors
	 */	
	public function getErrors()
	{
		return $this->_errors;
	}

	/**
	 * FandVReadXLSData::readExcelFile($excelFile, $sheetName)
	 * Read a spreadsheet into memory containing F and V Contract data
	 * and arrange into a multidimensional array
	 *
	 * @param $excelFile, $sheetName
	 */
	public function readExcelFile($excelFile, $sheetName) {
		try {
			$alphaA = range ( 'A', 'Z' );
			$alphaVar = range ( 'A', 'Z' );
			foreach ( $alphaA as $valueA ) {
				foreach ( $alphaA as $valueB ) {
					$alphaVar [] = $valueA . $valueB;
				}
			}
			// Construct multidimensional array template for contract data
			$dataTemplate = ( array ) array (
					"Contract" => "",
					"Business Unit" => "",
					"Customer" => "",
					"Contrib" => "",
					"Cost" => "",
					"Start Date" => "",
					"End Date" => "",
					"Days" => 0,
					"Rate" => "",
					"Truck Type" => "",
					"Trucks Linked" => "",
					"Routes Linked" => "",
					"RateType" => "",
					"DaysPerMonth" => "",
					"DaysPerTrip" => "",
					"FuelConsumption" => "",
					"FleetValues" => "",
					"ExpectedEmptyKms" => "",
					"ExpectedDistance" => "",
					"ContractId" => 0
			);
			// Construct multidimensional array template for storing links
			$storeComments = (array) array(
					"cellID" => "",
					"comment" => "",
					"contract" => "",
					"type" => ""
			);
			// Type cast necessary variables
			$contractData = ( array ) array ();
			// Create PHPExcel Reader Object
			$inputFileType = PHPExcel_IOFactory::identify($excelFile);
			$objReader = PHPExcel_IOFactory::createReader($inputFileType);
			//$objReader->setLoadSheetsOnly($sheetName); // Reader options
			$objPHPExcel = $objReader->load($excelFile); // Load worksheet into memory
			$worksheet = $objPHPExcel->getActiveSheet();
			// Read spreadsheet data and store into array
			foreach ($worksheet->getRowIterator() as $row) {
				$cellIterator = $row->getCellIterator();
				$cellIterator->setIterateOnlyExistingCells(false);
				foreach ($cellIterator as $cell) {
					$data[$cell->getColumn()] [$cell->getRow()] = $cell->getValue();
				}
			}
			
			// Get all comments within worksheet
			$comments = $worksheet->getComments(); $x = 0;
			// Filter each comment by truck and route link and store into array
			foreach($comments  as $cellID => $comment) {
				preg_match("#.*[a-z]#i", $cellID, $getCol);
				$contract = $worksheet->getCell($getCol[0] . "1")->getValue();
				$storeComments[$contract]["cellID"] = $cellID;
				preg_match("#[0-9]*$#", $cellID, $a);
				switch ($a[0]) {
					case "11":
						$storeComments[$contract]["Trucks"] = $comment->getText()->getPlainText();
						break;
					case "12":
						$storeComments[$contract]["Routes"] = $comment->getText()->getPlainText();;
						break;
				}
				$x++;
			}
			// Get last array with data
			$lastValue = count($data["A"]); $x = (int) 0;
			foreach ($data as $key => $value) {
				if (($value[1] != "") and ($key != "A")) {
					$contractData[$x] = $dataTemplate;
					$contractData[$x]["Contract"] = $value[1];
					$contractData[$x]["Customer"] = $value[2];
					$contractData[$x]["Contrib"] = $value[3];
					$contractData[$x]["Cost"] = $value[4];
					$contractData[$x]["Days"] = $value[5];
					$contractData[$x]["Rate"] = $value[6];
					$contractData[$x]["Business Unit"] = $value[7];
					$contractData[$x]["Start Date"] = $value[8];
					$contractData[$x]["End Date"] = $value[9];
					$contractData[$x]["Truck Type"] = $value[10];
					$contractData[$x]["Trucks Linked"] = $value[11];
					$contractData[$x]["Routes Linked"] = $value[12];
					$contractData[$x]["RateType"] = $value[13];
					$contractData[$x]["DaysPerMonth"] = $value[14];
					$contractData[$x]["DaysPerTrip"] = $value[15];
					$contractData[$x]["FuelConsumption"] = $value[16];
					$contractData[$x]["FleetValues"] = $value[17];
					$contractData[$x]["ExpectedEmptyKms"] = $value[18];
					$contractData[$x]["ExpectedDistance"] = $value[19];
					if (array_key_exists("20", $value)) {
						$contractData[$x]["ContractId"] = $value[20];
					}
					$searchArray = array($value[1], "Trucks");
					if ($value[11] != 0) {
						if (isset($storeComments[$value[1]]) != FALSE) {
							if (array_key_exists("Trucks", $storeComments[$value[1]]) != FALSE) {
								$contractData[$x]["Trucks Linked"] = $storeComments[$value[1]]["Trucks"];
							}
						}
					}
					$searchArray = array($value[1], "Routes");
					if ($value[12] != 0) {
						if (isset($storeComments[$value[1]]) != FALSE) {
							if (array_key_exists("Routes", $storeComments[$value[1]]) != FALSE) {
								$contractData[$x]["Routes Linked"] = $storeComments[$value[1]]["Routes"];
							}
						}
					}
					$x++;
				}
			}
			return $contractData;
		} catch ( Exception $e ) {
			echo "Caught exception: ", $e->getMessage (), "\n";
			exit ();
		}
	}

	// : Magic
	/**
	* FandVReadXLSData::__construct()
	* Class constructor
	*/
	public function __construct($file) {
		// Check if file exists
		if (!file_exists($file)) {
			$this->_errors[] = preg_replace('/%s/', $file, self::FILE_DOESNT_EXIST);
		} else {
			$this->setFileName($file);
			try {
				// Attempt to load spreadsheet into multidimensional formatted array
				$this->setData($this->readExcelFile($this->getFileName(), date("F Y")));
			} catch(Exception $e) {
				echo "Caught exception: ", $e->getMessage(), "\n";
			}
		}
		if (count($this->_errors) != 0) {
			return true;
		} else {
			return false;
		}
	}

	/**
	 * FandVReadXLSData::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End

	// : Private Functions
	
	/**
	 * FandVReadXLSData::setData($_var)
	 *
	 * @param string: $_var
	 */
	public function setData($_var) {
		$this->_data = $_var;
	}

	// : End
}