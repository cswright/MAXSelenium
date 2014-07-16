<?php
// Error reporting
error_reporting (E_ALL);

// : Includes
require_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';

// : End

/**
 * Object::ReadExcelFile
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Manline Group (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class ReadExcelFile {
	// : Constants
	const DISALLOWED_EXTENSION = "The file %s has an unsupported extension.";
	const FILE_DOESNT_EXIST = "The file %s could not be found.";
	const FILE_COULD_NOT_CREATE = "The file %s could not be created.";
	const FILE_NOT_READABLE = "The file %s could not be read. Probable causes: Incorrect permission or in a path that PHP cannot read.";
	const FOPEN_COULD_NOT_OPEN = "The file %s could not be opened.";
	const DS = DIRECTORY_SEPARATOR;

	// : Variables
	protected $_fileName;
	protected $_data;
	protected $_errors = array();

	// : Public functions
	// : Accessors

	/**
	 * ReadExcelFile::getFileName()
	 *
	 * @return string: $this->_FileName
	*/
	public function getFileName() {
		return $this->_fileName;
	}

	/**
	 * ReadExcelFile::setFileName($_setFile)
	 *
	 * @param string: $_setFile
	 */
	public function setFileName($_setFile) {
		$this->_fileName = $_setFile;
	}

	/**
	 * ReadExcelFile::getData()
	 *
	 * @param string: $this->_data;
	 */
	public function getData() {
		return $this->_data;
	}

	/**
	 * ReadExcelFile::getErrors
	 *
	 * @param string: $this->_errors
	 */
	public function getErrors()
	{
		return $this->_errors;
	}

	/**
	 * ReadExcelFile::readExcelFile($excelFile, $sheetName)
	 * Read a spreadsheet into memory containing F and V Contract data
	 * and arrange into a multidimensional array
	 *
	 * @param $excelFile, $sheetName
	 */
	public function readExcelFileData($excelFile, $sheetname) {
		try {
			// Type cast necessary variables
			$contractData = ( array ) array ();
			// Create PHPExcel Reader Object
			$inputFileType = PHPExcel_IOFactory::identify($excelFile);
			$objReader = PHPExcel_IOFactory::createReader($inputFileType);
			$objReader->setReadDataOnly(true);
			$objPHPExcel = $objReader->load($excelFile); // Load worksheet into memory
			$worksheet = $objPHPExcel->getSheetByName($sheetname);
			
			// Read spreadsheet data and store into array
			foreach ($worksheet->getRowIterator() as $row) {
				$cellIterator = $row->getCellIterator();
				$cellIterator->setIterateOnlyExistingCells(true);
				foreach ($cellIterator as $cell) {
					if ($cell->getRow() != 1) { 
						$data[$objPHPExcel->getActiveSheet()->getCell($cell->getColumn() . "1")->getValue()] [$cell->getRow() - 1] = $cell->getValue();
					}
				}
			}
			return $data;
		} catch ( Exception $e ) {
			echo "Caught exception: ", $e->getMessage (), "\n";
			return false;
		}
	}

	// : Magic
	/**
	* ReadExcelFile::__construct()
	* Class constructor
	*/
	public function __construct($file, $sheet) {
		// Check if file exists
		if (!file_exists($file)) {
			$this->_errors[] = preg_replace('/%s/', $file, self::FILE_DOESNT_EXIST);
		} else {
			// Get the extension of the filename
			$getExt = preg_split("/.*\./", $file);
			if (count($getExt) != 0)  {
				// Check that the file extension is a valid excel document
				if (($getExt[1] = "xls") || ($getExt[1] = "xlsx")) {
					$this->setFileName($file); # Set the filename
					// Try open the excel file and return data in an array
					try {
						// Set the data to return as the data from the excel spreadsheet
						$this->setData($this->readExcelFileData($this->getFileName(), $sheet));
					} catch(Exception $e) {
						$this->_errors[] = preg_replace('/%s/', $file, self::FOPEN_COULD_NOT_OPEN) . "\n" . $e->getMessage(); 
					}
				} else {
					$this->_errors[] = preg_replace('/%s/', $file, self::DISALLOWED_EXTENSION);
				}
			}
		}
		// Check if read of excel document was successful
		if (count($this->_errors) != 0) {
			return false;
		} else {
			return true;
		}
	}

	/**
	 * ReadExcelFile::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End

	// : Private Functions

	/**
	 * ReadExcelFile::setData($_var)
	 *
	 * @param string: $_var
	 */
	public function setData($_var) {
		$this->_data = $_var;
	}

	// : End
}