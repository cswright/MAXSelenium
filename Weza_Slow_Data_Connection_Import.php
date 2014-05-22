<?php
include_once 'PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriver.php';
include_once 'PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverWait.php';
include_once 'PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverBy.php';
include_once dirname ( __FILE__ ) . '/RatesReadXLSData.php';
include_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';
/**
 * PHPExcel_Writer_Excel2007
 */
include 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel/Writer/Excel2007.php';

/**
 * Object::Weza_Slow_Data_Connection_Import
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Barloworld Transport Solutions (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class Weza_Slow_Data_Connection_Import extends PHPUnit_Framework_TestCase {
	// : Constants
	const COULD_NOT_CONNECT_MYSQL = "Failed to connect to MySQL database";
	const MAX_NOT_RESPONDING = "Error: MAX does not seem to be responding";
	const PB_URL = "/Planningboard";
	const CUSTOMER_URL = "/DataBrowser?browsePrimaryObject=461&browsePrimaryInstance=";
	const LOCATION_BU_URL = "/DataBrowser?browsePrimaryObject=495&browsePrimaryInstance=";
	const OFF_CUST_BU_URL = "/DataBrowser?browsePrimaryObject=494&browsePrimaryInstance=";
	const RATEVAL_URL = "/DataBrowser?&browsePrimaryObject=udo_Rates&browsePrimaryInstance=%s&browseSecondaryObject=DateRangeValue&relationshipType=Rate";
	const DS = DIRECTORY_SEPARATOR;
	const BUNIT = "Freight";
	const CONTRIB = "Energy (tankers)";
	const COUNTRY = "South Africa";
	const CUSTOMER = "NCP Chlorochem - Chlorine";
	const PROVINCE = "Africa -- South Africa -- KZN";
	const TRUCKTYPE = "Fuel Tanker";
	const LIVE_URL = "https://login.max.bwtsgroup.com";
	const TEST_URL = "http://max.mobilize.biz";
	const INI_FILE = "rates_data.ini";
	const INI_DIR = "ini";
	const TEST_SESSION = "firefox";
	const XLS_CREATOR = "Weza_Slow_Data_Connection_Import.php";
	const XLS_TITLE = "Error Report";
	const XLS_SUBJECT = "Errors caught while creating rates for subcontracts";

	// : Variables
	protected static $driver;
	protected $_dummy;
	protected $_session;
	protected $lastRecord;
	protected $to = 'clintonabco@gmail.com';
	protected $subject = 'MAX Selenium script report';
	protected $message;
	protected $_username;
	protected $_password;
	protected $_welcome;
	protected $_mode;
	protected $_dataDir;
	protected $_errDir;
	protected $_scrDir;
	protected $_maxurl;
	protected $_error = array ();
	protected $_db;
	protected $_dbdsn = "mysql:host=%s;dbname=max2;charset=utf8;";
	protected $_dbuser = "root";
	protected $_dbpwd = "kaluma";
	protected $_dboptions = array (
			PDO::MYSQL_ATTR_INIT_COMMAND => 'SET NAMES utf8',
			PDO::ATTR_EMULATE_PREPARES => false,
			PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
			PDO::ATTR_PERSISTENT => true
	);
	protected $_myqueries = array (
			"select ID from udo_customerlocations where location_id IN (select ID from udo_location where name='%n' and _type='udo_Point') and customer_id IN (select ID from udo_customer where tradingName='%t');",
			"select ID from udo_offloadingcustomers where offloadingCustomer_id IN (select ID from udo_customer where tradingName='%o') and customer_id IN (select ID from udo_customer where tradingName='%t');",
			"select ID from udo_customer where tradingName='%t';",
			"select ID from udo_rates where route_id IN (select ID from udo_route where locationTo_id IN (select ID from udo_location where name='%t')) and objectregistry_id=%g and objectInstanceId=%c and truckDescription_id=%d and enabled=1 and model='%m' and businessUnit_id=%b and rateType_id=%r;",
			"select ID from objectregistry where handle = 'udo_Customer';"
	);

	// : Public functions
	// : Accessors

	// : End

	// : Magic
	/**
	* Weza_Slow_Data_Connection_Import::__construct()
	* Class constructor
	*/
	public function __construct() {
		$ini = dirname ( realpath ( __FILE__ ) ) . self::DS . self::INI_DIR . self::DS . self::INI_FILE;
		if (is_file ( $ini ) === FALSE) {
			echo "No " . self::INI_FILE . " file found. Please refer to documentation for script to determine which fields are required and their corresponding values." . PHP_EOL;
			return FALSE;
		}
		$data = parse_ini_file ( $ini );
		if ((array_key_exists ( "xls", $data ) && $data ["xls"]) && (array_key_exists ( "errordir", $data ) && $data ["errordir"]) && (array_key_exists ( "screenshotdir", $data ) && $data ["screenshotdir"]) && (array_key_exists ( "datadir", $data ) && $data ["datadir"]) && (array_key_exists ( "ip", $data ) && $data ["ip"]) && (array_key_exists ( "username", $data ) && $data ["username"]) && (array_key_exists ( "password", $data ) && $data ["password"]) && (array_key_exists ( "welcome", $data ) && $data ["welcome"]) && (array_key_exists ( "mode", $data ) && $data ["mode"])) {
			$this->_username = $data ["username"];
			$this->_password = $data ["password"];
			$this->_welcome = $data ["welcome"];
			$this->_dataDir = $data ["datadir"];
			$this->_errDir = $data ["errordir"];
			$this->_scrDir = $data ["screenshotdir"];
			$this->_mode = $data ["mode"];
			$this->_ip = $data ["ip"];
			$this->_xls = $data ["xls"];
			switch ($this->_mode) {
				case "live" :
					$this->_maxurl = self::LIVE_URL;
					break;
				default :
					$this->_maxurl = self::TEST_URL;
			}
		} else {
			echo "The correct data is not present in " . self::INI_FILE . ". Please confirm. Fields are username, password, welcome and mode" . PHP_EOL;
			return FALSE;
		}
	}

	/**
	 * Weza_Slow_Data_Connection_Import::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End

	/**
	 * Weza_Slow_Data_Connection_Import::setUp()
	 * Create new class object and initialize session for webdriver
	 */
	public function setUp() {
		self::$driver = new PHPWebDriver_WebDriver ();
		$this->_session = self::$driver->session ( self::TEST_SESSION );
	}

	/**
	 * Weza_Slow_Data_Connection_Import::testCreateContracts()
	 * Pull F and V Contract data and automate creation of F and V Contracts
	 */
	public function testCreateContracts() {
		$_sheetnames = ( array ) array (
				"Points",
				"Rates",
				"Script"
		);
		// : Pull data from correctly formatted xls spreadsheet
		if ($cPR = new RatesReadXLSData ( dirname ( __FILE__ ) . $this->_dataDir . self::DS . $this->_xls, $_sheetnames )) {
			// Get cities and save in correct naming format standard as per Meryle instruction
			$cities = $cPR->getCities ();
			// Get script data settings
			$settings = $cPR->getSettings ();
				
			try {
				// Initiate Session
				$session = $this->_session;
				$this->_session->setPageLoadTimeout ( 60 );
				$w = new PHPWebDriver_WebDriverWait ( $this->_session );

				// : Extract columns from the spreadsheet data
				$_xlsColumns = array (
						"Error_Msg",
						"Record Detail"
				);

				// : Setup local variables
				$_bu = $settings ["BusinessUnit"];
				$_customer = $settings ["Customer"];
				$_contrib = $settings ["ContribModel"];
				$_truckType = $settings ["TruckType"];
				$_startDate = $settings ["StartDate"];
				$_endDate = $settings ["EndDate"];
				$_rateType = $settings ["RateType"];

				// Insert IP address for MySQL Server supplied in rates_data.ini
				$_mysqlDsn = preg_replace ( "/%s/", $this->_ip, $this->_dbdsn );
				// Open keepalive connection to database
				$this->openDB ( $_mysqlDsn, $this->_dbuser, $this->_dbpwd, $this->_dboptions );

				// Get truck description ID
				$myQuery = "select ID from udo_truckdescription where description='$_truckType';";
				$result = $this->queryDB ( $myQuery );
				if (count ( $result ) != 0) {
					$trucktype_id = $result [0] ["ID"];
				} else {
					throw new Exception ( "Error: Truck description not found. Please check and amend truck description." );
				}

				// Get customer ID
				$myQuery = "select ID from udo_customer where tradingName='$_customer';";
				$result = $this->queryDB ( $myQuery );
				if (count ( $result ) != 0) {
					$customer_id = $result [0] ["ID"];
				} else {
					throw new Exception ( "Error: Customer not found. Please check and amend customer name." );
				}

				// Get rate type ID
				$myQuery = "select ID from udo_ratetype where name='$_rateType';";
				$result = $this->queryDB ( $myQuery );
				if (count ( $result ) != 0) {
					$rateType_id = $result [0] ["ID"];
				} else {
					throw new Exception ( "Error: Rate type not found. Please check and amend rate type name." );
				}

				// Get business unit ID
				$myQuery = "select ID from udo_businessunit where name='$_bu';";
				$result = $this->queryDB ( $myQuery );
				if (count ( $result ) != 0) {
					$bunit_id = $result [0] ["ID"];
				} else {
					throw new Exception ( "Error: Business unit not found. Please check and amend business unit name." );
				}

				// Get objectregistry_id for udo_Customer
				$myQuery = $this->_myqueries [4];
				$result = $this->queryDB ( $myQuery );
				if (count ( $result ) != 0) {
					$objectregistry_id = $result [0] ["ID"];
				} else {
					throw new Exception ( "Error: Object registry record for udo_customer not found." );
				}
				// : End
			} catch ( Exception $e ) {
				// Print error message
				print ($e->getMessage () . PHP_EOL) ;
				// Terminate application
				die ();
			}
				
			// : Login
			try {
				$this->_session->open ( $this->_maxurl );
				// : Wait for page to load and for elements to be present on page
				if ($this->_mode == "live") {
					$e = $w->until ( function ($session) {
						return $session->element ( 'css selector', "#contentFrame" );
					} );
					$iframe = $this->_session->element ( 'css selector', '#contentFrame' );
					$this->_session->switch_to_frame ( $iframe );
				}
				$e = $w->until ( function ($session) {
					return $session->element ( 'css selector', 'input[id=identification]' );
				} );
				// : End
				$this->assertElementPresent ( 'css selector', 'input[id=identification]' );
				$this->assertElementPresent ( 'css selector', 'input[id=password]' );
				$this->assertElementPresent ( 'css selector', 'input[name=submit][type=submit]' );
				$e->sendKeys ( $this->_username );
				$e = $this->_session->element ( 'css selector', 'input[id=password]' );
				$e->sendKeys ( $this->_password );
				$e = $this->_session->element ( 'css selector', 'input[name=submit][type=submit]' );
				$e->click ();
				// Switch out of frame
				if ($this->_mode == "live") {
					$this->_session->switch_to_frame ();
				}

				// : Wait for page to load and for elements to be present on page
				if ($this->_mode == "live") {
					$e = $w->until ( function ($session) {
						return $session->element ( 'css selector', "#contentFrame" );
					} );
					$iframe = $this->_session->element ( 'css selector', '#contentFrame' );
					$this->_session->switch_to_frame ( $iframe );
				}
				$e = $w->until ( function ($session) {
					return $session->element ( "xpath", "//*[text()='" . $this->_welcome . "']" );
				} );
				$this->assertElementPresent ( "xpath", "//*[text()='" . $this->_welcome . "']" );
				// Switch out of frame
				if ($this->_mode == "live") {
					$this->_session->switch_to_frame ();
				}
			} catch ( Exception $e ) {
				throw new Exception ( "Error: Failed to log into MAX." . PHP_EOL . $e->getMessage () );
			}
			// : End
				
			// : Load Planningboard to rid of iframe loading on every page from here on
			$this->_session->open ( $this->_maxurl . self::PB_URL );
			$e = $w->until ( function ($session) {
				return $session->element ( "xpath", "//*[contains(text(),'You Are Here') and contains(text(), 'Planningboard')]" );
			} );
			// : End
					
				// : Create Routes, Rates and Rate Values
				foreach ( $cities as $pointname ) {
					try {
						$this->lastRecord = $pointname;
						// : Get kms zone for this entry
						$kms = preg_split ( "/kms Zone.*/", $pointname );
						$kms = $kms [0];
						// : End
							
						// Correct hyphen conversion issue with spreadsheets
						$pointname = preg_replace ( "/–/", "-", $pointname );
							
						// : Create Rate Value for Route
						$myQuery = preg_replace ( "/%t/", $pointname, $this->_myqueries [3] );
						$myQuery = preg_replace ( "/%g/", $objectregistry_id, $myQuery );
						$myQuery = preg_replace ( "/%c/", $customer_id, $myQuery );
						$myQuery = preg_replace ( "/%d/", $trucktype_id, $myQuery );
						$myQuery = preg_replace ( "/%m/", $_contrib, $myQuery );
						$myQuery = preg_replace ( "/%b/", $bunit_id, $myQuery );
						$myQuery = preg_replace ( "/%r/", $rateType_id, $myQuery );
						$result = $this->queryDB ( $myQuery );
						if (count ( $result ) != 0) {
							foreach ( $result as $_rateRecord ) {
								$rate_id = $_rateRecord ["ID"];
								$rateurl = preg_replace ( "/%s/", $rate_id, $this->_maxurl . self::RATEVAL_URL );
								$this->_session->open ( $rateurl );
									
								// Wait for element = #button-create
								$e = $w->until ( function ($session) {
									return $session->element ( "css selector", "#button-create" );
								} );
								// Click element - #button-create
								$this->_session->element ( "css selector", "#button-create" )->click ();
									
								// Wait for element = #button-create
								$e = $w->until ( function ($session) {
									return $session->element ( "xpath", "//*[contains(text(),'Create Date Range Values')]" );
								} );
										
									$this->assertElementPresent ( "xpath", "//*[@id='DateRangeValue-2_0_0_beginDate-2']" );
									$this->assertElementPresent ( "xpath", "//*[@id='DateRangeValue-4_0_0_endDate-4']" );
									$this->assertElementPresent ( "xpath", "//*[@id='DateRangeValue-20_0_0_value-20']" );
									$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
										
									// Clear the begin date text field
									$this->_session->element ( "xpath", "//*[@id='DateRangeValue-2_0_0_beginDate-2']" )->clear ();
									// Paste startDate into begin date field
									$this->_session->element ( "xpath", "//*[@id='DateRangeValue-2_0_0_beginDate-2']" )->sendKeys ( $_startDate );
									// Clear the end date text field
									$this->_session->element ( "xpath", "//*[@id='DateRangeValue-4_0_0_endDate-4']" )->clear ();
									// Paste endDate into end date field
									$this->_session->element ( "xpath", "//*[@id='DateRangeValue-4_0_0_endDate-4']" )->sendKeys ( $_endDate );
									// Get the product name out the string
									$productname = preg_split ( "/^" . $kms . "kms Zone /", $pointname );
									// Format the string of the rate value xxx.xx
									$ratevalue = strval ( (number_format ( floatval ( $routes [$kms] [$productname [1]] ), 2, ".", "" )) );
									// Paste the formatted rate value into the value field
									$this->_session->element ( "xpath", "//*[@id='DateRangeValue-20_0_0_value-20']" )->sendKeys ( $ratevalue );
									// Click element - submit button
									$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							}
						} else {
							throw new Exception ( "Error: Rate id record not found." );
						}
					} catch ( Exception $e ) {
						echo "Error: " . $e->getMessage () . PHP_EOL;
						echo "Time of error: " . date ( "Y-m-d H:i:s" ) . PHP_EOL;
						echo "Last record: " . $this->lastRecord;
						$this->takeScreenshot ();
						$_erCount = count ( $this->_error );
						$this->_error [$_erCount + 1] ["error"] = $e->getMessage ();
						$this->_error [$_erCount + 1] ["record"] = $this->lastRecord;
					}
				}
				// : End
				// : End
					
				// : Tear Down
				$this->_session->element ( 'xpath', "//*[contains(@href,'/logout')]" )->click ();
				// Wait for page to load and for elements to be present on page
				$e = $w->until ( function ($session) {
					return $session->element ( 'css selector', 'input[id=identification]' );
				} );
				$this->assertElementPresent ( 'css selector', 'input[id=identification]' );
				$db = null;
				$this->_session->close ();
				// : End
				// : If errors occured. Create xls of entries that failed.
				if (count ( $this->_error ) != 0) {
					$_xlsfilename = (dirname ( __FILE__ ) . $this->_errDir . self::DS . date ( "Y-m-d_His_" ) . "MAXLiveNCP_" . ".xlsx");
					$this->writeExcelFile ( $_xlsfilename, $this->_error, $_xlsColumns );
					if (file_exists ( $_xlsfilename )) {
						print ("Excel error report written successfully to file: $_xlsfilename") ;
					} else {
						print ("Excel error report write unsuccessful") ;
					}
				}
				// : End
		} else {
			print ("Error: The excel spreadsheet, '" . $this->_xls . "', failed to load." . PHP_EOL) ;
		}
	}

	// : Private Functions

	/**
	 * MAXLive_Subcontractors::writeExcelFile($excelFile, $excelData)
	 * Create, Write and Save Excel Spreadsheet from collected data obtained from the variance report
	 *
	 * @param $excelFile, $excelData
	 */
	public function writeExcelFile($excelFile, $excelData, $columns) {
		try {
			// Check data validility
			if (count ( $excelData ) != 0) {

				// : Create new PHPExcel object
				print ("<pre>") ;
				print (date ( 'H:i:s' ) . " Create new PHPExcel object" . PHP_EOL) ;
				$objPHPExcel = new PHPExcel ();
				// : End

				// : Set properties
				print (date ( 'H:i:s' ) . " Set properties" . PHP_EOL) ;
				$objPHPExcel->getProperties ()->setCreator ( self::XLS_CREATOR );
				$objPHPExcel->getProperties ()->setLastModifiedBy ( self::XLS_CREATOR );
				$objPHPExcel->getProperties ()->setTitle ( self::XLS_TITLE );
				$objPHPExcel->getProperties ()->setSubject ( self::XLS_SUBJECT );
				// : End

				// : Setup Workbook Preferences
				print (date ( 'H:i:s' ) . " Setup workbook preferences" . PHP_EOL) ;
				$objPHPExcel->getDefaultStyle ()->getFont ()->setName ( 'Arial' );
				$objPHPExcel->getDefaultStyle ()->getFont ()->setSize ( 8 );
				$objPHPExcel->getActiveSheet ()->getPageSetup ()->setOrientation ( PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE );
				$objPHPExcel->getActiveSheet ()->getPageSetup ()->setPaperSize ( PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4 );
				$objPHPExcel->getActiveSheet ()->getPageSetup ()->setFitToWidth ( 1 );
				$objPHPExcel->getActiveSheet ()->getPageSetup ()->setFitToHeight ( 0 );
				// : End

				// : Set Column Headers
				$alphaVar = range ( 'A', 'Z' );
				print (date ( 'H:i:s' ) . " Setup column headers" . PHP_EOL) ;

				$i = 0;
				foreach ( $columns as $key ) {
					$objPHPExcel->getActiveSheet ()->setCellValue ( $alphaVar [$i] . "1", $key );
					$objPHPExcel->getActiveSheet ()->getStyle ( $alphaVar [$i] . '1' )->getFont ()->setBold ( true );
					$i ++;
				}

				// : End

				// : Add data from $excelData array
				print (date ( 'H:i:s' ) . " Add data from error array" . PHP_EOL) ;
				$rowCount = ( int ) 2;
				$objPHPExcel->setActiveSheetIndex ( 0 );
				foreach ( $excelData as $values ) {
					$i = 0;
					foreach ( $values as $key => $value ) {
						$objPHPExcel->getActiveSheet ()->getCell ( $alphaVar [$i] . strval ( $rowCount ) )->setValueExplicit ( $value, PHPExcel_Cell_DataType::TYPE_STRING );
						$i ++;
					}
					$rowCount ++;
				}
				// : End

				// : Setup Column Widths
				for($i = 0; $i <= count ( $columns ); $i ++) {
					$objPHPExcel->getActiveSheet ()->getColumnDimension ( $alphaVar [$i] )->setAutoSize ( true );
				}
				// : End

				// : Rename sheet
				print (date ( 'H:i:s' ) . " Rename sheet" . PHP_EOL) ;
				$objPHPExcel->getActiveSheet ()->setTitle ( self::XLS_TITLE );
				// : End

				// : Save spreadsheet to Excel 2007 file format
				print (date ( 'H:i:s' ) . " Write to Excel2007 format" . PHP_EOL) ;
				print ("</pre>" . PHP_EOL) ;
				$objWriter = new PHPExcel_Writer_Excel2007 ( $objPHPExcel );
				$objWriter->save ( $excelFile );
				$objPHPExcel->disconnectWorksheets ();
				unset ( $objPHPExcel );
				unset ( $objWriter );
				// : End
			} else {
				print ("<pre>") ;
				print_r ( "ERROR: The function was passed an empty array" );
				print ("</pre>") ;
				exit ();
			}
		} catch ( Exception $e ) {
			echo "Caught exception: ", $e->getMessage (), "\n";
			exit ();
		}
	}

	/**
	 * Weza_Slow_Data_Connection_Import::openDB($dsn, $username, $password, $options)
	 * Open connection to Database
	 *
	 * @param string: $dsn
	 * @param string: $username
	 * @param string: $password
	 * @param array: $options
	 */
	private function openDB($dsn, $username, $password, $options) {
		try {
			$this->_db = new PDO ( $dsn, $username, $password, $options );
		} catch ( PDOException $ex ) {
			return FALSE;
		}
	}

	/**
	 * MAXLive_Subcontractors::takeScreenshot()
	 * This is a function description for a selenium test function
	 *
	 * @param object: $_session
	 */
	private function takeScreenshot() {
		$_img = $this->_session->screenshot ();
		$_data = base64_decode ( $_img );
		$_file = dirname ( __FILE__ ) . $this->_scrDir . DIRECTORY_SEPARATOR . date ( "Y-m-d_His" ) . "_WebDriver.png";
		$_success = file_put_contents ( $_file, $_data );
		if ($_success) {
			return $_file;
		} else {
			return FALSE;
		}
	}

	/**
	 * Weza_Slow_Data_Connection_Import::assertElementPresent($_using, $_value)
	 * This is a function description for a selenium test function
	 *
	 * @param string: $_using
	 * @param string: $_value
	 */
	private function assertElementPresent($_using, $_value) {
		$e = $this->_session->element ( $_using, $_value );
		$this->assertEquals ( count ( $e ), 1 );
	}

	/**
	 * Weza_Slow_Data_Connection_Import::assertElementPresent($_title)
	 * This functions switches focus between each of the open windows
	 * and looks for the first window where the page title matches
	 * the given title and returns true else false
	 *
	 * @param string: $_title
	 * @param
	 *        	boolean: return
	 */
	private function selectWindow($_title) {
		try {
			$_results = ( array ) array ();
			// Store the current window handle value
			$_currentWin = $this->_session->window_handle ();
			// Get all open windows handles
			$e = $this->_session->window_handles ();
			if (count ( $e ) > 1) {
				foreach ( $e as $_browserWindow ) {
					$this->_session->focusWindow ( $_browserWindow );
					$_page_title = $this->_session->title ();
					preg_match ( "/^.+" . $_title . ".+/", $_page_title, $_results );
					if ((count ( $_results ) != 0) && ($_browserWindow != $_currentWin)) {
						return true;
					}
				}
			}
			$this->_session->focusWindow ( $_currentWin );
			return false;
		} catch ( Exception $e ) {
			return false;
		}
	}

	/**
	 * Weza_Slow_Data_Connection_Import::closeDB()
	 * Close connection to Database
	 */
	private function closeDB() {
		$this->_db = null;
	}

	/**
	 * Weza_Slow_Data_Connection_Import::queryDB($sqlquery)
	 * Pass MySQL Query to database and return output
	 *
	 * @param string: $sqlquery
	 * @param array: $result
	 */
	private function queryDB($sqlquery) {
		try {
			$result = $this->_db->query ( $sqlquery );
			return $result->fetchAll ( PDO::FETCH_ASSOC );
		} catch ( PDOException $ex ) {
			return FALSE;
		}
	}

	// : End
}