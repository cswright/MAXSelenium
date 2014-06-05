<?php
// : Includes
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriver.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverWait.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverBy.php');
include_once dirname ( __FILE__ ) . '/ReadExcelFile.php';
include_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';
/**
 * PHPExcel_Writer_Excel2007
 */
include 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel/Writer/Excel2007.php';
// : End

/**
 * Object::MAXLive_Route_Missing_Distance
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Barloworld Transport Solutions (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class MAXLive_Route_Missing_Distance extends PHPUnit_Framework_TestCase {
	// : Constants
	const DS = DIRECTORY_SEPARATOR;
	const PB_URL = "/Planningboard";
	const COULD_NOT_CONNECT_MYSQL = "Failed to connect to MySQL database";
	const MAX_NOT_RESPONDING = "Error: MAX does not seem to be responding";
	const ROUTE_URL = "/DataBrowser?browsePrimaryObject=511&browsePrimaryInstance=";
	const LIVE_URL = "https://login.max.bwtsgroup.com";
	const TEST_URL = "http://max.mobilize.biz";
	const INI_FILE = "user_data.ini";
	const INI_DIR = "ini";
	const TEST_SESSION = "firefox";
	const XLS_CREATOR = "MAXLive_Route_Missing_Distance.php";
	const XLS_TITLE = "Error Report";
	const XLS_SUBJECT = "Error caught while updating routes.";
	
	// : Variables
	protected static $driver;
	protected $_apikey = "AIzaSyBkTsZyk5_LSmDAYakeA5p_usX0WMQg7qM";
	protected $_dummy;
	protected $_session;
	protected $lastRecord;
	protected $to = 'clintonabco@gmail.com';
	protected $subject = 'MAX Selenium script report';
	protected $message;
	protected $_errors = array ();
	protected $_processed = array ();
	protected $_dataDir;
	protected $_maxurl;
	protected $_mode;
	protected $_ip;
	protected $_xls;
	protected $_wdport;
	protected $_username;
	protected $_password;
	protected $_welcome;
	protected $_browser;
	protected $_db;
	protected $_dbdsn = "mysql:host=%s;dbname=max2;charset=utf8;";
	protected $_dbuser = "root";
	protected $_dbpwd = "kaluma";
	protected $_cURLTimeout = 1000;
	protected $_dboptions = array (
			PDO::MYSQL_ATTR_INIT_COMMAND => 'SET NAMES utf8',
			PDO::ATTR_EMULATE_PREPARES => false,
			PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
			PDO::ATTR_PERSISTENT => true 
	);
	protected $_myqueries = array (
			"SELECT ID, expectedKms, duration FROM udo_route WHERE ID=%s;" 
	);
	
	// : Public Functions
	// : Accessors
	// : End
	
	// : Magic
	/**
	 * MAXLive_Route_Missing_Distance::__construct()
	 * Class constructor
	 */
	public function __construct() {
		$ini = dirname ( realpath ( __FILE__ ) ) . self::DS . self::INI_DIR . self::DS . self::INI_FILE;
		if (is_file ( $ini ) === FALSE) {
			echo "No " . self::INI_FILE . " file found. Please create it and populate it with the following data: username=x@y.com, password=`your password`, your name shown on MAX the welcome page welcome=`Joe Soap` and mode=`test` or `live`" . PHP_EOL;
			return FALSE;
		}
		$data = parse_ini_file ( $ini );
		if ((array_key_exists ( "screenshotdir", $data ) && $data ["screenshotdir"]) && (array_key_exists ( "errordir", $data ) && $data ["errordir"]) && (array_key_exists ( "xls", $data ) && $data ["xls"]) && (array_key_exists ( "wdport", $data ) && $data ["wdport"]) && (array_key_exists ( "datadir", $data ) && $data ["datadir"]) && (array_key_exists ( "ip", $data ) && $data ["ip"]) && (array_key_exists ( "username", $data ) && $data ["username"]) && (array_key_exists ( "password", $data ) && $data ["password"]) && (array_key_exists ( "welcome", $data ) && $data ["welcome"]) && (array_key_exists ( "mode", $data ) && $data ["mode"])) {
			$this->_username = $data ["username"];
			$this->_password = $data ["password"];
			$this->_welcome = $data ["welcome"];
			$this->_dataDir = $data ["datadir"];
			$this->_errDir = $data ["errordir"];
			$this->_scrDir = $data ["screenshotdir"];
			$this->_mode = $data ["mode"];
			$this->_ip = $data ["ip"];
			$this->_wdport = $data ["wdport"];
			$this->_browser = $data ["browser"];
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
	 * MAXLive_Route_Missing_Distance::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End
	public function setUp() {
		$wd_host = "http://localhost:$this->_wdport/wd/hub";
		self::$driver = new PHPWebDriver_WebDriver ( $wd_host );
		$this->_session = self::$driver->session ( $this->_browser );
	}
	
	// : Setters
	public function SetcURLTimeout($_Timeout) {
		if (is_int ( $_Timeout )) {
			$this->_cURLTimeout = $_Timeout;
			return TRUE;
		} else {
			return FALSE;
		}
	}
	// : End
	
	// : Getters
	public function GetcURLTimeout() {
		return $this->_cURLTimeout;
	}
	
	// : End
	
	/**
	 * MAXLive_Route_Missing_Distance::testFunctionTemplate
	 * This is a function description for a selenium test function
	 */
	public function testFunctionTemplate() {
		
		// Initiate Session
		$session = $this->_session;
		$this->_session->setPageLoadTimeout ( 120 );
		$w = new PHPWebDriver_WebDriverWait ( $this->_session );
		
		// Construct an array with the customer names to use with script
		$rate_id = ( string ) "";
		
		$_xlsfile = dirname ( realpath ( __FILE__ ) ) . self::DS . $this->_dataDir . self::DS . $this->_xls;
		
		if (file_exists ( $_xlsfile )) {
			
			try {
				// : Load XLS data
				$_xlsData = new ReadExcelFile ( $_xlsfile, "Sheet1" );
				$_data = $_xlsData->getData ();
				
				// Connect to database
				$_mysqlDsn = preg_replace ( "/%s/", $this->_ip, $this->_dbdsn );
				$this->openDB ( $_mysqlDsn, $this->_dbuser, $this->_dbpwd, $this->_dboptions );
				
				// : Login
				try {
					// Load MAX home page
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
					
					// : Load Planningboard to rid of iframe loading on every page from here on
					$this->_session->open ( $this->_maxurl . self::PB_URL );
					$e = $w->until ( function ($session) {
						return $session->element ( "xpath", "//*[contains(text(),'You Are Here') and contains(text(), 'Planningboard')]" );
					} );
					// : End
				} catch ( Exception $e ) {
					throw new Exception ( "Error: Failed to log into MAX." . PHP_EOL . $e->getMessage () );
				}
				// : End
				
				// : Main Loop
				$_duration = ( int ) 0;
				
				$_total = count ( $_data ["ID"] );
				if ($_total != 0) {
					for($x = 1; $x <= $_total; $x ++) {
						
						try {
							
							$_sqlquery = preg_replace ( "/%s/", $_data ["ID"] [$x], $this->_myqueries[0] );
							$_result = $this->queryDB ( $_sqlquery );
							if (count ( $_result ) != 0) {
								$_distance = $_result [0] ["expectedKms"];
								if ($_distance == "0.00") {
									$_from = "";
									$_to = "";
									$_distance = "";
									$_duration = 0;
									$_unit = "km";
									if ($_data ["LocationFrom"] [$x]) {
										$_from = $_data ["LocationFrom"] [$x];
									}
									if ($_data ["ParentFrom"] [$x]) {
										$_from .= " " . $_data ["ParentFrom"] [$x];
									}
									if ($_data ["LocationTo"] [$x]) {
										$_to = $_data ["LocationTo"] [$x];
									}
									if ($_data ["ParentTo"] [$x]) {
										$_to .= " " . $_data ["ParentTo"] [$x];
									}
									$_test = $this->getGoogleMapsDirectionsAPIData ( $_from, $_to, "false", "driving" );
									if (count ( $_test ["routes"] [0] ["legs"] [0] ["distance"] ) > 0) {
										$_distance = $_test ["routes"] [0] ["legs"] [0] ["distance"] ["text"];
										preg_match("/\s(m|km)$/", $_distance, $_matches);
										if (($_matches) && (count($_matches) > 0)) {
											$_unit = $_matches[1];
										}
										$_distance = preg_replace ( "/\s.?m$/", "", $_distance );
										$_distance = preg_replace ("/\,/", "", $_distance);
										if ($_unit == "m") {
											$_distance = strval(number_format((floatval($_distance) / 1000), 3, ".", "")); 
										}
										$_duration = number_format ( ((floatval ( $_distance ) / 80) * 60), 0, "", "" );
									} else {
										throw new Exception ( "ERROR: Distance not found using Google API for record: " . $_data ["ID"] [$x] );
									}
									
									$this->_session->open ( $this->_maxurl . self::ROUTE_URL . $_data ["ID"] [$x] );
									$e = $w->until ( function ($session) {
										return $session->element ( "css selector", ".toolbar-cell-update" );
									} );
									$this->_session->element ( "css selector", ".toolbar-cell-update" )->click ();
									$e = $w->until ( function ($session) {
										return $session->element ( "xpath", "//*[@id='udo_Route-4_0_0_expectedKms-4']" );
									} );
									// : Assert all elements are on the page
									$this->assertElementPresent ( "xpath", "//*[@id='udo_Route-6__0_locationFrom_id-6']" );
									$this->assertElementPresent ( "xpath", "//*[@id='udo_Route-7__0_locationTo_id-7']" );
									$this->assertElementPresent ( "xpath", "//*[@id='udo_Route-3_0_0_duration-3']" );
									$this->assertElementPresent ( "css selector", "input[name=save][type=submit]" );
									// : End
									
									if (($_distance != "0.00") && ($_distance != "")) {
										$this->_session->element ( "xpath", "//*[@id='udo_Route-4_0_0_expectedKms-4']" )->clear ();
										$this->_session->element ( "xpath", "//*[@id='udo_Route-3_0_0_duration-3']" )->clear ();
										
										$this->_session->element ( "xpath", "//*[@id='udo_Route-4_0_0_expectedKms-4']" )->sendKeys ( $_distance );
										$this->_session->element ( "xpath", "//*[@id='udo_Route-3_0_0_duration-3']" )->sendKeys ( strval ( $_duration ) );
									}
									$this->_session->element ( "css selector", "input[name=save][type=submit]" )->click ();
									
									$this->_processed [$x] ["Record"] = $_data ["ID"] [$x];
									$this->_processed [$x] ["Distance"] = $_distance;
									$this->_processed [$x] ["Duration"] = strval ( $_duration );
								} else {
									throw new Exception ("ERROR: Route already has distance saved.");
								}
							}
							else {
								throw new Exception ("ERROR: Route with ID not found");
							}
						} catch ( Exception $e ) {
							$_errCount = count ( $this->_errors ) + 1;
							$this->_errors [$_errCount] ["Message"] = $e->getMessage ();
							$this->_errors [$_errCount] ["Record"] = $_data ["ID"] [$x];
						}
					}
				}
				
				// : End
			} catch ( Exception $e ) {
				throw new Exception ( "Something went wrong when attempting to log into MAX, see error message below." . PHP_EOL . $e->getMessage () );
			}
			// : End
		}
		if (count ( $this->_errors ) != 0) {
			print_r ( $this->_errors );
		}
		
		print (PHP_EOL) ;
		print_r ( $this->_processed );
		
		// : Tear Down
		$this->_session->element ( 'xpath', "//*[contains(@href,'/logout')]" )->click ();
		// Wait for page to load and for elements to be present on page
		$e = $w->until ( function ($session) {
			return $session->element ( 'css selector', 'input[id=identification]' );
		} );
		$this->assertElementPresent ( 'css selector', 'input[id=identification]' );
		$this->_session->close ();
		// : End
	}
	
	// : Private Functions
	
	/**
	 * MAXLive_Route_Missing_Distance::writeExcelFile($excelFile, $excelData)
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
	 * MAXLive_Route_Missing_Distance::takeScreenshot()
	 * This is a function description for a selenium test function
	 *
	 * @param object: $_session        	
	 */
	private function takeScreenshot() {
		$_img = $this->_session->screenshot ();
		$_data = base64_decode ( $_img );
		$_file = dirname ( __FILE__ ) . DIRECTORY_SEPARATOR . "Screenshots" . DIRECTORY_SEPARATOR . date ( "Y-m-d_His" ) . "_WebDriver.png";
		$_success = file_put_contents ( $_file, $_data );
		if ($_success) {
			return $_file;
		} else {
			return FALSE;
		}
	}
	
	/**
	 * MAXLive_Route_Missing_Distance::assertElementPresent($_using, $_value)
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
	 * MAXLive_Route_Missing_Distance::openDB($dsn, $username, $password, $options)
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
	 * MAXLive_Route_Missing_Distance::closeDB()
	 * Close connection to Database
	 */
	private function closeDB() {
		$this->_db = null;
	}
	
	/**
	 * MAXLive_Route_Missing_Distance::queryDB($sqlquery)
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
	private function getGoogleMapsDirectionsAPIData($origin, $destination, $alternatives, $mode) {
		$origin = urlencode ( $origin );
		$destination = urlencode ( $destination );
		$alternatives = urlencode ( $alternatives );
		$mode = urlencode ( $mode );
		$url = "http://maps.googleapis.com/maps/api/directions/json?origin=$origin&destination=$destination&alternatives=$alternatives&mode=$mode";

		// create curl resource
		$ch = curl_init ();
		// set url
		curl_setopt ( $ch, CURLOPT_URL, $url );
		// return the transfer as a string
		curl_setopt ( $ch, CURLOPT_RETURNTRANSFER, 1 );
		
		// $output contains the output string
		$output = curl_exec ( $ch );
		
		// close curl resource to free up system resources
		curl_close ( $ch );
		
		$_result = json_decode ( $output, TRUE );
		if (count ( $_result ) != 0) {
			return $_result;
		} else {
			return FALSE;
		}
	}
	
	// : End
}
?>