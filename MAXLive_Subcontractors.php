<?php
// : Includes
require_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriver.php');
require_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverWait.php');
require_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverBy.php');
require_once dirname ( __FILE__ ) . '/ReadExcelFile.php';
require_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';
/**
 * PHPExcel_Writer_Excel2007
 */
include 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel/Writer/Excel2007.php';
// : End

/**
 * Object::MAXLive_Subcontractors
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Manline Group (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class MAXLive_Subcontractors extends PHPUnit_Framework_TestCase {
	// : Constants
	const DS = DIRECTORY_SEPARATOR;
	const PB_URL = "/Planningboard";
	const COULD_NOT_CONNECT_MYSQL = "Failed to connect to MySQL database";
	const MAX_NOT_RESPONDING = "Error: MAX does not seem to be responding";
	const SUBBIE_URL = "/DataBrowser?browsePrimaryObject=997&browsePrimaryInstance=";
	const CUSTOMER_URL = "/DataBrowser?browsePrimaryObject=461&browsePrimaryInstance=";
	const LOCATION_BU_URL = "/DataBrowser?browsePrimaryObject=495&browsePrimaryInstance=";
	const OFF_CUST_BU_URL = "/DataBrowser?browsePrimaryObject=494&browsePrimaryInstance=";
	const RATEVAL_URL = "/DataBrowser?browsePrimaryObject=udo_Rates&browsePrimaryInstance=%s&browseSecondaryObject=DateRangeValue&relationshipType=Rate";
	const BF = "0.00";
	const CONTRIB = "Freight (Long Distance)";
	const LIVE_URL = "https://login.max.bwtsgroup.com";
	const TEST_URL = "http://max.mobilize.biz";
	const INI_FILE = "subbies_data.ini";
	const INI_DIR = "ini";
	const TEST_SESSION = "firefox";
	const XLS_CREATOR = "MAXLive_Subcontractors.php";
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
	protected $_error = array ();
	protected $_files = array ();
	protected $_dataDir;
	protected $_maxurl;
	protected $_mode;
	protected $_ip;
	protected $_username;
	protected $_password;
	protected $_welcome;
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
			"select ID from udo_rates where route_id IN (select ID from udo_route where locationFrom_id IN (select ID from udo_location where name='%f') and locationTo_id IN (select ID from udo_location where name='%t')) and objectregistry_id=%g and objectInstanceId=%c and truckDescription_id=%d and enabled=1 and model IS NULL and businessUnit_id=%b and rateType_id=%r;",
			"select ID from objectregistry where handle = 'udo_Customer';",
			"select ID from objectregistry where handle = 'udo_subcontractor';",
			"select ID from udo_route where locationFrom_id IN (select ID from udo_location where name='%f') and locationTo_id IN (select ID from udo_location where name='%t');",
			"select ID from udo_subcontractor where name='%s';",
			"select ID, _type, parent_id, name from udo_location where name='%s';",
			"select ID, _type, parent_id, name from udo_location where ID=%s;",
			"select ID from %t where %f='%v';",
			"select ID from %t where %f like '%v';" 
	);
	
	// : Public Functions
	// : Accessors
	// : End
	
	// : Magic
	/**
	 * MAXLive_Subcontractors::__construct()
	 * Class constructor
	 */
	public function __construct() {
		$ini = dirname ( realpath ( __FILE__ ) ) . self::DS . self::INI_DIR . self::DS . self::INI_FILE;
		echo $ini;
		if (is_file ( $ini ) === FALSE) {
			echo "No " . self::INI_FILE . " file found. Please create it and populate it with the following data: username=x@y.com, password=`your password`, your name shown on MAX the welcome page welcome=`Joe Soap` and mode=`test` or `live`" . PHP_EOL;
			return FALSE;
		}
		$data = parse_ini_file ( $ini );
		if ((array_key_exists ( "dataDir", $data ) && $data ["dataDir"]) && (array_key_exists ( "ip", $data ) && $data ["ip"]) && (array_key_exists ( "username", $data ) && $data ["username"]) && (array_key_exists ( "password", $data ) && $data ["password"]) && (array_key_exists ( "welcome", $data ) && $data ["welcome"]) && (array_key_exists ( "mode", $data ) && $data ["mode"])) {
			$this->_username = $data ["username"];
			$this->_password = $data ["password"];
			$this->_welcome = $data ["welcome"];
			$this->_dataDir = $data ["dataDir"];
			$this->_mode = $data ["mode"];
			$this->_ip = $data ["ip"];
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
		// : Search for xls files in data dir and save into an array
		$_dir = dirname ( realpath ( __FILE__ ) ) ;
		$_dir = preg_replace("/\//", self::DS, $_dir);
		$_dir .= self::DS . $this->_dataDir;
		$_files = scandir ( $_dir );
		$_count = ( int ) 0;
		
		foreach ( $_files as $_file ) {
			preg_match ( "/^(.*)\.(.*)$/", $_file, $x );
			if (count ( $x ) != 0) {
				if ($x [2] == "xls") {
					$this->_files [$_count] ["customer"] = $x [1];
					$this->_files [$_count] ["filename"] = $x [1] . "." . $x [2];
					$_count ++;
				}
			}
		}
		// : End
	}
	
	/**
	 * MAXLive_Subcontractors::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End
	public function setUp() {
		self::$driver = new PHPWebDriver_WebDriver ();
		$this->_session = self::$driver->session ( self::TEST_SESSION );
	}
	
	/**
	 * MAXLive_Subcontractors::testFunctionTemplate
	 * This is a function description for a selenium test function
	 */
	public function testFunctionTemplate() {
		// Initiate Session
		$session = $this->_session;
		$this->_session->setPageLoadTimeout ( 60 );
		$w = new PHPWebDriver_WebDriverWait ( $this->_session );
		
		// Construct an array with the customer names to use with script
		$rate_id = ( string ) "";
		
		// Connect to database
		$_mysqlDsn = preg_replace ( "/%s/", $this->_ip, $this->_dbdsn );
		$this->openDB ( $_mysqlDsn, $this->_dbuser, $this->_dbpwd, $this->_dboptions );
		
		// : Query and save objectregistry_id for udo_subcontractor
		$myQuery = $this->_myqueries [5];
		$result = $this->queryDB ( $myQuery );
		$objectregistry_id = $result [0] ["ID"];
		// : End
		
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
			throw new Exception ("Something went wrong when attempting to log into MAX, see error message below." . PHP_EOL . $e->getMessage());
		}
		// : End
		
		// : Load Planningboard to rid of iframe loading on every page from here on
		$this->_session->open ( $this->_maxurl . self::PB_URL );
		$e = $w->until ( function ($session) {
			return $session->element ( "xpath", "//*[contains(text(),'You Are Here') and contains(text(), 'Planningboard')]" );
		} );
		// : End
		
		foreach ( $this->_files as $_xlsfile ) {
			$_customer = $_xlsfile ["customer"];
			// Prepare file path variable
			$_file = dirname ( __FILE__ ) . $this->_dataDir . DIRECTORY_SEPARATOR . $_xlsfile ["filename"];
			$_file = preg_replace ( "$\/$", DIRECTORY_SEPARATOR, $_file );

			if (file_exists ( $_file )) {
				
				// : Setup variables and data
				$this->_error = array ();
				echo $_file . PHP_EOL;
				$_xlsData = new ReadExcelFile ( $_file, "Sheet1" );
				$_data = $_xlsData->getData ();
				
				// : Extract columns from the spreadsheet data
				$_xlsColumns = array ();
				$_xlsColumns [] = "Error_Msg";
				foreach ( $_data as $key => $value ) {
					$_xlsColumns [] = $key;
				}
				
				// : Query and save subcontractor id
				$myQuery = preg_replace ( "/%s/", $_customer, $this->_myqueries [7] );
				$result = $this->queryDB ( $myQuery );
				$subbie_id = $result [0] ["ID"];
				// : End
				
				$_locations = ( array ) array (
						"LocationFrom" => "",
						"LocationTo" => "" 
				);
				$_count = count ( $_data ["LocationFrom"] );
				for($x = 1; $x <= $_count; $x ++) {
					
					try {
						$this->lastRecord = "Line: " . $x . ", From: " . $_data ["LocationFrom"] [$x] . ", To: " . $_data ["LocationTo"] [$x];
						$rate_id = NULL;
						$_vars = array (
								"TruckDescription" => array (
										"table" => "udo_truckdescription",
										"field" => "description",
										"value" => "" 
								),
								"BusinessUnit" => array (
										"table" => "udo_businessunit",
										"field" => "name",
										"value" => "" 
								),
								"RateType" => array (
										"table" => "udo_ratetype",
										"field" => "name",
										"value" => "" 
								) 
						);
						
						// : Get truck description and business unit ID from database
						foreach ( $_vars as $_key => $_var ) {
							$myQuery = preg_replace ( "/%v/", $_data [$_key] [$x], $this->_myqueries [10] );
							$myQuery = preg_replace ( "/%t/", $_var ["table"], $myQuery );
							$myQuery = preg_replace ( "/%f/", $_var ["field"], $myQuery );
							$_resultA = $this->queryDB ( $myQuery );
							if (count ( $_resultA ) != 0) {
								$_vars [$_key] ["value"] = $_resultA [0] ["ID"];
							} else {
								$myQuery = preg_replace ( "/%v/", $_data [$_key] [$x], $this->_myqueries [11] );
								$myQuery = preg_replace ( "/%t/", $_var ["table"], $myQuery );
								$myQuery = preg_replace ( "/%f/", $_var ["field"], $myQuery );
								$_resultB = $this->queryDB ( $myQuery );
								if (count ( $_resultB ) != 0) {
									$_vars [$_key] ["value"] = $_resultB [0] ["ID"];
								} else {
									throw new Exception ( "Cannot find " . $_key . " for route and rate. \nLine: " . ($x) . ", Customer: " . $_customer . "\nPlease check truck description and amend in spreadsheet." );
								}
							}
						}
						// : End
						
						// : Look for bad character conversion when using hypen in xls documents and convert
						$_data ["LocationFrom"] [$x] = preg_replace ( "/–/", "-", $_data ["LocationFrom"] [$x] );
						$_data ["LocationTo"] [$x] = preg_replace ( "/–/", "-", $_data ["LocationTo"] [$x] );
						// : End
						
						// : Loop query for parent location of location from and to names until we get the province location record for each
						foreach ( $_locations as $key => $_aLocation ) {
							$_type = "";
							$_parentid = "";
							$a = 1;
							$_location = $_data [$key] [$x];
							$aQuery = preg_replace ( "/%s/", $_location, $this->_myqueries [8] );
							$_result = $this->queryDB ( $aQuery );
							if (count ( $_result ) != 0) {
								if ((array_key_exists ( "_type", $_result [0] )) && (array_key_exists ( "name", $_result [0] )) && (array_key_exists ( "parent_id", $_result [0] )) && (array_key_exists ( "ID", $_result [0] ))) {
									$_type = $_result [0] ["_type"];
									$_parentid = $_result [0] ["parent_id"];
									$_location = $_result [0] ["name"];
								}
							} else {
								break;
							}
							
							while ( ($_type != "udo_Province") || ($a !== 3) ) {
								$aQuery = preg_replace ( "/%s/", $_parentid, $this->_myqueries [9] );
								$_result = $this->queryDB ( $aQuery );
								if (count ( $_result ) != 0) {
									if ((array_key_exists ( "_type", $_result [0] )) && (array_key_exists ( "name", $_result [0] )) && (array_key_exists ( "parent_id", $_result [0] )) && (array_key_exists ( "ID", $_result [0] ))) {
										$_type = $_result [0] ["_type"];
										$_parentid = $_result [0] ["parent_id"];
										$_location = $_result [0] ["name"];
										if ($_type == "udo_Province") {
											$_locations [$key] = $_result [0] ["name"];
											break;
										}
									}
								}
								$a ++;
							}
						}
						// : End
						
						// : Query database to find if route and rate exists
						$myQuery = preg_replace ( "/%f/", $_data ["LocationFrom"] [$x], $this->_myqueries [3] );
						$myQuery = preg_replace ( "/%t/", $_data ["LocationTo"] [$x], $myQuery );
						$myQuery = preg_replace ( "/%g/", $objectregistry_id, $myQuery );
						$myQuery = preg_replace ( "/%c/", $subbie_id, $myQuery );
						$myQuery = preg_replace ( "/%d/", $_vars ["TruckDescription"] ["value"], $myQuery );
						$myQuery = preg_replace ( "/%b/", $_vars ["BusinessUnit"] ["value"], $myQuery );
						$myQuery = preg_replace ( "/%r/", $_vars ["RateType"] ["value"], $myQuery );
						$result = $this->queryDB ( $myQuery );
						$this->lastRecord .= " ( " . $myQuery . " )";
						if (count ( $result ) != 0) {
							$rate_id = $result [0] ["ID"];
						} else {
							$rate_id = NULL;
						}
						
						// If route and rate does not already exist then create route and rate
						if ($rate_id == NULL) {
							
							// : Load Subbie page and Create Route
							// : #1 - Load page and wait an assert element to make sure has loaded
							$this->_session->open ( $this->_maxurl . self::SUBBIE_URL . $subbie_id );
							
							$this->_dummy = $_customer;
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'" . $this->_dummy . "')]" );
							} );
							// : #1 - End
							// : #2 - Assert elements are present and load rates page for customer
							$this->assertElementPresent ( "xpath", "//*[contains(text(),'" . $_customer . "')]" );
							$this->assertElementPresent ( 'css selector', 'span#subtabselector' );
							// : #2 - End
							// : #3 - Select option from selectbox to load rates for customer
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[contains(text(),'Rates for this Customer')]" )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							// : #3 - End
							// : #4 - Create new rate
							$this->assertElementPresent ( 'css selector', '#button-create' );
							$this->_session->element ( 'css selector', '#button-create' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the details of Rates')]" );
							} );
							$this->assertElementPresent ( 'css selector', '#udo_Rates-45__0_provinceFrom_id-45' );
							$this->assertElementPresent ( 'css selector', '#udo_Rates-46__0_cityFrom_id-46' );
							$this->assertElementPresent ( 'css selector', '#udo_Rates-47__0_provinceTo_id-47' );
							$this->assertElementPresent ( 'css selector', '#udo_Rates-48__0_cityTo_id-48' );
							$this->assertElementPresent ( 'css selector', '#udo_Rates-30__0_rateType_id-30' );
							$this->assertElementPresent ( 'css selector', '#udo_Rates-4__0_businessUnit_id-4' );
							$this->assertElementPresent ( 'css selector', '#udo_Rates-36__0_truckDescription_id-36' );
							$this->assertElementPresent ( 'css selector', '#udo_Rates-18_0_0_leadKms-18' );
							$this->assertElementPresent ( 'css selector', '#checkbox_udo_Rates-15_0_0_enabled-15' );
							$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-45__0_provinceFrom_id-45']/option[contains(text(),'" . $_locations ["LocationFrom"] . "')]" )->click ();
							$this->_dummy = $_data ["LocationFrom"] [$x];
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[@id='udo_Rates-46__0_cityFrom_id-46']/option[text()='" . $this->_dummy . "']" );
							} );
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-46__0_cityFrom_id-46']/option[text()='" . $_data ["LocationFrom"] [$x] . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-47__0_provinceTo_id-47']/option[contains(text(),'" . $_locations ["LocationTo"] . "')]" )->click ();
							$this->_dummy = $_data ["LocationTo"] [$x];
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[@id='udo_Rates-48__0_cityTo_id-48']/option[text()='" . $this->_dummy . "']" );
							} );
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-48__0_cityTo_id-48']/option[text()='" . $_data ["LocationTo"] [$x] . "']" )->click ();
							$this->_dummy = $_data ["LocationFrom"] [$x] . " TO " . $_data ["LocationTo"] [$x];
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'" . $this->_dummy . "')]" );
							} );
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-30__0_rateType_id-30']/option[text()='" . $_data ["RateType"] [$x] . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-4__0_businessUnit_id-4']/option[text()='" . $_data ["BusinessUnit"] [$x] . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-36__0_truckDescription_id-36']/option[text()='" . $_data ["TruckDescription"] [$x] . "']" )->click ();
							;
							if ($_data ["LeadKms"] [$x] != "0") {
								$this->_session->element ( 'css selector', '#udo_Rates-18_0_0_leadKms-18' )->sendKeys ( $_data ["LeadKms"] [$x] );
							}
							$this->_session->element ( 'css selector', '#checkbox_udo_Rates-15_0_0_enabled-15' )->click ();
							$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							// : End
							
							// : Query database to find newly created route to create a daterangevalue rate for the route
							$myQuery = preg_replace ( "/%f/", $_data ["LocationFrom"] [$x], $this->_myqueries [3] );
							$myQuery = preg_replace ( "/%t/", $_data ["LocationTo"] [$x], $myQuery );
							$myQuery = preg_replace ( "/%g/", $objectregistry_id, $myQuery );
							$myQuery = preg_replace ( "/%c/", $subbie_id, $myQuery );
							$myQuery = preg_replace ( "/%d/", $_vars ["TruckDescription"] ["value"], $myQuery );
							$myQuery = preg_replace ( "/%b/", $_vars ["BusinessUnit"] ["value"], $myQuery );
							$myQuery = preg_replace ( "/%r/", $_vars ["RateType"] ["value"], $myQuery );
							$result = $this->queryDB ( $myQuery );
							$this->lastRecord .= " ( " . $myQuery . " )";
							if (count ( $result ) != 0) {
								$rate_id = $result [0] ["ID"];
							} else {
								$rate_id = NULL;
							}
							// : End
						}
						
						if ($rate_id != NULL) {
							$rateurl = preg_replace ( "/%s/", $rate_id, $this->_maxurl . self::RATEVAL_URL );
							$this->_session->open ( $rateurl );
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							$this->assertElementPresent ( 'css selector', '#button-create' );
							$this->_session->element ( 'css selector', '#button-create' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Date Range Values')]" );
							} );
							
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-20_0_0_value-20' );
							$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
							
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->sendKeys ( date ( "Y-m-01 00:00:00", strtotime ( "-1 month" ) ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->sendKeys ( strval ( $_data ["Rate"] [$x] ) );
							$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
						}
						// : End
					} catch ( Exception $e ) {
						echo "Error: " . $e->getMessage () . PHP_EOL;
						echo "Time of error: " . date ( "Y-m-d H:i:s" ) . PHP_EOL;
						echo "Last record: " . $this->lastRecord;
						$this->takeScreenshot ();
						$_erCount = count ( $this->_error );
						$this->_error [$_erCount + 1] ["error"] = $e->getMessage ();
						foreach ( $_data as $key => $value ) {
							$this->_error [$_erCount + 1] [$key] = $value [$x];
						}
					}
				}
				
				// : If errors occured. Create xls of entries that failed.
				if (count ( $this->_error ) != 0) {
					$_xlsfilename = (dirname ( __FILE__ ) . self::DS . "error_reports" . self::DS . date ( "Y-m-d_His_" ) . "MAXLiveSubbies_" . $_customer . ".xlsx");
					$this->writeExcelFile ( $_xlsfilename, $this->_error, $_xlsColumns );
					if (file_exists ( $_xlsfilename )) {
						print ("Excel error report written successfully to file: $_xlsfilename") ;
					} else {
						print ("Excel error report write unsuccessful") ;
					}
				}
				// : End
			}
		}
		
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
	 * MAXLive_Subcontractors::takeScreenshot()
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
	 * MAXLive_Subcontractors::assertElementPresent($_using, $_value)
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
	 * MAXLive_Subcontractors::openDB($dsn, $username, $password, $options)
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
	 * MAXLive_Subcontractors::closeDB()
	 * Close connection to Database
	 */
	private function closeDB() {
		$this->_db = null;
	}
	
	/**
	 * MAXLive_Subcontractors::queryDB($sqlquery)
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
?>