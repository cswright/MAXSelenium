<?php

// : Error reporting settings
error_reporting(E_ALL);
ini_set('log_errors','1');
ini_set('error_log', dirname(__FILE__) . '/my-errors.log');
ini_set('display_errors','0');
// : End

// : Includes
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriver.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverWait.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverBy.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverProxy.php');
include_once dirname ( __FILE__ ) . '/FandVReadXLSData.php';
include_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';
/**
 * PHPExcel_Writer_Excel2007
 */
include 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel/Writer/Excel2007.php';
// : End

/**
 * Object::MAXLive_CreateFandVContracts
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Manline Group (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class MAXLive_CreateFandVContracts extends PHPUnit_Framework_TestCase {
	// : Constants
	const PB_URL = "/Planningboard";
	const COULD_NOT_CONNECT_MYSQL = "Failed to connect to MySQL database";
	const MAX_NOT_RESPONDING = "Error: MAX does not seem to be responding";
	const CUSTOMER_URL = "/DataBrowser?browsePrimaryObject=461&browsePrimaryInstance=";
	const LOCATION_BU_URL = "/DataBrowser?browsePrimaryObject=495&browsePrimaryInstance=";
	const OFF_CUST_BU_URL = "/DataBrowser?browsePrimaryObject=494&browsePrimaryInstance=";
	const RATEVAL_URL = "/DataBrowser?browsePrimaryObject=udo_Rates&browsePrimaryInstance=%s&browseSecondaryObject=DateRangeValue&relationshipType=Rate";
	const CONTRIB = "Freight (Long Distance)";
	const LIVE_URL = "https://login.max.bwtsgroup.com";
	const TEST_URL = "http://max.mobilize.biz";
	const INI_FILE = "fandv_data.ini";
	const INI_DIR = "ini";
	const TEST_SESSION = "firefox";
	const CUSTOMERURL = "/DataBrowser?browsePrimaryObject=461&browsePrimaryInstance=%s&browseSecondaryObject=910&useDataViewForSecondary=758&tab_id=61";
	const FANDVURL = "/DataBrowser?browsePrimaryObject=910&browsePrimaryInstance=";
	const RATEDATAURL = "/DataBrowser?browsePrimaryObject=udo_Rates&browsePrimaryInstance=";
	const DS = DIRECTORY_SEPARATOR;
	const XLS_CREATOR = "MAXLive_CreateFandVContracts.php";
	const XLS_TITLE = "Error Report";
	const XLS_SUBJECT = "Errors caught while creating F & V contracts";
	
	// : Variables
	protected static $driver;
	protected $_dummy;
	protected $_session;
	protected $lastRecord;
	protected $to = 'clintonabco@gmail.com';
	protected $subject = 'MAX Selenium script report';
	protected $message;
	protected $_maxurl;
	protected $_mode;
	protected $_error = array ();
	protected $_wdport;
	protected $var_rate_id;
	protected $_username;
	protected $_password;
	protected $_welcome;
	protected $_xls;
	protected $_ip;
	protected $_proxyip;
	protected $_data;
	protected $_dataDir;
	protected $_errDir;
	protected $_scrDir;
	protected $_db;
	protected $_browser;
	protected $_dbdsn = "mysql:host=%s;dbname=max2;charset=utf8;";
	protected $_dbuser = "root";
	protected $_dbpwd = "kaluma";
	protected $_dboptions = array (
			PDO::MYSQL_ATTR_INIT_COMMAND => 'SET NAMES utf8',
			PDO::ATTR_EMULATE_PREPARES => false,
			PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
			PDO::ATTR_PERSISTENT => true 
	);
	
	// : Public functions
	// : Accessors
	
	// : End
	
	// : Magic
	/**
	 * MAXLive_CreateFandVContracts::__construct()
	 * Class constructor
	 */
	public function __construct() {
		$ini = dirname ( realpath ( __FILE__ ) ) . self::DS . self::INI_DIR . self::DS . self::INI_FILE;
		if (is_file ( $ini ) === FALSE) {
			echo "File $ini not found. Please create it and populate it with the following data: username=x@y.com, password=`your password`, your name shown on MAX the welcome page welcome=`Joe Soap` and mode=`test` or `live`" . PHP_EOL;
			return FALSE;
		}
		$data = parse_ini_file ( $ini );
		if ((array_key_exists ( "proxy", $data ) && $data ["proxy"]) && (array_key_exists ( "browser", $data ) && $data ["browser"]) && (array_key_exists ( "ip", $data ) && $data ["ip"]) && (array_key_exists ( "datadir", $data ) && $data ["datadir"]) && (array_key_exists ( "username", $data ) && $data ["username"]) && (array_key_exists ( "xls", $data ) && $data ["xls"]) && (array_key_exists ( "password", $data ) && $data ["password"]) && (array_key_exists ( "welcome", $data ) && $data ["welcome"]) && (array_key_exists ( "mode", $data ) && $data ["mode"])) {
			$this->_username = $data ["username"];
			$this->_password = $data ["password"];
			$this->_welcome = $data ["welcome"];
			$this->_dataDir = $data ["datadir"];
			$this->_errDir = $data ["errordir"];
			$this->_scrDir = $data ["screenshotdir"];
			$this->_xls = $data ["xls"];
			$this->_ip = $data ["ip"];
			$this->_proxyip = $data ["proxy"];
			$this->_wdport = $data ["wdport"];
			$this->_browser = $data ["browser"];
			$this->_mode = $data ["mode"];
			switch ($this->_mode) {
				case "live" :
					$this->_maxurl = self::LIVE_URL;
					break;
				default :
					$this->_maxurl = self::TEST_URL;
			}
		} else {
			echo "The correct data is not present in $ini. Please confirm the following fields are present in the file: username, password, welcome, mode, dataDir, ip, proxy, browser and xls." . PHP_EOL;
			return FALSE;
		}
	}
	
	/**
	 * MAXLive_CreateFandVContracts::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End
	
	/**
	 * MAXLive_CreateFandVContracts::setUp()
	 * Setup instance
	 */
	public function setUp() {
		$wd_host = "http://localhost:$this->_wdport/wd/hub";
		self::$driver = new PHPWebDriver_WebDriver ( $wd_host );
		$desired_capabilities = array();
		$proxy = new PHPWebDriver_WebDriverProxy();
		$proxy->httpProxy = $this->_proxyip;
		$proxy->add_to_capabilities($desired_capabilities);
		$this->_session = self::$driver->session ( $this->_browser, $desired_capabilities );
	}
	
	/**
	 * MAXLive_CreateFandVContracts::testCreateContracts()
	 * Pull F and V Contract data and automate creation of F and V Contracts
	 */
	public function testCreateContracts() {
		
		// : Pull F and V Contract data from correctly formatted xls spreadsheet
		$_datadir = preg_replace ( '/\//', self::DS, $this->_dataDir );
		$file = dirname ( __FILE__ ) . $_datadir . $this->_xls;
		if (file_exists ( $file )) {
			
			// Initiate Session
			$session = $this->_session;
			$this->_session->setPageLoadTimeout ( 90 );
			$w = new PHPWebDriver_WebDriverWait ( $this->_session );
			
			// : Get xls data
			$FandVContract = new FandVReadXLSData ( $file );
			$this->_data = $FandVContract->getData ();
			// : End
			
			// : Build columns to be used when creating the error report spreadsheet
			$_xlsColumns = array (
					"Error_Msg"
			);
			
			foreach ( $this->_data[0] as $key => $value ) {
				$_xlsColumns[] = $key; 
			}
			// : End
			
			// : Connect to database
			$_mysqlDsn = preg_replace ( "/%s/", $this->_ip, $this->_dbdsn );
			$this->openDB ( $_mysqlDsn, $this->_dbuser, $this->_dbpwd, $this->_dboptions );
			// : End
			
			// : Login
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
			// : End
			
			// : Load Planningboard to rid of iframe loading on every page from here on
			$this->_session->open ( $this->_maxurl . self::PB_URL );
			$e = $w->until ( function ($session) {
				return $session->element ( "xpath", "//*[contains(text(),'You Are Here') and contains(text(), 'Planningboard')]" );
			} );
			// : End
			
			// : Main Loop
			foreach ( $this->_data as $key => $value ) {
				try {
					// : Check customer and truck type exist in database
					$customer_id = NULL;
					$trucktype_id = NULL;
					// Store each record currently been processed for error reporting purposes
					$this->lastRecord = 'Contract: ' . $value ['Contract'] . ', Customer: ' . $value ['Customer'] . ', VRate: ' . strval ( $value ['Rate'] ) . ', TruckType: ' . $value ['Truck Type'] . ', Fixed Cost: ' . $value ['Cost'];
					// Get customer ID from database
					$customer_id = $this->queryDB ( "select ID from udo_customer where tradingName = '" . $value ["Customer"] . "';" );
					// Get truckdescription ID from database
					$trucktype_id = $this->queryDB ( "select ID from udo_truckdescription where description='" . $value ["Truck Type"] . "';" );
					if ((($customer_id != NULL) && (count ( $customer_id ) != 0)) && (($trucktype_id != NULL) && (count ( $trucktype_id ) != 0))) {
						$customer_id = $customer_id [0] ["ID"];
						$trucktype_id = $trucktype_id [0] ["ID"];
						$url = preg_replace ( '/%s/', $customer_id, self::CUSTOMERURL );
						$this->_session->open ( $this->_maxurl . $url );
					} else {
						throw new Exception ( "Customer not found: " . $value ["Customer"] );
					}
					// : End
					// Wait for element = #subtabselector
					$e = $w->until ( function ($session) {
						return $session->element ( 'css selector', '#subtabselector' );
					} );
					$this->assertElementPresent ( 'css selector', '#subtabselector' );
					$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='F and V Contracts']" )->click ();
					$e = $w->until ( function ($session) {
						return $session->element ( 'css selector', '#button-create' );
					} );
					$this->assertElementPresent ( 'css selector', '#button-create' );
					$this->_session->element ( 'css selector', '#button-create' )->click ();
					$e = $w->until ( function ($session) {
						return $session->element ( 'css selector', '#udo_FandVContract-2__0_businessUnit_id-2' );
					} );

					$this->assertElementPresent ( "xpath", "//*[contains(text(),'Create F and V Contracts')]" );
					$this->assertElementPresent ( "xpath", "//*[contains(text(),'" . $value ["Customer"] . "')]" );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-2__0_businessUnit_id-2' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-16_0_0_startDate-16' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-7_0_0_endDate-7' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-8_0_0_fixedContribution-8' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-9_0_0_fixedCost-9' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-28__0_truckDescription_id-28' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-30_0_0_variableCost-30' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-29__0_rateType_id-29' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-4_0_0_creditPercentage-4' );
					$this->assertElementPresent ( 'css selector', '#udo_FandVContract-13_0_0_numberOfDays-13' );
					$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
					
					$this->_session->element ( "xpath", "//*[@id='udo_FandVContract-2__0_businessUnit_id-2']/option[text()='" . $value ["Business Unit"] . "']" )->click ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-16_0_0_startDate-16' )->clear ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-16_0_0_startDate-16' )->sendKeys ( $value ['Start Date'] );
					$this->_session->element ( 'css selector', '#udo_FandVContract-7_0_0_endDate-7' )->clear ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-7_0_0_endDate-7' )->sendKeys ( $value ['End Date'] );
					$this->_session->element ( 'css selector', '#udo_FandVContract-8_0_0_fixedContribution-8' )->clear ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-8_0_0_fixedContribution-8' )->sendKeys ( strval ( number_format ( floatval ( $value ['Contrib'] ), 2, '.', '' ) ) );
					$this->_session->element ( 'css selector', '#udo_FandVContract-9_0_0_fixedCost-9' )->clear ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-9_0_0_fixedCost-9' )->sendKeys ( strval ( number_format ( floatval ( $value ['Cost'] ), 2, '.', '' ) ) );
					$this->_session->element ( "xpath", "//*[@id='udo_FandVContract-28__0_truckDescription_id-28']/option[text()='" . $value ["Truck Type"] . "']" )->click ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-30_0_0_variableCost-30' )->clear ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-30_0_0_variableCost-30' )->sendKeys ( strval ( $value ['Rate'] ) );
					$this->_session->element ( "xpath", "//*[@id='udo_FandVContract-29__0_rateType_id-29']/option[text()='" . $value ["RateType"] . "']" )->click ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-4_0_0_creditPercentage-4' )->clear ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-4_0_0_creditPercentage-4' )->sendKeys ( "0.00" );
					$this->_session->element ( 'css selector', '#udo_FandVContract-13_0_0_numberOfDays-13' )->clear ();
					$this->_session->element ( 'css selector', '#udo_FandVContract-13_0_0_numberOfDays-13' )->sendKeys ( strval ( $value ["Days"] ) );
					$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();

					$startDate = $value ["Start Date"];
					$endDate = $value ["End Date"];
					$startDate = date ( "Y-m-d H:i:s", strtotime ( "-2 hours", strtotime ( $startDate ) ) );
					$endDate = date ( "Y-m-d H:i:s", strtotime ( "-2 hours", strtotime ( $endDate ) ) );
					$myquery = array (
							"SELECT ID, variableCostRate_id FROM udo_fandvcontract WHERE fixedCost='" . strval ( number_format ( floatval ( $value ["Cost"] ), 2, "", "" ) ) . "' AND fixedContribution='" . strval ( number_format ( floatval ( $value ["Contrib"] ), 2, "", "" ) ) . "' AND startDate='" . $startDate . "' AND endDate='" . $endDate . "' AND businessUnit_id IN (SELECT ID FROM udo_businessunit WHERE name='" . $value ["Business Unit"] . "') AND customer_id IN (SELECT ID FROM udo_customer WHERE tradingName='" . $value ["Customer"] . "');",
							"SELECT description FROM udo_truckdescription WHERE ID IN (SELECT truckDescription_id FROM udo_rates WHERE ID=%s);",
							"SELECT value FROM daterangevalue WHERE objectInstanceid=%s and type='Rate';",
							"SELECT ID FROM udo_truck WHERE fleetnum='%t';" 
					);
					$myresult = $this->queryDB ( $myquery [0] );
					$fandvid = 0;
					if (count ( $myresult ) != 0) {
						foreach ( $myresult as $rows ) {
							$this->var_rate_id = $rows ["variableCostRate_id"];
							$aQuery = preg_replace ( '/%s/', $rows ["variableCostRate_id"], $myquery [1] );
							$truckDesc = $this->queryDB ( $aQuery );
							$aQuery = preg_replace ( '/%s/', $rows ["variableCostRate_id"], $myquery [2] );
							$varRate = $this->queryDB ( $aQuery );
							if (($value ["Rate"] != "0") && ($value ["Rate"] != "0.00")) {
								$_vrate = strval ( number_format ( floatval ( $value ["Rate"] ), 2, "", "" ) ) . ".0000";
							} else {
								$_vrate = "0.0000";
							}
							if (($truckDesc [0] ["description"] === $value ["Truck Type"]) && ($varRate [0] ["value"] === $_vrate)) {
								$fandvid = intval ( $rows ["ID"] );
							}
						}
					}
					
					if ($fandvid != 0) {
						if ($value ["Trucks Linked"] != "0") {
							$a = explode ( "\n", $value ["Trucks Linked"] );
							foreach ( $a as $fleetnum ) {
								if ($fleetnum != "") {
									$truckQuery = preg_replace ( '/%t/', $fleetnum, $myquery [3] );
									$result = $this->queryDB ( $truckQuery );
									if (count ( $result ) != 0) {
										foreach ( $result as $x ) {
											try {
												// : Load F and V page and goto truck links
												$this->_session->open ( $this->_maxurl . self::FANDVURL . strval ( $fandvid ) );
												// Wait for element = #subtabselector
												$e = $w->until ( function ($session) {
													return $session->element ( 'css selector', '#subtabselector' );
												} );
												// Select Truck Links from the selectbox
												$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='F and V Contracts - Truck']" )->click ();
												// : End
												// Wait for element = #button-create
												$e = $w->until ( function ($session) {
													return $session->element ( 'css selector', '#button-create' );
												} );
												// Click element = #button-create
												$this->_session->element ( 'css selector', '#button-create' )->click ();
												// Wait for element = truck_id field
												$e = $w->until ( function ($session) {
													return $session->element ( 'css selector', '#udo_FandVContractTruck_link-7__0_truck_id-7' );
												} );
												// : Assert elements are present on the page
												$this->assertElementPresent ( 'css selector', '#udo_FandVContractTruck_link-16_0_0_beginDate-16' );
												$this->assertElementPresent ( 'css selector', '#udo_FandVContractTruck_link-17_0_0_endDate-17' );
												$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
												// : End
												// Select truck from selectbox
												$this->_session->element ( "xpath", "//*[@id='udo_FandVContractTruck_link-7__0_truck_id-7']/option[text()='" . $fleetnum . "']" )->click ();
												// : Clear begin and end date fields
												$this->_session->element ( 'css selector', '#udo_FandVContractTruck_link-16_0_0_beginDate-16' )->clear ();
												$this->_session->element ( 'css selector', '#udo_FandVContractTruck_link-17_0_0_endDate-17' )->clear ();
												// : End
												// : Insert start and end date values
												$this->_session->element ( 'css selector', '#udo_FandVContractTruck_link-16_0_0_beginDate-16' )->sendKeys ( $value ["Start Date"] );
												$this->_session->element ( 'css selector', '#udo_FandVContractTruck_link-17_0_0_endDate-17' )->sendKeys ( $value ["End Date"] );
												// : End
												// Click element = submit button
												$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
											} catch ( Exception $e ) {
												$this->saveError ( $e, "ERROR: Failed to create truck link: " . $fleetnum, $key );
											}
										}
									}
								}
							}
						}
					}
					if ($value ["Routes Linked"] != "0") {
						$a = explode ( "\n", $value ["Routes Linked"] );

						foreach ( $a as $route ) {
							if ($route != "") {
								try {
									
									// : Load F and V page and goto truck links
									$this->_session->open ( $this->_maxurl . self::FANDVURL . strval ( $fandvid ) );
									// Wait for element = #subtabselector
									$e = $w->until ( function ($session) {
										return $session->element ( 'css selector', '#subtabselector' );
									} );
									$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='Routes in the F&V Contract']" )->click ();
									$e = $w->until ( function ($session) {
										return $session->element ( 'css selector', '#button-create' );
									} );
									// : End
									
									// : Use preg_match to build needed values
									preg_match ( '/^(.*TO.*)\[/', $route, $result );
									$route_name = $result [1];
									
									preg_match ( '/TO\s(.*)\[.*$/', $route, $result );
									$fromLocation = $result [1];
									
									preg_match ( '/(.*)\sTO/', $route, $result );
									$toLocation = $result [1];
									
									preg_match ( '/TO.*\[(.*)\]$/', $route, $result );
									$leadKms = $result [1];
									// : End

									$this->_session->element ( 'css selector', '#button-create' )->click ();
									$e = $w->until ( function ($session) {
										return $session->element ( "xpath", "//*[contains(text(),'Create F and V Contracts - Route')]" );
									} );
									// : Assert all elements are on the page
									$this->assertElementPresent ( 'css selector', '#udo_FandVContractRoute_link-6__0_route_id-6' );
									$this->assertElementPresent ( 'css selector', '#udo_FandVContractRoute_link-4_0_0_leadKms-4' );
									$this->assertElementPresent ( 'css selector', 'input[type=submit][name=save]' );
									// : End
									// Select the route from the selectbox
									$this->_session->element ( "xpath", "//*[@id='udo_FandVContractRoute_link-6__0_route_id-6']/option[text()='" . $route_name . "']" )->click ();
									// Clear the leadKms field
									$this->_session->element ( 'css selector', '#udo_FandVContractRoute_link-4_0_0_leadKms-4' )->clear ();
									if ($leadKms != "0") {
										// IF leadKms is not 0 then insert the leadkms value into the leadkms field
										$this->_session->element ( 'css selector', '#udo_FandVContractRoute_link-4_0_0_leadKms-4' )->sendKeys ( strval ( $leadKms ) );
									}
									// Click element = submit button
									$this->_session->element ( 'css selector', 'input[type=submit][name=save]' )->click ();
									
								} catch ( Exception $e ) {
									$this->saveError ( $e, "ERROR: Failed to create route link: " . $route_name, $key );
								}
							}
						}
					}
					if ($this->var_rate_id != NULL) {
						try {
							// Load rate data page for contract
							$this->_session->open ( $this->_maxurl . self::RATEDATAURL . $this->var_rate_id );
							// Wait for element = #subtabselector
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='DaysPerMonth Values']" )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#button-create' );
							} );
							$this->_session->element ( 'css selector', '#button-create' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Date Range Values')]" );
							} );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-20_0_0_value-20' );
							$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->sendKeys ( strval ( $value ["Start Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->sendKeys ( strval ( $value ["End Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->sendKeys ( strval ( $value ["DaysPerMonth"] ) );
							$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
						} catch ( Exception $e ) {
							$this->saveError ( $e, "ERROR: Caught exception while creating DaysPerMonth value", $key );
						}
						try {
							// Load rate data page for contract
							$this->_session->open ( $this->_maxurl . self::RATEDATAURL . $this->var_rate_id );
							// Wait for element = #subtabselector
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='DaysPerTrip Values']" )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#button-create' );
							} );
							$this->_session->element ( 'css selector', '#button-create' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Date Range Values')]" );
							} );
							// : Assert elements are present on the page
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-20_0_0_value-20' );
							$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
							// : End
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->sendKeys ( strval ( $value ["Start Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->sendKeys ( strval ( $value ["End Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->sendKeys ( strval ( $value ["DaysPerTrip"] ) );
							$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
						} catch ( Exception $e ) {
							$this->saveError ( $e, "ERROR: Caught exception while creating DaysPerTrip value", $key );
						}
						try {
							// Load rate data page for contract
							$this->_session->open ( $this->_maxurl . self::RATEDATAURL . $this->var_rate_id );
							// Wait for element = #subtabselector
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='ExpectedDistance Values']" )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#button-create' );
							} );
							$this->_session->element ( 'css selector', '#button-create' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Date Range Values')]" );
							} );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-20_0_0_value-20' );
							$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->sendKeys ( strval ( $value ["Start Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->sendKeys ( strval ( $value ["End Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->sendKeys ( strval ( $value ["ExpectedDistance"] ) );
							$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
						} catch ( Exception $e ) {
							$this->saveError ( $e, "ERROR: Caught exception while creating ExpectedDistance value", $key );
						}
						
						try {
							// Load rate data page for contract
							$this->_session->open ( $this->_maxurl . self::RATEDATAURL . $this->var_rate_id );
							// Wait for element = #subtabselector
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='ExpectedEmptyKms Values']" )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#button-create' );
							} );
							$this->_session->element ( 'css selector', '#button-create' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Date Range Values')]" );
							} );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-20_0_0_value-20' );
							$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->sendKeys ( strval ( $value ["Start Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->sendKeys ( strval ( $value ["End Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->sendKeys ( strval ( $value ["ExpectedEmptyKms"] ) );
							$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
						} catch ( Exception $e ) {
							$this->saveError ( $e, "ERROR: Caught exception while creating ExpectedEmptyKms value", $key );
						}
						
						try {
							// Load rate data page for contract
							$this->_session->open ( $this->_maxurl . self::RATEDATAURL . $this->var_rate_id );
							// Wait for element = #subtabselector
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='Fleet Values']" )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#button-create' );
							} );
							$this->_session->element ( 'css selector', '#button-create' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Date Range Values')]" );
							} );
							// : Assert elements are present on the page
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-20__0_value-20' );
							$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
							// : End
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->sendKeys ( $value ["Start Date"] );
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->sendKeys ( $value ["End Date"] );
							try {
								$myresult = $this->queryDB ( "select id from udo_fleet where ID=" . $value ["FleetValues"] );
								$fleet = $myresult [0] ["id"];
								$this->_session->element ( "xpath", "//*[@id='DateRangeValue-20__0_value-20']/option[text()='" . $fleet . "']" )->click ();
								$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
							} catch ( PHPWebDriver_NoSuchElementWebDriverError $e ) {
								$this->assertElementPresent ( 'css selector', 'input[type=submit][name=abort]' );
								$this->_session->element ( 'css selector', 'input[type=submit][name=abort]' )->click ();
							}
						} catch ( Exception $e ) {
							$this->saveError ( $e, "ERROR: Caught exception while creating Fleet value", $key );
						}
						
						try {
							// Load rate data page for contract
							$this->_session->open ( $this->_maxurl . self::RATEDATAURL . $this->var_rate_id );
							// Wait for element = #subtabselector
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
							
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='FuelConsumptionForRoute Values']" )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#button-create' );
							} );
							$this->_session->element ( 'css selector', '#button-create' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Date Range Values')]" );
							} );
							// : Assert elements are present on the page
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' );
							$this->assertElementPresent ( 'css selector', '#DateRangeValue-20_0_0_value-20' );
							$this->assertElementPresent ( 'css selector', 'input[name=save][type=submit]' );
							// : End
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-2_0_0_beginDate-2' )->sendKeys ( strval ( $value ["Start Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-4_0_0_endDate-4' )->sendKeys ( strval ( $value ["End Date"] ) );
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->clear ();
							$this->_session->element ( 'css selector', '#DateRangeValue-20_0_0_value-20' )->sendKeys ( strval ( $value ["FuelConsumption"] ) );
							$this->_session->element ( 'css selector', 'input[name=save][type=submit]' )->click ();
							$e = $w->until ( function ($session) {
								return $session->element ( 'css selector', '#subtabselector' );
							} );
						} catch ( Exception $e ) {
							$this->saveError ( $e, "ERROR: Caught exception while creating FuelConsumption value", $key );
						}
					}
				} catch ( Exception $e ) {
					$this->saveError ( $e, "ERROR: Caught in main loop.", $key );
				}
			}
			// : If errors occured. Create xls of entries that failed.
			if (count ( $this->_error ) != 0) {
				$_xlsfilename = (dirname ( __FILE__ ) . $this->_errDir . self::DS . date ( "Y-m-d_His_" ) . $this->_wdport . "_MAXLiveFandV.xlsx");
				$this->writeExcelFile ( $_xlsfilename, $this->_error, $_xlsColumns );
				if (file_exists ( $_xlsfilename )) {
					print ("Excel error report written successfully to file: $_xlsfilename") ;
				} else {
					print ("Excel error report write unsuccessful") ;
				}
			}
			// : End
		}
		
		// : End
		
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
	 * MAXLive_CreateFandVContracts::assertElementPresent($_using, $_value)
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
	 * MAXLive_CreateFandVContracts::openDB($dsn, $username, $password, $options)
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
	 * MAXLive_CreateFandVContracts::closeDB()
	 * Close connection to Database
	 */
	private function closeDB() {
		$this->_db = null;
	}
	
	/**
	 * MAXLive_CreateFandVContracts::queryDB($sqlquery)
	 * Pass MySQL Query to database and return output
	 *
	 * @param string: $sqlquery        	
	 * @param array: $result        	
	 */
	private function queryDB($sqlquery) {
		try {
			$result = $this->_db->query ( $sqlquery );
			return $result->fetchAll ( PDO::FETCH_ASSOC );
		} catch ( PDOException $ex1 ) {
			try {
				// Disconnect from database
				$this->_db = null;
				// : Reconnect to database
				$_mysqlDsn = preg_replace ( "/%s/", $this->_ip, $this->_dbdsn );
				$this->openDB ( $_mysqlDsn, $this->_dbuser, $this->_dbpwd, $this->_dboptions );
				// : End
				// Reattempt to run query
				$result = $this->_db->query ( $sqlquery );
				// Return result
				return $result->fetchAll ( PDO::FETCH_ASSOC );
			} catch (PDOException $ex2) {
				return FALSE;
			}
		}
	}
	
	/**
	 * MAXLive_CreateFandVContracts::takeScreenshot()
	 * This is a function description for a selenium test function
	 *
	 * @param object: $_session        	
	 */
	private function takeScreenshot() {
		$_img = $this->_session->screenshot ();
		$_data = base64_decode ( $_img );
		$_file = dirname ( __FILE__ ) . $this->_scrDir . DIRECTORY_SEPARATOR . date ( "Y-m-d_His" ) . "_F&VContracts.png";
		$_success = file_put_contents ( $_file, $_data );
		if ($_success) {
			return $_file;
		} else {
			return FALSE;
		}
	}
	
	/**
	 * MAXLive_CreateFandVContracts::saveError($_error, $_message, $_key)
	 * Add error message and lastRecord to error array to keep record
	 * of all errors that occur during runtime of script
	 *
	 * @param object: $_error        	
	 * @param string: $_message
	 * @param integer: $_key
	 */
	private function saveError($_error, $_message, $_key) {
		$this->takeScreenshot ();
		$_erCount = count ( $this->_error );
		$this->_error [$_erCount + 1] ["Error_Message"] = $_message . PHP_EOL . $_error->getMessage ();
		foreach ( $this->_data[$_key] as $key => $value ) {
			$this->_error [$_erCount + 1] [$key] = $value;
		}
	}
	
	/**
	 * MAXLive_CreateFandVContracts::writeExcelFile($excelFile, $excelData)
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
	
	// : End
}