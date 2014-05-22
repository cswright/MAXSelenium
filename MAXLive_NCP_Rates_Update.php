<?php
include_once 'PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriver.php';
include_once 'PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverWait.php';
include_once 'PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverBy.php';
include_once dirname ( __FILE__ ) . '/RatesReadXLSData.php';
require_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';

/**
 * Object::MAXLive_NCP_Rates_Update
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Manline Group (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class MAXLive_NCP_Rates_Update extends PHPUnit_Framework_TestCase {
	// : Constants
	const COULD_NOT_CONNECT_MYSQL = "Failed to connect to MySQL database";
	const MAX_NOT_RESPONDING = "Error: MAX does not seem to be responding";
	const PB_URL = "/Planningboard";
	const CUSTOMER_URL = "/DataBrowser?browsePrimaryObject=461&browsePrimaryInstance=";
	const LOCATION_BU_URL = "/DataBrowser?browsePrimaryObject=495&browsePrimaryInstance=";
	const OFF_CUST_BU_URL = "/DataBrowser?browsePrimaryObject=494&browsePrimaryInstance=";
	const RATEVAL_URL = "/DataBrowser?&browsePrimaryObject=udo_Rates&browsePrimaryInstance=%s&browseSecondaryObject=DateRangeValue&relationshipType=Rate";
	const DS = DIRECTORY_SEPARATOR;
	const BF = "0.00";
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
	protected $_username;
	protected $_password;
	protected $_welcome;
	protected $_mode;
	protected $_dataDir;
	protected $_errDir;
	protected $_scrDir;
	protected $_modeRates;
	protected $_modeUpdates;
	protected $_modeLocations;
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
			"select ID from udo_rates where route_id IN (select ID from udo_route where locationFrom_id IN (select ID from udo_location where name='%f') and locationTo_id IN (select ID from udo_location where name='%t')) and objectregistry_id=%g and objectInstanceId=%c and truckDescription_id=%d and enabled=1 and model='%m' and businessUnit_id=%b and rateType_id=%r;",
			"select ID from objectregistry where handle = 'udo_Customer';" 
	);
	
	// : Public functions
	// : Accessors
	
	// : End
	
	// : Magic
	/**
	 * MAXLive_NCP_Rates_Update::__construct()
	 * Class constructor
	 */
	public function __construct() {
		$ini = dirname ( realpath ( __FILE__ ) ) . self::DS . self::INI_DIR . self::DS . self::INI_FILE;
		if (is_file ( $ini ) === FALSE) {
			echo "No " . self::INI_FILE . " file found. Please refer to documentation for script to determine which fields are required and their corresponding values." . PHP_EOL;
			return FALSE;
		}
		$data = parse_ini_file ( $ini );
		if ((array_key_exists ( "updates", $data ) && $data ["updates"]) && (array_key_exists ( "rates", $data ) && $data ["rates"]) && (array_key_exists ( "locations", $data ) && $data ["locations"]) && (array_key_exists ( "xls", $data ) && $data ["xls"]) && (array_key_exists ( "errordir", $data ) && $data ["errordir"]) && (array_key_exists ( "screenshotdir", $data ) && $data ["screenshotdir"]) && (array_key_exists ( "datadir", $data ) && $data ["datadir"]) && (array_key_exists ( "ip", $data ) && $data ["ip"]) && (array_key_exists ( "username", $data ) && $data ["username"]) && (array_key_exists ( "password", $data ) && $data ["password"]) && (array_key_exists ( "welcome", $data ) && $data ["welcome"]) && (array_key_exists ( "mode", $data ) && $data ["mode"])) {
			$this->_username = $data ["username"];
			$this->_password = $data ["password"];
			$this->_welcome = $data ["welcome"];
			$this->_dataDir = $data ["datadir"];
			$this->_errDir = $data ["errordir"];
			$this->_scrDir = $data ["screenshotdir"];
			$this->_modeLocations = $data ["locations"];
			$this->_modeRates = $data ["rates"];
			$this->_modeUpdates = $data ["updates"];
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
	 * MAXLive_NCP_Rates_Update::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End
	
	/**
	 * MAXLive_NCP_Rates_Update::setUp()
	 * Create new class object and initialize session for webdriver
	 */
	public function setUp() {
		self::$driver = new PHPWebDriver_WebDriver ();
		$this->_session = self::$driver->session ( self::TEST_SESSION );
	}
	
	/**
	 * MAXLive_NCP_Rates_Update::testCreateContracts()
	 * Pull F and V Contract data and automate creation of F and V Contracts
	 */
	public function testCreateContracts() {
		$_sheetnames = ( array ) array (
				"points",
				"rates" 
		);
		// : Pull data from correctly formatted xls spreadsheet
		if ($cPR = new RatesReadXLSData ( dirname ( __FILE__ ) . self::DS . "Data" . self::DS . $this->_xls, $_sheetnames )) {
			// Get cities and save in correct naming format standard as per Meryle instruction
			$cities = $cPR->getCities ();
			// Store points
			$points = $cPR->getPoints ();
			// Store products
			$products = $cPR->getProducts ();
			// Store routes and rates
			$routes = $cPR->getRoutes ();
			
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
				
				// : Connect to database
				$_mysqlDsn = preg_replace ( "/%s/", $this->_ip, $this->_dbdsn );
				$this->openDB ( $_mysqlDsn, $this->_dbuser, $this->_dbpwd, $this->_dboptions );
				
				$myQuery = "select ID from udo_truckdescription where description='" . self::TRUCKTYPE . "';";
				$result = $this->queryDB ( $myQuery );
				$trucktype_id = $result [0] ["ID"];
				
				$myQuery = "select ID from udo_customer where tradingName='" . self::CUSTOMER . "';";
				$result = $this->queryDB ( $myQuery );
				$customer_id = $result [0] ["ID"];
				
				$myQuery = "select ID from udo_ratetype where name='Flat';";
				$result = $this->queryDB ( $myQuery );
				$rateType_id = $result [0] ["ID"];
				
				$myQuery = "select ID from udo_businessunit where name='" . self::BUNIT . "';";
				$result = $this->queryDB ( $myQuery );
				$bunit_id = $result [0] ["ID"];
				
				$myQuery = $this->_myqueries [4];
				$result = $this->queryDB ( $myQuery );
				
				// Store objectregistry_id for udo_Customer
				$objectregistry_id = $result [0] ["ID"];
				
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
					throw new Exception ( "Something went wrong when attempting to log into MAX, see error message below." . PHP_EOL . $e->getMessage () );
				}
				// : End
				
				// : Load Planningboard to rid of iframe loading on every page from here on
				$this->_session->open ( $this->_maxurl . self::PB_URL );
				$e = $w->until ( function ($session) {
					return $session->element ( "xpath", "//*[contains(text(),'You Are Here') and contains(text(), 'Planningboard')]" );
				} );
				// : End
				
				// : Main Loop
				// : Create Zones
				/*
				 * $this->load ( $this->_maxurl . "/Country_Tab/zones?&tab_id=192" ); $this->set_implicit_wait ( 60000 ); foreach ( $routes as $key => $value ) { $this->assert_element_present ( "css=div.toolbar-cell-create" ); $this->get_element ( "css=div.toolbar-cell-create" )->click (); $this->set_implicit_wait ( 60000 ); $this->assert_element_present ( "//*[contains(text(),'Create Zones')]" ); $this->assert_element_present ( "name=udo_Zone[0][name]" ); $this->assert_element_present ( "//input[@name='udo_Zone[0][fleet]' and @value='Energy (tankers)']" ); $this->assert_element_present ( "name=udo_Zone[0][country_id]" ); $this->assert_element_present ( "name=udo_Zone[0][blackoutFactor]" ); $this->assert_element_present ( "css=input[type=submit][name=save]" ); $this->get_element ( "name=udo_Zone[0][name]" )->send_keys ( $key . "kms Zone" ); $this->get_element ( "//input[@name='udo_Zone[0][fleet]' and @value='Energy (tankers)']" )->click (); $this->get_element ( "name=udo_Zone[0][country_id]" )->select_label ( self::COUNTRY ); $this->get_element ( "name=udo_Zone[0][blackoutFactor]" )->send_keys ( self::BF ); $this->get_element ( "css=input[type=submit][name=save]" )->click (); $this->set_implicit_wait ( 60000 ); }
				 */
				// : End
				
				// : Create Cities
				/*
				 * $this->load ( $this->_maxurl . "/Country_Tab/cities?&tab_id=50" ); $this->set_implicit_wait ( 600000 ); foreach ( $cities as $city ) { $this->assert_element_present ( "css=div.toolbar-cell-create" ); $this->get_element ( "css=div.toolbar-cell-create" )->click (); $this->set_implicit_wait ( 60000 ); $this->assert_element_present ( "//*[contains(text(),'Capture the details of City')]" ); $this->assert_element_present ( "name=udo_City[0][name]" ); $this->assert_element_present ( "name=udo_City[0][parent_id]" ); $this->assert_element_present ( "name=checkbox_udo_City_0_active" ); $this->assert_element_present ( "css=input[type=submit][name=save]" ); $this->get_element ( "name=udo_City[0][name]" )->send_keys ( $city ); $this->get_element ( "name=udo_City[0][parent_id]" )->select_label ( self::PROVINCE ); $this->get_element ( "name=checkbox_udo_City_0_active" )->click (); $this->get_element ( "css=input[type=submit][name=save]" )->click (); $this->set_implicit_wait ( 60000 ); $this->assert_element_present ( "css=div.toolbar-cell-create" ); $this->get_element ( "css=div.toolbar-cell-create" )->click (); $this->set_implicit_wait ( 60000 ); $this->assert_element_present ( "//*[contains(text(),'Create Zones - City')]" ); $this->assert_element_present ( "name=udo_ZoneCity_link[0][zone_id]" ); $this->assert_element_present ( "css=input[type=submit][name=save]" );
				 */
				// $zone_id = preg_split ( "/kms.*/", $city );
				/*
				 * $this->get_element ( "name=udo_ZoneCity_link[0][zone_id]" )->select_label ( $zone_id [0] . "kms Zone " . self::CONTRIB ); $this->get_element ( "css=input[type=submit][name=save]" )->click (); $this->set_implicit_wait ( 60000 ); $this->assert_element_present ( "css=input[type=submit][name=save]" ); $this->get_element ( "css=input[type=submit][name=save]" )->click (); $this->set_implicit_wait ( 600000 ); }
				 */
				// Add MySQL Query to check if record exists after it has been created
				
				// : End
				
				// : Create and link points and offloading customers to the Customer
				if ($this->_modeLocations == "true") {
					foreach ( $points as $point ) {
						$pointnames = preg_grep ( "/^" . $point ["Kms"] . "kms.*/", $cities );
						
						foreach ( $pointnames as $pointname ) {
							$pointname = preg_replace ( "/–/", "-", $pointname );
							$this->lastRecord = $point ["LocationTo"] . " (" . $pointname . ")";
							
							// Load MAX customer page
							$this->_session->open ( $this->_maxurl . self::CUSTOMER_URL . $customer_id );
							// Wait for element = #subtabselector
							$e = $w->until ( function ($session) {
								return $session->element ( "css selector", "#subtabselector" );
							} );
							// Select option from select box
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='Locations']" )->click ();
							// Wait for element = #button-create
							$e = $w->until ( function ($session) {
								return $session->element ( "css selector", "#button-create" );
							} );
							// Click element - button
							$this->_session->element ( "css selector", "#button-create" )->click ();
							// Wait for element
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[@id='udo_CustomerLocations-5__0_location_id-5']" );
							} );
							$this->assertElementPresent ( "link text", "Create Location" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_CustomerLocations-5__0_location_id-5']" );
							$this->_session->element ( "link text", "Create Location" )->click ();
							$this->selectWindow ( "Create Location" );
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the TYPE of Location')]" );
							} );
							$this->assertElementPresent ( "css selector", "#udo_Location-17__0__type-17" );
							$this->assertElementPresent ( "css selector", "input[name=save][type=submit]" );
							$this->_session->element ( "xpath", "//*[@id='udo_Location-17__0__type-17']/option[text()='Point']" )->click ();
							$this->_session->element ( "css selector", "input[name=save][type=submit]" )->click ();
							// Wait for element = Page heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the details of Point')]" );
							} );
							// Assert all elements on current page are present
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Point-14_0_0_name-14']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Point-15__0_parent_id-15']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Point-32_0_0_pointType_id-32[2]']" );
							$this->assertElementPresent ( "xpath", "//*[@id='checkbox_udo_Point-2_0_0_active-2']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							// Enter name of new location in text field
							$this->_session->element ( "xpath", "//*[@id='udo_Point-14_0_0_name-14']" )->sendKeys ( $point ["LocationTo"] . " (" . $pointname . ")" );
							// Select parent location from select box
							$this->_session->element ( "xpath", "//*[@id='udo_Point-15__0_parent_id-15']/option[text()='" . self::PROVINCE . " -- " . $pointname . "']" )->click ();
							// Check the offloading point checkbox
							$this->_session->element ( "xpath", "//*[@id='udo_Point-32_0_0_pointType_id-32[2]']" )->click ();
							// Check the active checkbox
							$this->_session->element ( "xpath", "//*[@id='checkbox_udo_Point-2_0_0_active-2']" )->click ();
							// Click the submit button
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							$this->selectWindow ( "Create Customer Locations" );
							// Wait for element
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the details of Customer Locations')]" );
							} );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_CustomerLocations-5__0_location_id-5']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_CustomerLocations-8__0_type-8']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							
							// Select new location from select box
							$this->_session->element ( "xpath", "//*[@id='udo_CustomerLocations-5__0_location_id-5']/option[text()='" . $point ["LocationTo"] . " (" . $pointname . ")" . "']" )->click ();
							// Select offloading as type for new location from select box
							$this->_session->element ( "xpath", "//*[@id='udo_CustomerLocations-8__0_type-8']/option[text()='Offloading']" )->click ();
							// Click the submit button
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							// : Create Business Unit Link for Point Link
							$myQuery = preg_replace ( "/%n/", $point ["LocationTo"] . " (" . $pointname . ")", $this->_myqueries [0] );
							$myQuery = preg_replace ( "/%t/", self::CUSTOMER, $myQuery );
							$result = $this->queryDB ( $myQuery );
							$location_id = $result [0] ["ID"];
							$this->_session->open ( $this->_maxurl . self::LOCATION_BU_URL . $location_id );
							// Wait for element
							$e = $w->until ( function ($session) {
								return $session->element ( "css selector", "#button-create" );
							} );
							// Click element = button-create
							$this->_session->element ( "css selector", "#button-create" )->click ();
							
							// Wait for element = Page heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Customer Locations - Business Unit')]" );
							} );
							
							$this->assertElementPresent ( "xpath", "//*[@id='udo_CustomerLocationsBusinessUnit_link-2__0_businessUnit_id-2']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							
							$this->_session->element ( "xpath", "//*[@id='udo_CustomerLocationsBusinessUnit_link-2__0_businessUnit_id-2']/option[text()='" . self::BUNIT . "']" )->click ();
							// Click the submit button
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							// : Load customer data browser page for Customer
							$this->_session->open ( $this->_maxurl . self::CUSTOMER_URL . $customer_id );
							
							// : Create and link Offloading Customer
							
							// Wait for element = Page heading
							$e = $w->until ( function ($session) {
								return $session->element ( "css selector", "#subtabselector" );
							} );
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='Offloading Customers where Customer is " . self::CUSTOMER . "']" )->click ();
							
							// Wait for element = Page heading
							$e = $w->until ( function ($session) {
								return $session->element ( "css selector", "#button-create" );
							} );
							$this->_session->element ( "css selector", "#button-create" )->click ();
							
							// Wait for element = Page heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the details of Offloading Customers')]" );
							} );
							
							$this->assertElementPresent ( "link text", "Create Customer" );
							$this->_session->element ( "link text", "Create Customer" )->click ();
							
							$this->selectWindow ( "Create Customer" );
							
							// Wait for element = Page heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Customer')]" );
							} );
							
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Customer-22_0_0_tradingName-22']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Customer-12_0_0_legalName-12']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Customer-32_0_0_customerType_id-32[11]']" );
							$this->assertElementPresent ( "xpath", "//*[@id='checkbox_udo_Customer-2_0_0_active-2']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							
							$this->_session->element ( "xpath", "//*[@id='udo_Customer-22_0_0_tradingName-22']" )->sendKeys ( $point ["LocationTo"] . " (" . $pointname . ")" );
							$this->_session->element ( "xpath", "//*[@id='udo_Customer-12_0_0_legalName-12']" )->sendKeys ( $point ["LocationTo"] . " (" . $pointname . ")" );
							$this->_session->element ( "xpath", "//*[@id='udo_Customer-32_0_0_customerType_id-32[11]']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='checkbox_udo_Customer-2_0_0_active-2']" )->click ();
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							$this->selectWindow ( "Create Offloading Customer" );
							
							// Wait for element = Page heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the details of Offloading Customers')]" );
							} );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_OffloadingCustomers-3__0_customer_id-3']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_OffloadingCustomers-6__0_offloadingCustomer_id-6']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							
							$this->_session->element ( "xpath", "//*[@id='udo_OffloadingCustomers-3__0_customer_id-3']/option[text()='" . self::CUSTOMER . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_OffloadingCustomers-6__0_offloadingCustomer_id-6']/option[text()='" . $point ["LocationTo"] . " (" . $pointname . ")" . "']" )->click ();
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							// : Create Business Unit Link for Offloading Customer Link
							$myQuery = preg_replace ( "/%o/", $point ["LocationTo"] . " (" . $pointname . ")", $this->_myqueries [1] );
							$myQuery = preg_replace ( "/%t/", self::CUSTOMER, $myQuery );
							$result = $this->queryDB ( $myQuery );
							$offloadingcustomer_id = $result [0] ["ID"];
							$this->_session->open ( $this->_maxurl . self::OFF_CUST_BU_URL . $offloadingcustomer_id );
							
							// Wait for element = #subtabselector
							$e = $w->until ( function ($session) {
								return $session->element ( "css selector", "#subtabselector" );
							} );
							$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='Offloading Customers - Business Unit']" );
							
							// Wait for element = #button-create
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//div[@id='button-create']" );
							} );
							$this->_session->element ( "xpath", "//div[@id='button-create']" )->click ();
							
							// Wait for element = Page Heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Create Offloading Customers - Business Unit')]" );
							} );
							
							$this->assertElementPresent ( "xpath", "//*[@id='udo_OffloadingCustomersBusinessUnit_link-2__0_businessUnit_id-2']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							
							$this->_session->element ( "xpath", "//*[@id='udo_OffloadingCustomersBusinessUnit_link-2__0_businessUnit_id-2']/option[text()='" . self::BUNIT . "']" )->click ();
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							// Wait for element = #button-create
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//div[@id='button-create']" );
							} );
						}
					}
				}
				
				// Add MySQL Query to check if record exists after it has been created
				
				// : Create
				
				// : End
				
				// : Create Routes, Rates and Rate Values
				if ($this->_modeRates == "true") {
					foreach ( $cities as $pointname ) {
						
						$kms = preg_split ( "/kms Zone.*/", $pointname );
						$kms = $kms [0];
						// Load the MAX customer page
						$this->_session->open ( $this->_maxurl . self::CUSTOMER_URL . $customer_id );
						
						// Wait for element = #subtabselector
						$e = $w->until ( function ($session) {
							return $session->element ( "css selector", "#subtabselector" );
						} );
						// Select Rates from the select box
						$this->_session->element ( "xpath", "//*[@id='subtabselector']/select/option[text()='Rates']" )->click ();
						
						$myQuery = "select name from udo_location where ID IN (select parent_id from udo_location where name='" . $points [1] ["LocationFrom"] . "');";
						$result = $this->queryDB ( $myQuery );
						if (count($result) != 0) {
							$locationFrom_id = $result [0] ["name"];
						} else {
							$locationFrom_id = $points [1] ["LocationFrom"];
						}
						
						$pointname = preg_replace ( "/–/", "-", $pointname );
						if ($this->_modeUpdates == "false") {
							// Wait for element = #button-create
							$e = $w->until ( function ($session) {
								return $session->element ( "css selector", "#button-create" );
							} );
							// Click element - #button-create
							$this->_session->element ( "css selector", "#button-create" )->click ();
							
							// Wait for element Page Heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the details of Rates')]" );
							} );
							
							$this->assertElementPresent ( "link text", "Create Route" );
							$this->_session->element ( "link text", "Create Route" )->click ();
							$this->selectWindow ( "Create Route" );
							
							// Wait for element Page Heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the details of Route')]" );
							} );
							
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Route-6__0_locationFrom_id-6']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Route-7__0_locationTo_id-7']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Route-4_0_0_expectedKms-4']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Route-3_0_0_duration-3']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							
							try {
								$this->_session->element ( "xpath", "//*[@id='udo_Route-6__0_locationFrom_id-6']/option[text()=' . $locationFrom_id . ']" )->click ();
							} catch ( Exception $e ) {
								$locationFrom_id = $points [1] ["LocationFrom"];
								$this->_session->element ( "xpath", "//*[@id='udo_Route-6__0_locationFrom_id-6']/option[text()=' . $locationFrom_id . ']" )->click ();
							}
							
							$this->_session->element ( "xpath", "//*[@id='udo_Route-7__0_locationTo_id-7']/option[text()='" . $pointname . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_Route-4_0_0_expectedKms-4']" )->sendKeys ( $kms );
							// Calculate duration from kms value at 60K/H
							$duration = strval ( number_format ( (floatval ( $kms ) / 80) * 60, 0, "", "" ) );
							$this->_session->element ( "xpath", "//*[@id='udo_Route-3_0_0_duration-3']" )->sendKeys ( $duration );
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							$this->selectWindow ( "Create Rates" );
							
							// Wait for element Page Heading
							$e = $w->until ( function ($session) {
								return $session->element ( "xpath", "//*[contains(text(),'Capture the details of Rates')]" );
							} );
							
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Rates-31__0_route_id-31']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Rates-30__0_rateType_id-30']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Rates-4__0_businessUnit_id-4']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Rates-36__0_truckDescription_id-36']" );
							$this->assertElementPresent ( "xpath", "//*[@id='udo_Rates-20__0_model-20']" );
							$this->assertElementPresent ( "xpath", "//*[@id='checkbox_udo_Rates-15_0_0_enabled-15']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-31__0_route_id-31']/option[text()='" . $locationFrom_id . " TO " . $pointname . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-30__0_rateType_id-30']/option[text()='Flat']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-4__0_businessUnit_id-4']/option[text()='" . self::BUNIT . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-36__0_truckDescription_id-36']/option[text()='" . self::TRUCKTYPE . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='udo_Rates-20__0_model-20']/option[text()='" . self::CONTRIB . "']" )->click ();
							$this->_session->element ( "xpath", "//*[@id='checkbox_udo_Rates-15_0_0_enabled-15']" )->click ();
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							sleep ( 5 );
						}
						
						// : Create Rate Value for Route
						// $myQuery = preg_replace ( "/%f/", $points [1] ["LocationFrom"], $this->_myqueries [3] );
						$myQuery = preg_replace ( "/%f/", $locationFrom_id, $this->_myqueries [3] );
						$myQuery = preg_replace ( "/%t/", $pointname, $myQuery );
						$myQuery = preg_replace ( "/%g/", $objectregistry_id, $myQuery );
						$myQuery = preg_replace ( "/%c/", $customer_id, $myQuery );
						$myQuery = preg_replace ( "/%d/", $trucktype_id, $myQuery );
						$myQuery = preg_replace ( "/%m/", self::CONTRIB, $myQuery );
						$myQuery = preg_replace ( "/%b/", $bunit_id, $myQuery );
						$myQuery = preg_replace ( "/%r/", $rateType_id, $myQuery );
						$result = $this->queryDB ( $myQuery );
						if (count ( $result ) != 0) {
							$rate_id = $result [0] ["ID"];
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
							$this->assertElementPresent ( "xpath", "//*[@id='DateRangeValue-20_0_0_value-20']" );
							$this->assertElementPresent ( "css selector", "input[type=submit][name=save]" );
							
							$this->_session->element ( "xpath", "//*[@id='DateRangeValue-2_0_0_beginDate-2']" )->clear ();
							$this->_session->element ( "xpath", "//*[@id='DateRangeValue-2_0_0_beginDate-2']" )->sendKeys ( date ( "Y-m-01 00:00:00" ) );
							$productname = preg_split ( "/^" . $kms . "kms Zone /", $pointname );
							$ratevalue = strval ( (number_format ( floatval ( $routes [$kms] [$productname [1]] ), 2, ".", "" )) );
							$this->_session->element ( "xpath", "//*[@id='DateRangeValue-20_0_0_value-20']" )->sendKeys ( $ratevalue );
							$this->_session->element ( "css selector", "input[type=submit][name=save]" )->click ();
							
							sleep ( 2 );
						} else {
							$_erCount = count ( $this->_error );
							$this->_error [$_erCount + 1] ["error"] = $e->getMessage ();
							$this->_error [$_erCount + 1] ["record"] = $this->lastRecord;
						}
					}
				}
				// : End
				// : End
				
			} catch ( Exception $e ) {
				echo "Error: " . $e->getMessage () . PHP_EOL;
				echo "Time of error: " . date ( "Y-m-d H:i:s" ) . PHP_EOL;
				echo "Last record: " . $this->lastRecord;
				$this->takeScreenshot ();
				$_erCount = count ( $this->_error );
				$this->_error [$_erCount + 1] ["error"] = $e->getMessage ();
				$this->_error [$_erCount + 1] ["record"] = $this->lastRecord;
			}
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
	 * MAXLive_NCP_Rates_Update::openDB($dsn, $username, $password, $options)
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
	 * MAXLive_NCP_Rates_Update::assertElementPresent($_using, $_value)
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
	 * MAXLive_NCP_Rates_Update::assertElementPresent($_title)
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
	 * MAXLive_NCP_Rates_Update::closeDB()
	 * Close connection to Database
	 */
	private function closeDB() {
		$this->_db = null;
	}
	
	/**
	 * MAXLive_NCP_Rates_Update::queryDB($sqlquery)
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