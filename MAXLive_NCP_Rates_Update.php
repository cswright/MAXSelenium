<?php
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriver.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverWait.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverBy.php');
include_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';
include_once dirname ( __FILE__ ) . '/RatesReadXLSData.php';

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
	const DATAFILE = "2014-02-03_NCPData.xls";
	const PROVINCE = "Africa -- South Africa -- KZN";
	const TRUCKTYPE = "Fuel Tanker";
	
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
	protected $_files;
	protected $_error = array ();
	protected $_db;
	protected $_dbdsn = "mysql:host=192.168.1.43;dbname=max2;charset=utf8;";
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
		// : Pull data from correctly formatted xls spreadsheet
		if ($cPR = new RatesReadXLSData ( dirname ( __FILE__ ) . self::DS . "Data" . self::DS . $this->_xls )) {
			// Get cities and save in correct naming format standard as per Meryle instruction
			$cities = $cPR->getCities (); 
			// Store points
			$points = $cPR->getPoints (); 
			// Store products
			$products = $cPR->getProducts ();
			// Store routes and rates
			$routes = $cPR->getRoutes (); 
			
			try {
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
				$this->set_implicit_wait ( 15000 );
				$this->load ( $this->_maxurl . "/home" );
				$this->set_implicit_wait ( 60000 );
				$this->is_element_present ( "id=identification" );
				$this->is_element_present ( "id=password" );
				$this->is_element_present ( "name=submit" );
				$this->get_element ( "id=identification" )->send_keys ( $this->_username );
				$this->get_element ( "id=password" )->send_keys ( $this->_password );
				$this->get_element ( "name=submit" )->click ();
				$this->set_implicit_wait ( 60000 );
				$this->assert_string_present ( "Welcome " . $this->_welcome );
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
				foreach ( $points as $point ) {
					$pointnames = preg_grep ( "/^" . $point ["Kms"] . "kms.*/", $cities );
					
					foreach ( $pointnames as $pointname ) {
						$pointname = preg_replace ( "/–/", "-", $pointname );
						$this->lastRecord = $point ["LocationTo"] . " (" . $pointname . ")";
						$this->load ( $this->_maxurl . self::CUSTOMER_URL . $customer_id );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "css=span#subtabselector" );
						$this->get_element ( "css=span#subtabselector" )->select_label_old ( "Locations" );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "css=div#button-create" );
						$this->get_element ( "css=div#button-create" )->click ();
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//*[contains(text(),'Capture the details of Customer Locations')]" );
						$this->assert_element_present ( "link=Create Location" );
						$this->assert_element_present ( "name=udo_CustomerLocations[0][location_id]" );
						$this->get_element ( "link=Create Location" )->click ();
						$this->set_implicit_wait ( 60000 );
						$this->select_window_pattern ( "Create Location" );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//*[contains(text(),'Capture the TYPE of Location')]" );
						$this->assert_element_present ( "name=udo_Location[0][_type]" );
						$this->assert_element_present ( "css=input[type=submit][name=save]" );
						$this->get_element ( "name=udo_Location[0][_type]" )->select_label ( "Point" );
						$this->get_element ( "css=input[type=submit][name=save]" )->click ();
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "name=udo_Point[0][name]" );
						$this->assert_element_present ( "name=udo_Point[0][parent_id]" );
						$this->assert_element_present ( "name=udo_Point[0][pointType_id][2]" );
						$this->assert_element_present ( "name=checkbox_udo_Point_0_active" );
						$this->assert_element_present ( "css=input[type=submit][name=save]" );
						$this->get_element ( "name=udo_Point[0][name]" )->send_keys ( $point ["LocationTo"] . " (" . $pointname . ")" );
						$this->get_element ( "name=udo_Point[0][parent_id]" )->select_label ( self::PROVINCE . " -- " . $pointname );
						$this->get_element ( "name=udo_Point[0][pointType_id][2]" )->click ();
						$this->get_element ( "name=checkbox_udo_Point_0_active" )->click ();
						$this->get_element ( "css=input[type=submit][name=save]" )->click ();
						$this->select_window_pattern ( "Create Customer Locations" );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//*[contains(text(),'Capture the details of Customer Locations')]" );
						$this->assert_element_present ( "name=udo_CustomerLocations[0][location_id]" );
						$this->assert_element_present ( "name=udo_CustomerLocations[0][type]" );
						$this->assert_element_present ( "css=input[type=submit][name=save]" );
						$this->get_element ( "name=udo_CustomerLocations[0][location_id]" )->select_label ( $point ["LocationTo"] . " (" . $pointname . ")" );
						$this->get_element ( "name=udo_CustomerLocations[0][type]" )->select_label ( "Offloading" );
						$this->get_element ( "css=input[type=submit][name=save]" )->click ();
						$this->set_implicit_wait ( 60000 );
						
						// : Create Business Unit Link for Point Link
						$myQuery = preg_replace ( "/%n/", $point ["LocationTo"] . " (" . $pointname . ")", $this->_myqueries [0] );
						$myQuery = preg_replace ( "/%t/", self::CUSTOMER, $myQuery );
						$result = $this->queryDB ( $myQuery );
						$location_id = $result [0] ["ID"];
						$this->load ( $this->_maxurl . self::LOCATION_BU_URL . $location_id );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "css=div#button-create" );
						$this->get_element ( "css=div#button-create" )->click ();
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//*[contains(text(),'Create Customer Locations - Business Unit')]" );
						$this->assert_element_present ( "name=udo_CustomerLocationsBusinessUnit_link[0][businessUnit_id]" );
						$this->assert_element_present ( "css=input[type=submit][name=save]" );
						$this->get_element ( "name=udo_CustomerLocationsBusinessUnit_link[0][businessUnit_id]" )->select_label ( self::BUNIT );
						$this->get_element ( "css=input[type=submit][name=save]" )->click ();
						$this->set_implicit_wait ( 60000 );
						
						// : Load customer data browser page for Customer
						$this->load ( $this->_maxurl . self::CUSTOMER_URL . $customer_id );
						$this->set_implicit_wait ( 60000 );
						
						// : Create and link Offloading Customer
						$this->assert_element_present ( "css=span#subtabselector" );
						$this->get_element ( "css=span#subtabselector" )->select_label_old ( "Offloading Customers where Customer is " . self::CUSTOMER );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "css=div#button-create" );
						$this->get_element ( "css=div#button-create" )->click ();
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//*[contains(text(),'Capture the details of Offloading Customers')]" );
						$this->assert_element_present ( "link=Create Customer" );
						$this->get_element ( "link=Create Customer" )->click ();
						$this->select_window_pattern ( "Create Customer" );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//*[contains(text(),'Create Customer')]" );
						$this->assert_element_present ( "name=udo_Customer[0][tradingName]" );
						$this->assert_element_present ( "name=udo_Customer[0][legalName]" );
						$this->assert_element_present ( "name=udo_Customer[0][customerType_id][11]" );
						$this->assert_element_present ( "name=checkbox_udo_Customer_0_active" );
						$this->assert_element_present ( "css=input[type=submit][name=save]" );
						$this->get_element ( "name=udo_Customer[0][tradingName]" )->send_keys ( $point ["LocationTo"] . " (" . $pointname . ")" );
						$this->get_element ( "name=udo_Customer[0][legalName]" )->send_keys ( $point ["LocationTo"] . " (" . $pointname . ")" );
						$this->get_element ( "name=udo_Customer[0][customerType_id][11]" )->click ();
						$this->get_element ( "name=checkbox_udo_Customer_0_active" )->click ();
						$this->get_element ( "css=input[type=submit][name=save]" )->click ();
						$this->select_window_pattern ( "Create Offloading Customers" );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//*[contains(text(),'Capture the details of Offloading Customers')]" );
						$this->assert_element_present ( "name=udo_OffloadingCustomers[0][customer_id]" );
						$this->assert_element_present ( "name=udo_OffloadingCustomers[0][offloadingCustomer_id]" );
						$this->assert_element_present ( "css=input[type=submit][name=save]" );
						$this->get_element ( "name=udo_OffloadingCustomers[0][customer_id]" )->select_label ( self::CUSTOMER );
						$this->get_element ( "name=udo_OffloadingCustomers[0][offloadingCustomer_id]" )->select_label ( $point ["LocationTo"] . " (" . $pointname . ")" );
						$this->get_element ( "css=input[type=submit][name=save]" )->click ();
						$this->set_implicit_wait ( 60000 );
						// : Create Business Unit Link for Offloading Customer Link
						$myQuery = preg_replace ( "/%o/", $point ["LocationTo"] . " (" . $pointname . ")", $this->_myqueries [1] );
						$myQuery = preg_replace ( "/%t/", self::CUSTOMER, $myQuery );
						$result = $this->queryDB ( $myQuery );
						$offloadingcustomer_id = $result [0] ["ID"];
						$this->load ( $this->_maxurl . self::OFF_CUST_BU_URL . $offloadingcustomer_id );
						$this->set_implicit_wait ( 60000 );
						$this->get_element_present ( "css=span#subtabselector" )->select_label_old ( "Offloading Customers - Business Unit" );
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//div[@id='button-create']" );
						$this->get_element ( "//div[@id='button-create']" )->click ();
						$this->set_implicit_wait ( 60000 );
						$this->assert_element_present ( "//*[contains(text(),'Create Offloading Customers - Business Unit')]" );
						$this->assert_element_present ( "name=udo_OffloadingCustomersBusinessUnit_link[0][businessUnit_id]" );
						$this->assert_element_present ( "css=input[type=submit][name=save]" );
						$this->get_element ( "name=udo_OffloadingCustomersBusinessUnit_link[0][businessUnit_id]" )->select_label ( self::BUNIT );
						$this->get_element ( "css=input[type=submit][name=save]" )->click ();
						$this->set_implicit_wait ( 60000 );
					}
				}
				
				// Add MySQL Query to check if record exists after it has been created
				
				// : Create
				
				// : End
				
				// : Create Routes, Rates and Rate Values
				
				foreach ( $cities as $pointname ) {
					$kms = preg_split ( "/kms Zone.*/", $pointname );
					$kms = $kms [0];
					$this->load ( $this->_maxurl . self::CUSTOMER_URL . $customer_id );
					$this->set_implicit_wait ( 60000 );
					$this->assert_element_present ( "css=span#subtabselector" );
					$this->get_element ( "css=span#subtabselector" )->select_label_old ( "Rates" );
					$this->set_implicit_wait ( 60000 );
					$pointname = preg_replace ( "/–/", "-", $pointname );
					$this->assert_element_present ( "css=div#button-create" );
					$this->get_element ( "css=div#button-create" )->click ();
					$this->set_implicit_wait ( 60000 );
					$this->assert_element_present ( "//*[contains(text(),'Capture the details of Rates')]" );
					$this->assert_element_present ( "link=Create Route" );
					$this->get_element ( "link=Create Route" )->click ();
					$this->select_window_pattern ( "Create Route" );
					$this->set_implicit_wait ( 60000 );
					$this->assert_element_present ( "//*[contains(text(),'Capture the details of Route')]" );
					$this->assert_element_present ( "name=udo_Route[0][locationFrom_id]" );
					$this->assert_element_present ( "name=udo_Route[0][locationTo_id]" );
					$this->assert_element_present ( "name=udo_Route[0][expectedKms]" );
					$this->assert_element_present ( "name=udo_Route[0][duration]" );
					$this->assert_element_present ( "css=input[type=submit][name=save]" );
					$myQuery = "select name from udo_location where ID IN (select parent_id from udo_location where name='" . $points [1] ["LocationFrom"] . "');";
					$result = $this->queryDB ( $myQuery );
					$locationFrom_id = $result [0] ["name"];
					try {
						$this->get_element ( "name=udo_Route[0][locationFrom_id]" )->select_label ( $locationFrom_id );
					} catch ( Exception $e ) {
						$locationFrom_id = $points [1] ["LocationFrom"];
						$this->get_element ( "name=udo_Route[0][locationFrom_id]" )->select_label ( $locationFrom_id );
					}
					// $this->get_element ( "name=udo_Route[0][locationFrom_id]" )->select_label ( $points [1] ["LocationFrom"] );
					$this->get_element ( "name=udo_Route[0][locationTo_id]" )->select_label ( $pointname );
					$this->get_element ( "name=udo_Route[0][expectedKms]" )->send_keys ( $kms );
					$duration = strval ( number_format ( (floatval ( $kms ) / 80) * 60, 0, "", "" ) ); // Calculate duration from kms value at 60K/H
					$this->get_element ( "name=udo_Route[0][duration]" )->send_keys ( strval ( $duration ) );
					$this->get_element ( "css=input[type=submit][name=save]" )->click ();
					$this->select_window_pattern ( "Create Rates" );
					$this->set_implicit_wait ( 60000 );
					$this->assert_element_present ( "//*[contains(text(),'Capture the details of Rates')]" );
					$this->assert_element_present ( "name=udo_Rates[0][route_id]" );
					$this->assert_element_present ( "name=udo_Rates[0][rateType_id]" );
					$this->assert_element_present ( "name=udo_Rates[0][businessUnit_id]" );
					$this->assert_element_present ( "name=udo_Rates[0][truckDescription_id]" );
					$this->assert_element_present ( "name=udo_Rates[0][model]" );
					$this->assert_element_present ( "name=checkbox_udo_Rates_0_enabled" );
					$this->assert_element_present ( "css=input[type=submit][name=save]" );
					// $this->get_element ( "name=udo_Rates[0][route_id]" )->select_label ( $points [1] ["LocationFrom"] . " TO " . $pointname );
					$this->get_element ( "name=udo_Rates[0][route_id]" )->select_label ( $locationFrom_id . " TO " . $pointname );
					$this->get_element ( "name=udo_Rates[0][rateType_id]" )->select_label ( "Flat" );
					$this->get_element ( "name=udo_Rates[0][businessUnit_id]" )->select_label ( self::BUNIT );
					$this->get_element ( "name=udo_Rates[0][truckDescription_id]" )->select_label ( self::TRUCKTYPE );
					$this->get_element ( "name=udo_Rates[0][model]" )->select_label ( self::CONTRIB );
					$this->get_element ( "name=checkbox_udo_Rates_0_enabled" )->click ();
					$this->get_element ( "css=input[type=submit][name=save]" )->click ();
					$this->set_implicit_wait ( 60000 );
					
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
					$rate_id = $result [0] ["ID"];
					$rateurl = preg_replace ( "/%s/", $rate_id, $this->_maxurl . self::RATEVAL_URL );
					$this->load ( $rateurl );
					$this->set_implicit_wait ( 60000 );
					$this->assert_element_present ( "css=div#button-create" );
					$this->get_element ( "css=div#button-create" )->click ();
					$this->set_implicit_wait ( 60000 );
					$this->assert_element_present ( "//*[contains(text(),'Create Date Range Values')]" );
					$this->assert_element_present ( "name=DateRangeValue[0][beginDate]" );
					$this->assert_element_present ( "name=DateRangeValue[0][value]" );
					$this->assert_element_present ( "css=input[type=submit][name=save]" );
					$this->get_element ( "name=DateRangeValue[0][beginDate]" )->clear ();
					$this->get_element ( "name=DateRangeValue[0][beginDate]" )->send_keys ( date ( "Y-m-01 00:00:00" ) );
					$productname = preg_split ( "/^" . $kms . "kms Zone /", $pointname );
					$ratevalue = strval ( (number_format ( floatval ( $routes [$kms] [$productname [1]] ), 2, ".", "" )) );
					$this->get_element ( "name=DateRangeValue[0][value]" )->send_keys ( $ratevalue );
					$this->get_element ( "css=input[type=submit][name=save]" )->click ();
					$this->set_implicit_wait ( 60000 );
				}
				// : End
				// : End
				
				// : Teardown
				$this->load ( $this->_maxurl . "/logout" );
				$this->set_implicit_wait ( 60000 );
				$this->is_element_present ( "id=identification" );
				// Close database connection
				$db = null;
			} catch ( Exception $e ) {
				echo "Error: " . $e->getMessage () . PHP_EOL;
				echo "Time of error: " . date ( "Y-m-d H:i:s" ) . PHP_EOL;
				echo "Last record: " . $this->lastRecord;
				$this->message = "Error: " . $e->getMessage ();
				$this->message = wordwrap ( $this->message, 70, "\r\n" );
				mail ( $this->to, $this->subject, $this->message );
				$db = null;
				$this->TakeScreenshot ();
			}
			// : End
		}
		else {
			print("Error: The excel spreadsheet, '" . $this->_xls . "', failed to load." . PHP_EOL);
		}
	}
	
	/**
	 * MAXLive_NCP_Rates_Update::tearDown()
	 * tear down instance
	 */
	public function tearDown() {
		if ($this->driver) {
			if ($this->hasFailed ()) {
				$this->driver->set_sauce_context ( "passed", false );
			} else {
				$this->driver->set_sauce_context ( "passed", true );
			}
			$this->driver->quit ();
		}
		parent::tearDown ();
	}
	
	// : Private Functions
	
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
	 * MAXLive_NCP_Rates_Update::TakeScreenshot()
	 * Take a screenshot of the browser content
	 */
	private function TakeScreenshot() {
		file_put_contents ( "screenshot-" . date ( "Y-m-d H:i:s" ) . ".png", $this->get_screenshot () );
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