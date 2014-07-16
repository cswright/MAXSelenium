<?php
// : Includes
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriver.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverWait.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverBy.php');
include_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverProxy.php');
include_once dirname ( __FILE__ ) . '/ReadExcelFile.php';
include_once 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel.php';
/**
 * PHPExcel_Writer_Excel2007
 */
include 'PHPUnit/Extensions/PHPExcel/Classes/PHPExcel/Writer/Excel2007.php';
// : End

/**
 * Object::T24Live_Weza_Main
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Barloworld Transport Solutions (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class T24Live_Weza_Main extends PHPUnit_Framework_TestCase {
	// : Constants
	const DS = DIRECTORY_SEPARATOR;
	const PB_URL = "/Planningboard";
	const COULD_NOT_CONNECT_MYSQL = "Failed to connect to MySQL database";
	const MAX_NOT_RESPONDING = "Error: MAX does not seem to be responding";
	const ROUTE_URL = "/DataBrowser?browsePrimaryObject=856&browsePrimaryInstance=";
	const LIVE_URL = "https://t24.max.bwtsgroup.com";
	const TEST_URL = "http://t24.mobilize.biz";
	const INI_FILE = "t24_data.ini";
	const INI_DIR = "ini";
	const TEST_SESSION = "firefox";
	const XLS_ERR_REPORT_AUTHOR = "T24Live_Weza_Main.php";
	const XLS_ERR_REPORT_TITLE = "Error Report";
	const XLS_ERR_REPORT_SUBJECT = "Errors and failed entries caught while processing offline trip data sheet";
	
	// : Variables
	protected static $driver;
	protected $_dummy;
	protected $_session;
	protected $lastRecord;
	protected $to = 'clintonabco@gmail.com';
	protected $subject = 'T24 Selenium script report';
	protected $message;
	protected $_errors = array ();
	protected $_processed = array ();
	protected $_dataDir;
	protected $_maxurl;
	protected $_mode;
	protected $_ip;
	protected $_xls;
	protected $_truckDescription;
	protected $_wdport;
	protected $_proxyip;
	protected $_username;
	protected $_password;
	protected $_welcome;
	protected $_browser;
	protected $_db;
	protected $_dbdsn = "mysql:host=%s;dbname=application_3;charset=utf8;";
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
	
	// : Setters
	// : End
	
	// : Getters
	// : End
	
	// : Magic
	/**
	 * T24Live_Weza_Main::__construct()
	 * Class constructor
	 */
	public function __construct() {
		$ini = dirname ( realpath ( __FILE__ ) ) . self::DS . self::INI_DIR . self::DS . self::INI_FILE;
		if (is_file ( $ini ) === FALSE) {
			echo "No " . self::INI_FILE . " file found. Please create it and populate it with the following data: username=x@y.com, password=`your password`, your name shown on MAX the welcome page welcome=`Joe Soap` and mode=`test` or `live`" . PHP_EOL;
			return FALSE;
		}
		$data = parse_ini_file ( $ini );
		if ((array_key_exists ( "screenshotdir", $data ) && $data ["screenshotdir"]) && (array_key_exists ( "errordir", $data ) && $data ["errordir"]) && (array_key_exists ( "xls", $data ) && $data ["xls"]) && (array_key_exists ( "wdport", $data ) && $data ["wdport"]) && (array_key_exists ( "datadir", $data ) && $data ["datadir"]) && (array_key_exists ( "ip", $data ) && $data ["ip"]) && (array_key_exists ( "username", $data ) && $data ["username"]) && (array_key_exists ( "password", $data ) && $data ["password"]) && (array_key_exists ( "welcome", $data ) && $data ["welcome"]) && (array_key_exists ( "mode", $data ) && $data ["mode"]) && (array_key_exists ( "truckDescription", $data ) && $data ["truckDescription"])) {
			$this->_username = $data ["username"];
			$this->_password = $data ["password"];
			$this->_welcome = $data ["welcome"];
			$this->_dataDir = $data ["datadir"];
			$this->_errDir = $data ["errordir"];
			$this->_scrDir = $data ["screenshotdir"];
			$this->_mode = $data ["mode"];
			$this->_ip = $data ["ip"];
			$this->_wdport = $data ["wdport"];
			$this->_proxyip = $data["proxy"];
			$this->_browser = $data ["browser"];
			$this->_xls = $data ["xls"];
			$this->_truckDescription = $data["truckDescription"];
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
	 * T24Live_Weza_Main::__destruct()
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
        $desired_capabilities = array();
		$proxy = new PHPWebDriver_WebDriverProxy();
		$proxy->httpProxy = $this->_proxyip;
        $proxy->add_to_capabilities($desired_capabilities);
		$this->_session = self::$driver->session ( $this->_browser, $desired_capabilities );
	}
	
	// : Selenium Test Functions
	
	/**
	 * T24Live_Weza_Main::testFunctionTemplate
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
	
	// : End
	
	// : Private Functions
	/**
	 * T24Live_Weza_Main::writeCSVFile($_data, $_file)
	 * Write CSV File of given array
	 *
	 * @param array: $_data
	 * @param string: $_file
	 */
	private function writeCSVFile($_data, $_file) {
		
	}
	
	/**
	 * T24Live_Weza_Main::readCSVFile($_file)
	 * Read CSV File of given array
	 *
	 * @param array: $_data
	 * @param string: $_file
	 */
	private function readCSVFile($_data, $_file) {
		
	}
	
	/**
	 * T24Live_Weza_Main::takeScreenshot()
	 * This is a function description for a selenium test function
	 *
	 * @param object: $_file        	
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
	 * T24Live_Weza_Main::assertElementPresent($_using, $_value)
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
	 * T24Live_Weza_Main::openDB($dsn, $username, $password, $options)
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
	 * T24Live_Weza_Main::closeDB()
	 * Close connection to Database
	 */
	private function closeDB() {
		$this->_db = null;
	}
	
	/**
	 * T24Live_Weza_Main::queryDB($sqlquery)
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