<?php
require_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriver.php');
require_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverWait.php');
require_once ('PHPUnit/Extensions/php-webdriver/PHPWebDriver/WebDriverBy.php');

/**
 * Object::web_driver_template_for_max_tests
 *
 * @author Clinton Wright
 * @author cwright@bwtsgroup.com
 * @copyright 2011 onwards Manline Group (Pty) Ltd
 * @license GNU GPL
 * @see http://www.gnu.org/copyleft/gpl.html
 */
class web_driver_template_for_max_tests extends PHPUnit_Framework_TestCase {
	// : Constants
	const DS = DIRECTORY_SEPARATOR;
	const LIVE_URL = "https://login.max.bwtsgroup.com";
	const TEST_URL = "http://max.mobilize.biz";
	
	// : Variables
	protected static $driver;
	protected $_maxurl;
	protected $_mode;
	protected $_username;
	protected $_password;
	protected $_welcome;
	
	// : Public Functions
	// : Accessors
	// : End
	
	// : Magic
	/**
	 * web_driver_template_for_max_tests::__construct()
	 * Class constructor
	 */
	public function __construct() {
		$ini = dirname ( realpath ( __FILE__ ) ) . self::DS . "user_data.ini";
		echo $ini;
		if (is_file ( $ini ) === FALSE) {
			echo "No " . self::INI_FILE . " file found. Please create it and populate it with the following data: username=x@y.com, password=`your password`, your name shown on MAX the welcome page welcome=`Joe Soap` and mode=`test` or `live`" . PHP_EOL;
			return FALSE;
		}
		$data = parse_ini_file ( $ini );
		if ((array_key_exists ( "username", $data ) && $data ["username"]) && (array_key_exists ( "password", $data ) && $data ["password"]) && (array_key_exists ( "welcome", $data ) && $data ["welcome"]) && (array_key_exists ( "mode", $data ) && $data ["mode"])) {
			$this->_username = $data ["username"];
			$this->_password = $data ["password"];
			$this->_welcome = $data ["welcome"];
			$this->_mode = $data ["mode"];
			switch ($this->_mode) {
				case "live" :
					$this->_maxurl = self::LIVE_URL;
					break;
				default :
					$this->_maxurl = self::TEST_URL;
			}
		} else {
			echo "The correct data is not present in user_data.ini. Please confirm. Fields are username, password, welcome and mode" . PHP_EOL;
			return FALSE;
		}
	}
	
	/**
	 * web_driver_template_for_max_tests::__destruct()
	 * Class destructor
	 * Allow for garbage collection
	 */
	public function __destruct() {
		unset ( $this );
	}
	// : End
	
	public function setUp() {
		self::$driver = new PHPWebDriver_WebDriver ();
	}
	
	/**
	 * web_driver_template_for_max_tests::testFunctionTemplate
	 * This is a function description for a selenium test function
	 */
	public function testFunctionTemplate() {
		$session = self::$driver->session (); // default session if not specified - firefox
		$session->setPageLoadTimeout ( 60 );
		$session->open ( $this->_maxurl );
		$w = new PHPWebDriver_WebDriverWait ( $session );
		
		// : Wait for page to load and for elements to be present on page
		$e = $w->until ( function ($session) {
			return $session->element ( 'css selector', "#contentFrame" );
		} );
		$iframe = $session->element ( 'css selector', '#contentFrame' );
		$session->switch_to_frame ( $iframe );
		$e = $w->until ( function ($session) {
			return $session->element ( 'css selector', 'input[id=identification]' );
		} );
		// : End
		
		// : Login
		$e = $session->element ( 'css selector', 'input[id=identification]' );
		$this->assertEquals ( count ( $e ), 1 );
		$e->sendKeys ( 'cwright@bwtsgroup.com' );
		$e = $session->element ( 'css selector', 'input[id=password]' );
		$e->sendKeys ( 'F@$aZ2r5StuC' );
		$e = $session->element ( 'css selector', 'input[name=submit][type=submit]' );
		$e->click ();
		// Switch out of frame
		$session->switch_to_frame ();
		
		// : Wait for page to load and for elements to be present on page
		$e = $w->until ( function ($session) {
			return $session->element ( 'css selector', "#contentFrame" );
		} );
		$iframe = $session->element ( 'css selector', '#contentFrame' );
		$session->switch_to_frame ( $iframe );
		$e = $w->until ( function ($session) {
			return $session->element ( 'xpath', "//*[contains(@href,'/logout')]" );
		} );
		// : End
		// : End
		
		// : Tear Down
		$e = $session->element ( 'xpath', "//*[contains(@href,'/logout')]" );
		$e->click ();
		// Switch out of frame
		$session->switch_to_frame ();
		
		// : Wait for page to load and for elements to be present on page
		$e = $w->until ( function ($session) {
			return $session->element ( 'css selector', "#contentFrame" );
		} );
		$iframe = $session->element ( 'css selector', '#contentFrame' );
		$session->switch_to_frame ( $iframe );
		$e = $w->until ( function ($session) {
			return $session->element ( 'css selector', 'input[id=identification]' );
		} );
		// : End
		
		$e = $session->element ( 'css selector', 'input[id=identification]' );
		$this->assertEquals ( count ( $e ), 1 );
		
		// Take a screenshot
		$this->takeScreenshot ( $session );
		
		$session->close ();
		// : End
	}
	
	// : Private Functions
	
	/**
	 * web_driver_template_for_max_tests::takeScreenshot($_session)
	 * This is a function description for a selenium test function
	 *
	 * @param object: $_session        	
	 */
	private function takeScreenshot($_session) {
		$_img = $_session->screenshot ();
		$_data = base64_decode ( $_img );
		$_file = dirname ( __FILE__ ) . DIRECTORY_SEPARATOR . "Screenshots" . DIRECTORY_SEPARATOR . date ( "Y-m-d_His" ) . "_WebDriver.png";
		$_success = file_put_contents ( $_file, $_data );
		if ($_success) {
			return $_file;
		} else {
			return FALSE;
		}
	}
	
	// : End
}
?>