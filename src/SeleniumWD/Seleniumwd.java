package SeleniumWD;

/*######################################################################################################
 'Project Name		: Open2Test framework for Selenium Web Driver
 'Author		: Open2Test
 'Version	    	: V 3.0.9
 'Date of Creation	: 18-August-2015
 '######################################################################################################
 */
//V2- Enhanced with context support
//V3 - Enahanced with FileUploadDownload Management
import static org.junit.Assert.*;

import org.junit.After;
import org.junit.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.Alert; //Updated on 16.12.2013 for Dialog support
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.JavascriptExecutor;
//import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;//Updated on 16.12.2013 for Dialog support
//import org.openqa.selenium.NoSuchElementException;
//import org.openqa.selenium.NoSuchWindowException;
import org.openqa.selenium.OutputType;
//import org.openqa.selenium.StaleElementReferenceException;
//import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
//import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
//import org.openqa.selenium.interactions.Actions;
//import org.openqa.selenium.android.AndroidDriver;
//import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.RemoteWebDriver;



//import java.awt.event.ActionEvent;
import java.io.*;

import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
//import org.openqa.selenium.safari.SafariDriver;
//import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
//import org.openqa.selenium.support.ui.WebDriverWait;



//import autoitx4java.AutoItX;
//import com.thoughtworks.selenium.Selenium;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
//import java.util.Calendar;
//import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet; //Updated on 16.12.2013 for page support
import java.util.Iterator; //Updated on 16.12.2013 for page support
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set; //Updated on 16.12.2013 for page support
import java.util.Stack;
//import java.util.Stack;
import java.util.concurrent.TimeUnit;
import java.lang.System;
import java.lang.reflect.Method;
//import java.lang.reflect.Method;
//import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit; // Updated 11.10.2014 screenshots
//import java.awt.datatransfer.Clipboard;
//import java.awt.datatransfer.DataFlavor;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;

import javax.imageio.ImageIO;

public class SeleniumWD {
	int scriptfound = 0;
	Setting setObj = new Setting(); // Updated on 16.12.2013
	static BufferedWriter bw = null;
	static BufferedWriter bw1 = null;
	static RemoteWebDriver D8;
	String browserver = null;
	String mm = null;
	String dd = null;
	String yyyy = null;
	String monthpart;
	String monthpartjap;
	Date cur_dt = null;
	String TestSuite;
	String TestScript;
	static String globalCompName;
	static String ObjectRepository;
	int startrow = 0;
	static String ReportsPath;
	String strResultPath;
	String[] TCNm = null;
	static String exeStatus = "True";
	static String ScreenReport;
	static int rowcnt;
	static int dtrownum = 1;
	static int ORrowcount = 0;
	static String ORvalname = "";
	static String ORvalue = "";
	static String Action = "";
	static String cCellData = "";
	static String dCellData = "";
	static String htmlname = "";
	static String cCellObjType = "";// included on 29-Nov-2013 for storing
	static String cCellObjName = ""; // included on 29-Nov-2013 for storing
										// object name
	String[] cCellDataVal = null;
	String[] dCellDataVal = null;
	String ObjectSet;
	String[] ObjectSet1 = null;
	String ObjectSet2 = "";
	String ObjectSetVal = "";
	static Sheet DTsheet = null;
	static Sheet ORsheet;
	String Searchtext;
	String parentWindowHandle;// Updated on 16.12.2013 for page support
	String parentWindowHandle1;
	String windowHandle;// Updated on 16.12.2013 for page support
	String ieDriverPath;// Updated on 16.12.2013 for outside browser setting
	String tempTestReportPath;// Updated on 16.12.2013 for outside temporary
								// report path

	static int iflag = 0;

	static int screenshotflag = 0;
	static int loopflag = 0;
	static int j = 0;
	static int loopsize = -1;
	int[] loopstart = new int[1];
	int[] loopcount = new int[1];
	int[] loopend = new int[1];
	static int[] loopcnt = new int[1];
	static int[] dtrownumloop = new int[1];
	boolean captureperform = false;
	static boolean capturecheckvalue = false;
	static boolean capturestorevalue = false;
	static Sheet TScsheet;
	static Workbook TScworkbook;
	static String execpath;
	static int TScrowcount = 0;
	static int loopnum = 1;
	static int windowFound = 0;
	static String TScname;
	static String ActionVal;
	String BrowserType;// Assign with either FF or IE
	static int DTcolumncount = 0;
	static WebElement elem = null;
	Alert dialogSwitch = null; // Updated on 20-Dec-2013
	// Updated on on 16.12.2013
	private static Map<String, String> map = new HashMap<String, String>();
	// private static Map<String, Float> mapint = new HashMap<String, Float>();
	public static int devFlag = 0;
	public static String varname;
	public static int initscreen = 1;
	public static int initscreenFlag = 1;
	String BrowserName;
	public static String objectType;
	public static int objFoundFlag = 0;
	public static String capString = "";
	static int locatorError = 0;
	static int ScreenshotTypeFlag = 0;
	int varUpdateReport;
	int today = 0;
	String selectedDateClass = "";
	int lastSelecteddateJQ = 0;
	int lastSelectedmonthJQ = 0;
	int lastSelectedyearJQ = 0;
	static int conditionline = 0;
	static int reporttype = 0;
	static String ReusableComponents;
	static String ExpectedReportData = "";
	static String ActualReportData = "";
	static String OptionalString = "";
	static String OptionalString2 = "";
	static String[] getCellArray;

	/*
	 * This function reads the selenium utility file and identifies where Object
	 * Repository, Test Suite & Test Scripts are located
	 */

	@Test
	public void ReadUtilFile() {
		Workbook w1 = null;
		try {
			if (!setObj.utilityFilePath.toString().endsWith("xls")) {
				System.out
						.println("Utility file defined in Setting.Java is not a correct one. Give the correct (Xls) Format of Utility File");
				return;
			} else {
				w1 = Workbook.getWorkbook(new File(setObj.utilityFilePath));
			}

		} catch (BiffException e) {
			System.out
					.println("Selenium Utility file is not accessible. Check your Utility file in the given path");
			e.printStackTrace();
			return;
		} catch (IOException e) {
			System.out
					.println("Utility file defined in Setting.Java is not a correct one. Give the correct Utility file path in Setting.java");
			e.printStackTrace();
			System.out.println(e);
			return;
		} catch (NullPointerException e) {
			System.out
					.println("The format of Utility file is not a corect one. Give the correct (Xls) Format of Utility File.");
			return;
		}
		try {
			Sheet sheet = w1.getSheet(0);
			TestSuite = sheet.getCell(1, 1).getContents();
			TestScript = sheet.getCell(1, 2).getContents();
			ObjectRepository = sheet.getCell(1, 3).getContents();
			ReportsPath = sheet.getCell(1, 5).getContents();
			ScreenReport = sheet.getCell(1, 6).getContents();
			BrowserName = sheet.getCell(1, 7).getContents();
			tempTestReportPath = sheet.getCell(1, 9).getContents();
			ieDriverPath = sheet.getCell(1, 8).getContents();
			try {
				execpath = sheet.getCell(1, 10).getContents();
				ReusableComponents = sheet.getCell(1, 11).getContents();

			} catch (ArrayIndexOutOfBoundsException exp10) {
				System.out
						.println("WARNING: If you want to use file upload/ downoad automation feature, use the new format of Utility(Configuration) Excel Which contains the path for the FileManager EXE. ");

			}
			if (!TestSuite.endsWith(".xls")) {
				if (TestSuite.toString() == "") {
					System.out
							.println("No proper TestSuite is defined in Selenium Utility file. Give the correct TestSuite name.");
				} else {
					System.out
							.println("Invalid File Format of TestSuite in Selenium Utility file: "
									+ TestSuite + ". Give only xls file format");
				}

			}

			if (!ObjectRepository.endsWith(".xls")) {
				if (ObjectRepository.toString() == "") {
					System.out
							.println("No proper Object repository is defined in Selenium Utility excel. Give the correct Object repository. ");
				} else {
					System.out
							.println("Invalid File Format of Object Repository : "
									+ ObjectRepository
									+ " in Selenium Utility file. Give only xls file format");
				}

			}
			if (BrowserName.toString() == "") {
				System.out
						.println("The browser type in the Selenium Utility file is not proper. Give a valid browser type");
			}

			for (int z = 0; z < 1; z++) {
				loopstart[z] = 0;
				loopend[z] = 0;
				loopcnt[z] = 0;
				dtrownumloop[z] = 1;
				loopcount[z] = 0;
			}

		} catch (ArrayIndexOutOfBoundsException ex) {
			System.out
					.println("The format of Utility file is not a corect one. Give the correct (Xls) Format of Utility File.");
			System.out.println(ex);
			return;
		} catch (Exception e) {
			System.out.println(e);
		}
		try {

			setBrowser(BrowserName);

		} catch (Exception e) {
			System.out
					.println("Could not launch the Browser. Check the browser settings");
			e.printStackTrace();
			return;
		}
		try {
			FindExecTestscript(TestSuite, TestScript, ObjectRepository);

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public void setBrowser(String BrowserType) {
		try {
			//Closing IEDriverServer and Internet Explorer before starting Driver  //Added by Santosh 23/11/2017
			Runtime.getRuntime().exec("taskkill /IM IEDriverServer.exe /F");
			Runtime.getRuntime().exec("taskkill /IM iexplore.exe /F");		
			
			
			System.out.println(BrowserType);

			switch (BrowserType.toUpperCase()) {
			case "IE":
				System.setProperty("webdriver.ie.driver", ieDriverPath);
				// Changed on 16/12/2013
				DesiredCapabilities capability = DesiredCapabilities
						.internetExplorer();
				capability
						.setCapability(
								InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,
								true);
				capability.setCapability("useLegacyInternalServer", true);
				capability.setCapability("nativeEvents", false);
				D8 = new InternetExplorerDriver(capability);
				D8.getWindowHandle();				
				D8.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				D8.manage().window().maximize();
				Capabilities cap = ((RemoteWebDriver) D8).getCapabilities();
				browserver = cap.getVersion();

				break;
			case "FF":
				D8 = new FirefoxDriver();
				D8.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				D8.manage().window().maximize();
				break;

			default:
				System.out.println("Error : Invalid browser type");
			}
		} catch (Exception e1) {
			System.out.println("Check the Test Suite file contents");
			System.out.println(e1);
		}
	}

	public void FindExecTestscript(String TestSuite, String TestScript,
			String ObjectRepository) throws Exception {
		try {
			// int TSrowcount = 0;
			FileInputStream fs = null;
			WorkbookSettings ws = null;
			try {
				fs = new FileInputStream(new File(TestSuite));
			} catch (Exception e) {
				System.out
						.println("No proper TestSuite is defined in Selenium Utility file. Give the correct TestSuite name");
				System.out.println(e);
				return;
			}

			ws = new WorkbookSettings();
			ws.setLocale(new Locale("en", "EN"));
			Workbook TSworkbook = Workbook.getWorkbook(fs, ws);
			Sheet TSsheet = TSworkbook.getSheet(0);
			// TSrowcount = TSsheet.getRows();
			cur_dt = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			String strTimeStamp = dateFormat.format(cur_dt);
			String[] dateArray = strTimeStamp.split("-");
			today = Integer.parseInt(dateArray[2]);
			String rp = ReportsPath;
			if (rp == "") {
				// if results path is passed as null, by
				// default 'C:\' drive is used

				// Updated on on 12/16/2013
				rp = tempTestReportPath;
			}

			if (!rp.endsWith("\\")) { // checks whether the path ends with
										// '/'
				rp = rp + "\\";
			}
			strResultPath = rp + "Log";
			String htmlname1 = rp + "Log" + "/Test_Suite_" + strTimeStamp
					+ ".html";
			File f = new File(strResultPath);
			if (strResultPath != f.getAbsolutePath().toString()) {
				System.out
						.println("Report will be printed in the following path since THE REPORT PATH WAS NOT GIVEN / THE GIVEN PATH IS INCORRECT.");
				System.out.println(f.getAbsolutePath().toString());
			}
			f.mkdirs();
			bw1 = new BufferedWriter(new FileWriter(htmlname1));
			bw1.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
			bw1.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
			bw1.write("<TR><TD COLSPAN=6 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Testcase Name</B></FONT></TD><TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Status</B></FONT></TD></TR>");

			for (int i = 0; i < TSsheet.getRows(); i++) {
				String TSvalidate = "r";
				if (((TSsheet.getCell(0, i).getContents())
						.equalsIgnoreCase(TSvalidate) == true))

				{
					scriptfound = 1;
					String ScriptName = TSsheet.getCell(1, i).getContents();
					if (!ScriptName.endsWith(".xls")) {

						System.out
								.println("Invalid File Format of TestScript: "
										+ ScriptName
										+ "in Test Suite.Give only xls format");

					}
					try {
						ExecKeywordScript(ScriptName, TestScript,
								ObjectRepository);
					} catch (FileNotFoundException e) {
						System.out
								.println("Invalid File:"
										+ ScriptName
										+ " is not available.Give the correct TestScript");
					}
					if (exeStatus.equalsIgnoreCase("Failed")) {
						bw1.write("<TR><TD COLSPAN=6 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>"
								+ TCNm[0]
								+ "</B></FONT></TD><TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=RED SIZE=2><B>"
								+ exeStatus + "</B></FONT></TD></TR>");
					} else {

						bw1.write("<TR><TD COLSPAN=6 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>"
								+ TCNm[0]
								+ "</B></FONT></TD><TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>"
								+ exeStatus + "</B></FONT></TD></TR>");
					}

				}

			}
			if (!(scriptfound == 1)) {
				System.out
						.println("Invalid File Format of Testsuite/ No Test Script to execute. Give the correct Testsuite/TestScript. ");
			}

			bw1.close();
		} catch (Exception e) {
			System.out.println(e);
			bw.close();
			bw1.close();
		}
	}

	public void ExecKeywordScript(String scriptName, String TestScript,
			String ObjectRepository) throws Exception

	{
		// Report header
		cur_dt = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		String strTimeStamp = dateFormat.format(cur_dt);

		if (ReportsPath == "") {

			ReportsPath = tempTestReportPath;
		}

		if (!ReportsPath.endsWith("\\"))

		{

			ReportsPath = ReportsPath + "\\";
		}
		TCNm = scriptName.split("\\.");
		strResultPath = ReportsPath + "Log" + "/" + TCNm[0] + "/";
		String htmlname = ReportsPath + "Log" + "/" + TCNm[0] + "/"
				+ strTimeStamp + ".html";
		File f = new File(strResultPath);
		f.mkdirs();
		bw = new BufferedWriter(new FileWriter(htmlname));
		bw.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TR><TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Test Case Name:</B></FONT></TD><TD COLSPAN=6 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>"
				+ TCNm[0] + "</B></FONT></TD></TR>");
		bw.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TR COLS=6><TD BGCOLOR=#FFCC99 WIDTH=3%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Row</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=15%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Keyword</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Object</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Action</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Execution Time</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Status</B></FONT></TD></TR>");

		exeStatus = "Pass";
		String scriptPath = TestScript + scriptName;
		TScname = scriptName;
		FileInputStream fs1 = null;
		WorkbookSettings ws1 = null;
		fs1 = new FileInputStream(new File(scriptPath));
		ws1 = new WorkbookSettings();
		ws1.setLocale(new Locale("en", "EN"));
		Workbook TScworkbook = Workbook.getWorkbook(fs1, ws1);
		TScsheet = TScworkbook.getSheet(0);
		TScrowcount = TScsheet.getRows();

		rowcnt = 0;

		for (j = 0; j < TScrowcount; j++)

		{
			rowcnt = rowcnt + 1;
			String TSvalidate = "r";
			if (((TScsheet.getCell(0, j).getContents())
					.equalsIgnoreCase(TSvalidate) == true))

			{
				Action = TScsheet.getCell(1, j).getContents();
				cCellData = TScsheet.getCell(2, j).getContents();
				dCellData = TScsheet.getCell(3, j).getContents();
				String ORPath = ObjectRepository;
				FileInputStream fs2 = null;
				WorkbookSettings ws2 = null;

				try {
					fs2 = new FileInputStream(new File(ORPath));
					ws2 = new WorkbookSettings();
					ws2.setLocale(new Locale("en", "EN"));
				} catch (Exception e) {
					System.out.println("Object Repository File not found");

				}

				try {
					Workbook ORworkbook = Workbook.getWorkbook(fs2, ws2);
					ORsheet = ORworkbook.getSheet(0);
					ORrowcount = ORsheet.getRows();
					ActionVal = Action.toLowerCase();
					iflag = 0;

				} catch (Exception e)

				{
					fail("Object Repository format is not correct.Check the File format");
					System.out.print(e);
				}
				bcellAction(scriptName);

			}

		}
		bw.close();

	}

	public static void screenshot(int loopn, int rown, String Sname)
			throws Exception {
		try {			
			if (devFlag == 0) {
				screenshotflag = screenshotflag + 1;
				DateFormat dateFormat = new SimpleDateFormat(
						"yyyyMMdd-HH-mm-ss");
				Date date = new Date();
				String filenamer = "";
				String strTime = dateFormat.format(date);
				Sname = Sname.substring(0, Sname.indexOf("."));

				ScreenReport = ScreenReport.toLowerCase();
				System.out.println("ScreenReport: "+ScreenReport);
				if (ScreenReport == "") {
					ScreenReport = ReportsPath;
					initscreenFlag = 2;
				}

				if (!ScreenReport.endsWith("\\")) {
					ScreenReport = ScreenReport + "\\";

				}
				if (!(ScreenReport.contains("screen"))) {
					ScreenReport = ScreenReport + "Screenshot\\";

				}
				initscreen = initscreen + 1;
				if (initscreen == 2 && initscreenFlag == 2) {
					System.out
							.println("Screenshots will be captured in the following path since the SCREENSHOTS PATH IS NOT GIVEN.");
					System.out.println(ScreenReport);

				}

				else {
					if (initscreen == 2) {
						File f1s = new File(ScreenReport);
						if (!ScreenReport.substring(0,
								ScreenReport.length() - 1).equalsIgnoreCase(

						f1s.getAbsolutePath().toString()))

						{
							System.out
									.println("Screenshots will be captured in the following path since the given SCREENSHOTS PATH IS NOT AVAILABLE.");
							System.out
									.println(f1s.getAbsolutePath().toString());

						}
						f1s.delete();
					}
				}
				if (loopflag == 0) {

					filenamer = ScreenReport + Sname + "\\" + Sname + "_"
							+ screenshotflag + "_rowno_" + (j + 1) + "_"
							+ strTime + ".png";

				} else {

					filenamer = ScreenReport + Sname + "\\" + Sname + "_"
							+ screenshotflag + "_loop_"
							+ (loopcnt[loopsize] + 1) + "_rowno_" + (j + 1)
							+ "_" + strTime + ".png";
				}

				if (ScreenshotTypeFlag == 0) {
					System.out.println("inside screenshot");
					File screenshot = D8.getScreenshotAs(OutputType.FILE);
					FileUtils.copyFile(screenshot, new File(filenamer));

				} else {

					BufferedImage image = new Robot()
							.createScreenCapture(new Rectangle(Toolkit
									.getDefaultToolkit().getScreenSize()));
					File outputfile = new File(filenamer);
					outputfile.mkdirs();
					ImageIO.write(image, "png", outputfile);
					Thread.sleep(3000);
				}
			}
			/*
			 * else { Update_Report("notForDevice1"); }
			 */
		}

		catch (Exception e) {
			System.out
					.println("Screenshot was not printed. Check the file path");
			Update_Report("screenshot", e);
		}
	}

	public static void Update_Report(String Res_type) throws IOException {
		try {
			String str_time;
			// String[] str_rep = new String[2];
			Date exec_time = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			str_time = dateFormat.format(exec_time);

			if (Res_type.startsWith("executed")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
						+ "Passed" + "</FONT></TD></TR>");
			} else if (Res_type.startsWith("failed")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");

			} else if (Res_type.startsWith("NoWindowFound")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Fail" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number: "
						+ (j + 1)
						+ ". Description: The Window not Found</div></th></FONT></TR>");
			} else if (Res_type.startsWith("loop")) {

				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
						+ Res_type + "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("callactionstart")) {

				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
						+ " Start of CallAction : '"
						+ globalCompName
						+ "' execution" + "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("callactionend")) {

				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
						+ " End of CallAction : '"
						+ globalCompName
						+ "' execution" + "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("callactionfnf")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Fail" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number: "
						+ (j + 1)
						+ ". Description: The Action(TestScript) name given is not available in the given path. Check the file path and action name.</div></th></FONT></TR>");
			} else if (Res_type.startsWith("callactionff")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Fail" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number: "
						+ (j + 1)
						+ ". Description: The Action(TestScript) format is not in supported. Give the file name with xls extension only.</div></th></FONT></TR>");
			}

			else if (Res_type.startsWith("missing")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
						+ (j + 1)
						+ ".Description: The Datatable column name not found</div></th></FONT></TR>");
			} else if (Res_type.startsWith("ObjectLocator")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
						+ (j + 1)
						+ ". Description: Object Locator is wrong or not supported. Supported Locators are Id,Name,Xpath& CSS</div></th></FONT></TR>");
			}

			else if (Res_type.startsWith("tooManyArguments")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED >"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "WARNING" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>WARNING @ Line NO:  "
						+ (j + 1)
						+ ". Pass only one argument if you are using data or environment variables. Otherwise only the first variable will be considered</div></th></FONT></TR>");
			} else if (Res_type.startsWith("NoBlankAvailable")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED >"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error @ Line NO:  "
						+ (j + 1)
						+ ". No Blank Value is available in the Element</div></th></FONT></TR>");
			} else if (Res_type.startsWith("objNotFound")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in test step number "
						+ (j + 1)
						+ ". Object with the given object name is not found in objectrepository</div></th></FONT></TR>");

			} else if (Res_type.startsWith("keyword")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ Action
						+ ": Keyword not supported by Open2Test</div></th></FONT></TR>");
			} else if (Res_type.startsWith("nodatatable")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Data table not available, Ensure whether you have imported the right Data Table.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("action")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ dCellData
						+ " : Action not supported by Open2Test</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("action1")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ dCellData
						+ " : The action is not supported for this type of object</div></th></FONT></TR>");
			} else if (Res_type.startsWith("objectmiss")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number "
						+ (j + 1)
						+ ".Object not found in the page</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("property")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Property not supported by Open2Test</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("property1")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Property not supported for this kind of object</div></th></FONT></TR>");
			} else if (Res_type.startsWith("condFailed")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (conditionline + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Condition Failed. Statements within the condition and End Condition will not be executed</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("invaliddate1")) {

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Invalid date has been given. Give the proper date format(It should be mm-dd-yyyy format) for a valid date in test script.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("invaliddate")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given Date cannot be set. Check the date is valid and in allowed range</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("prevmonth")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the previous month is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("nextmonth")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the next month is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("titlemonth")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the title month is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("titleyear")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the titleyear is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("titledefault")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the title year and month  is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("monthnotidentified")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Month in the calendar is not identified. English and Japanese character set are supported.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("invalidmonth")) {
				// System.out.println("fixed");

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given month is invalid. Correct the month and retry.</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("filePathNotFound")) {

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given file does not exist.Check the File path and File Name.</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("filePathNotFound1")) {

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given file path does not exist.Check the File path.</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("filePathNotFound2")) {

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "FilePath is not Given.Give a valid file path</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("calendaraction")) {

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "For SetDate Operation ObjectType Should be calendar and Object Name should start with cal_</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("userdefined")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Function is not executed. Error exists in  User Defined function. Correct the User Defined Function</div></th></FONT></TR>");
			}

			else if (Res_type.equalsIgnoreCase("getcelldata")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Invalid Syntax for getcelldata. Pass Exact number of arguments</div></th></FONT></TR>");
			}
		} catch (Exception e1) {
			System.out
					.println("Report will not be printed. Check the file path."
							+ e1);
		}

	}

	public static void Update_Report(String Res_type, Exception e)
			throws IOException {
		try {
			String str_time;
			Date exec_time = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			str_time = dateFormat.format(exec_time);

			exeStatus = "Failed";
			if (Res_type.startsWith("failed")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ e.toString().substring(
								e.toString().indexOf(":") + 1,
								e.toString().indexOf(".",
										e.toString().indexOf(":") + 1) + 1)
						+ "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("keyword")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number "
						+ (j + 1)
						+ ". Description: Keyword not supported by Open2Test</div></th></FONT></TR>");
			} else if (Res_type.startsWith("screenshot")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number "
						+ (j + 1)
						+ ". Screenshot not printed</div></th></FONT></TR>");
			}
		} catch (Exception e4) {
			System.out
					.println("Report will not be printed. Check the file path."
							+ e4);
		}

	}

	public static void Update_Report(String Res_type, String ObjectSet,
			String ObjectSetVal) throws IOException // //Updated on 16.12.2013
													// for Adding the values
													// instead variable names in
													// the report
	{
		try {
			String str_time;
			// String[] str_rep = new String[2];
			Date exec_time = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			str_time = dateFormat.format(exec_time);
			if (reporttype == 1) {
			} else {
			}
			if (Res_type.startsWith("executed")
					&& ObjectSet.equalsIgnoreCase("page")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " ; "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
						+ "Passed" + "</FONT></TD></TR>");
			} else if (Res_type.startsWith("executed")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
						+ "Passed" + "</FONT></TD></TR>");
			} else if (Res_type.startsWith("NoWindowFound")
					&& ObjectSet.equalsIgnoreCase("page")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " ; "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number: "
						+ (j + 1)
						+ ". Description: The Window not Found</div></th></FONT></TR>");
			}

			else if (Res_type.startsWith("failed")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
			} else if (Res_type.startsWith("loop")) {

				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
						+ Res_type + "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("missing")) {
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
						+ (j + 1)
						+ ". Description: The Datatable column name not found</div></th></FONT></TR>");
			} else if (Res_type.startsWith("ObjectLocator")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
						+ (j + 1)
						+ ". Description: Object Locator is wrong or not supported. Supported Locators are Id,Name,Xpath& CSS</div></th></FONT></TR>");
			} else if (Res_type.startsWith("Property")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in test step number "
						+ (j + 1)
						+ ". This property is not supported by Open2Test</div></th></FONT></TR>");
			}

		} catch (Exception e1) {
			System.out
					.println("Report will not be printed. Check the file path."
							+ e1);
		}
	}

	private static void Func_StoreCheck() throws Exception {
		// TODO Auto-generated method stub

		try {
			String actval = null;
			String expval = null;
			Boolean boolval = null;
			String varname;
			String[] cCellDataValCh = cCellData.split(";");
			String objectType = cCellDataValCh[0];
			String ObjectValCh = "";
			try {
				// Inserted on 16/12/2012. For holding Object Type
				ObjectValCh = cCellDataValCh[1];

				// Inserted on 16/12/2012. For Holding Object Name
			} catch (ArrayIndexOutOfBoundsException e) {
				ObjectValCh = "";
			}
			String[] dCellDataValCh = dCellData.split(":");

			String ObjectSetCh = dCellDataValCh[0];

			String ObjectSetValCh = "";
			int DTcolumncountCh = 0;
			if (objectType.equalsIgnoreCase("textelement")) {
				ObjectSetValCh = dCellData.replaceFirst("text:", "");
			} else {

				if (dCellDataValCh.length == 2) {
					ObjectSetValCh = dCellDataValCh[1];
				}
			}
			if (dCellDataValCh.length == 2) {
				if (ObjectSetValCh.substring(0, 1).equalsIgnoreCase("#")) {
					if (map.get(ObjectSetValCh.substring(1,
							(ObjectSetValCh.length()))) != null) {
						ObjectSetValCh = map.get(ObjectSetValCh.substring(1,
								(ObjectSetValCh.length())));

					} else {
						ObjectSetValCh = "";
					}
				}
			}
			if (objectType.equalsIgnoreCase("page")
					|| objectType.equalsIgnoreCase("dialog")) {

				objFoundFlag = 1;

			} else

			{
				for (int k = 0; k < ORrowcount; k++) {
					// String ORName = ORsheet.getCell(1, k).getContents();

					if (((ORsheet.getCell(1, k).getContents())
							.equalsIgnoreCase(ObjectValCh) == true)) {
						String[] ORcellData = new String[3];
						ORcellData = (ORsheet.getCell(4, k).getContents())
								.split("=", 2);

						ORvalname = ORcellData[0];
						ORvalue = ORcellData[1];
						k = ORrowcount + 1;
						objFoundFlag = 1;
					}

				}
			}

			if (objFoundFlag == 1) {
				objFoundFlag = 0;
				if (dCellDataValCh.length == 2) {
					if (ObjectSetValCh.contains("dt_")) {
						iflag = 0;
						String ObjectSetValtableheader[] = ObjectSetValCh
								.split("_");
						int column = 0;
						String Searchtext = ObjectSetValtableheader[1];
						try {
							DTcolumncountCh = DTsheet.getColumns();
						} catch (NullPointerException e) {
							return;
						}

						for (column = 0; column < DTcolumncountCh; column++) {
							if (Searchtext.equalsIgnoreCase(DTsheet.getCell(
									column, 0).getContents()) == true) {
								ObjectSetValCh = DTsheet.getCell(column,
										dtrownum).getContents();
								iflag = 1;
								if (ObjectSetValCh.length() == 0) {
									ObjectSetValCh = "";
								}
							}

						}
						if (iflag == 0) {
							ObjectSetValCh = "nodatarow";
						}
					}

					if (ObjectSetValCh.contains("dt_")) {
						String ObjectSetValtableheader[] = ObjectSetValCh
								.split("_");
						int column = 0;
						String Searchtext = ObjectSetValtableheader[1];
						DTcolumncountCh = DTsheet.getColumns();

						for (column = 0; column < DTcolumncountCh; column++) {
							if (Searchtext.equalsIgnoreCase(DTsheet.getCell(
									column, 0).getContents()) == true) {
								ObjectSetValCh = DTsheet.getCell(column,
										dtrownum).getContents();

								iflag = 1;
								if (ObjectSetValCh.length() == 0) {
									ObjectSetValCh = "";
								}
							}

						}
						if (iflag == 0) {
							ORvalname = "exit";
						}
					}
				}
				switch (ObjectSetCh.toLowerCase()) {
				case "enabled":
					if (objectType.equalsIgnoreCase("textbox")
							|| objectType.equalsIgnoreCase("combobox")
							|| objectType.equalsIgnoreCase("listbox")
							|| objectType.equalsIgnoreCase("radiobutton")
							|| objectType.equalsIgnoreCase("button")
							|| objectType.equalsIgnoreCase("checkbox")
							|| objectType.equalsIgnoreCase("textarea")
							|| objectType.equalsIgnoreCase("image")
							|| objectType.equalsIgnoreCase("table")
							|| objectType.equalsIgnoreCase("link")
							|| objectType.equalsIgnoreCase("element")) {
						Func_FindObj(ORvalname, ORvalue);
						boolval = elem.isEnabled();
						actval = boolval.toString();
					} else {
						Update_Report("property1");
					}
					break;
				case "text":
					if (objectType.equalsIgnoreCase("textbox")
							|| objectType.equalsIgnoreCase("combobox")
							|| objectType.equalsIgnoreCase("textarea")
							|| objectType.equalsIgnoreCase("radiobutton")
							|| objectType.equalsIgnoreCase("checkbox")
							|| objectType.equalsIgnoreCase("button")) {

						Func_FindObj(ORvalname, ORvalue);
						actval = elem.getAttribute("value");

					} else if (objectType.equalsIgnoreCase("textelement")
							|| objectType.equalsIgnoreCase("element")
							) {

						Func_FindObj(ORvalname, ORvalue);
						actval = elem.getText();
					} else {

						Update_Report("property1");
					}
					break;
				case "value":
					if (objectType.equalsIgnoreCase("combobox")) {
						Func_FindObj(ORvalname, ORvalue);
						actval = new Select(elem).getFirstSelectedOption()
								.getText().toString();
					} else {
						Update_Report("property1");
					}
					break;
				case "visible":
					if (objectType.equalsIgnoreCase("textbox")
							|| objectType.equalsIgnoreCase("combobox")
							|| objectType.equalsIgnoreCase("listbox")
							|| objectType.equalsIgnoreCase("radiobutton")
							|| objectType.equalsIgnoreCase("button")
							|| objectType.equalsIgnoreCase("checkbox")
							|| objectType.equalsIgnoreCase("textarea")
							|| objectType.equalsIgnoreCase("image")
							|| objectType.equalsIgnoreCase("table")
							|| objectType.equalsIgnoreCase("link")
							|| objectType.equalsIgnoreCase("element")) {
						Func_FindObj(ORvalname, ORvalue);
						boolval = elem.isDisplayed();
						actval = boolval.toString();
					} else {
						Update_Report("property1");
					}

					break;
				case "checked":
					if ((objectType.equalsIgnoreCase("radiobutton")
							|| objectType.equalsIgnoreCase("checkbox") || objectType
								.equalsIgnoreCase("element"))) {
						Func_FindObj(ORvalname, ORvalue);
						boolval = elem.isSelected();
						actval = boolval.toString();

					} else {
						Update_Report("property1");
					}
					break;
				case "linktext":
					if (objectType.equalsIgnoreCase("link")) {
						Func_FindObj(ORvalname, ORvalue);						
						actval = elem.getText();						
					} else {
						Update_Report("Property1");
					}
					break;
				case "pagetitle":
					if (ObjectValCh != "") {
						actval = ObjectValCh;
					} else {
						actval = D8.getTitle();
					}
					break;

				case "exist":
					boolval = false;
					actval = boolval.toString();

					if ((objectType.equalsIgnoreCase("page")) == true

							&& (D8.getTitle().toString()
									.equalsIgnoreCase(ObjectValCh)) == true) {

						actval = "true";
					} else {
						if (objectType.equalsIgnoreCase("page")) {

							String currentWindowHandle = D8.getWindowHandle();
							int windowFound = 0;
							WebDriver newWindow = null;
							Set<String> al = new HashSet<String>();
							al = D8.getWindowHandles();
							Iterator<String> windowIterator = al.iterator();

							if (D8.getTitle().toString()
									.equalsIgnoreCase(ObjectValCh) != true) {
								while (windowIterator.hasNext()) {
									String windowHandle = windowIterator.next();
									newWindow = D8.switchTo().window(
											windowHandle);
									if (newWindow.getTitle().toString()
											.equalsIgnoreCase(ObjectValCh) == true) {
										boolval = true;
										actval = boolval.toString();
										windowFound = 1;
										break;
									}

								}
								if (windowFound != 1) {
									boolval = false;

									actval = boolval.toString();
								}
								D8.switchTo().window(currentWindowHandle);
							}

						} else {

							if (objectType.equalsIgnoreCase("dialog") == true) {
								try {

									Alert dialogExist = null;
									dialogExist = D8.switchTo().alert();
									if (dialogExist.toString() != null) {
										boolval = true;
										actval = boolval.toString();
									} else {
										boolval = false;
										actval = boolval.toString();
									}

								} catch (NoAlertPresentException e) {

									boolval = false;
									actval = boolval.toString();

								}

							}
						}

					}

					break;
				case "rowcount":
					Func_FindObj(ORvalname, ORvalue);
					List<WebElement> rows = elem.findElements(By.tagName("tr"));
					Integer rowCount = rows.size();
					if (rowCount == 0) {
						// WebElement rows1=elem.findElement(By.tagName("tr"));
						rowCount = 1;
					}
					actval = rowCount.toString();
					break;

				case "columncount":
					WebElement headerRow = null;
					Func_FindObj(ORvalname, ORvalue);
					List<WebElement> rows1 = elem
							.findElements(By.tagName("tr"));
					try {
						headerRow = rows1.get(rows1.size()-2);
					} catch (Exception Ed) {
						try
						{
							headerRow = rows1.get(1);
						}
						catch(Exception Ed1)
						{
							headerRow = rows1.get(0);
						}
					}

					List<WebElement> columns = headerRow.findElements(By
							.tagName("th"));
					Integer colCount = columns.size();
					if (colCount == 0) {
						List<WebElement> columns0 = headerRow.findElements(By
								.tagName("td"));
						colCount = columns0.size();
						if (colCount == 0) {
							WebElement columns1 = headerRow.findElement(By
									.tagName("th"));
							if (columns1 != null) {
								colCount = 1;
							} else {
								WebElement columns2 = headerRow.findElement(By
										.tagName("td"));
								if (columns2 != null) {
									colCount = 1;
								}
							}
						}

					}
					actval = colCount.toString();
					break;
				case "getcelldata":

					Func_FindObj(ORvalname, ORvalue);
					try {
						List<WebElement> cellRows = elem.findElements(By
								.tagName("tr"));
						getCellArray = dCellData.split(":");
						String cellData = "";
						int rowNumber = Integer.parseInt(getCellArray[2]);
						int colNumber = Integer.parseInt(getCellArray[3]);
						System.out.println("rownum - " + rowNumber
								+ "  col num - " + colNumber);
						WebElement reqrow = cellRows.get(rowNumber - 1);
						List<WebElement> reqcolmns = reqrow.findElements(By
								.tagName("td"));
						WebElement reqcellData = reqcolmns.get(colNumber - 1);
						cellData = reqcellData.getText();
						actval = cellData.toString();
						ObjectSetValCh = getCellArray[1];
					} catch (Exception e) {
						Update_Report("getcelldata");
					}
					break;

				default:
					actval = "Invalid syntax";
					Update_Report("property");
					break;
				}

				if ((ActionVal).equalsIgnoreCase("check")) {
					expval = ObjectSetValCh;
					if (expval.equalsIgnoreCase("On"))
						expval = "True";
					else if (expval.equalsIgnoreCase("Off"))
						expval = "False";

					if (expval.equalsIgnoreCase(actval)) {
						System.out
								.println("Actual value matches with expected value. Actual and expected value is "
										+ actval);
						if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {

							Update_Report("executed");
						} else {
							Update_Report("executed", ObjectSetCh,
									ObjectSetValCh);
						}
						if (capturecheckvalue == true) {
							screenshot(loopnum, TScrowcount, TScname);
						}
					}else if(actval.trim().contains(expval.trim())) {	
						System.out
						.println("Actual value contains in the expected value. Expected Value is: "+ expval.trim() +", Actual value is: "
								+ actval);
							Update_Report("executed", ObjectSetCh,
									ObjectSetValCh);
					} else {
						System.out
								.println("Actual value doesn't match with expected value. Actual value is "
										+ actval
										+ " expected value is "
										+ expval);

						if (ORvalname == "exit") {
							Update_Report("missing", ObjectSetCh,
									ObjectSetValCh);

						} else {
							if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
								Update_Report("failed");
							} else {
								Update_Report("failed", ObjectSetCh,
										ObjectSetValCh);
							}

						}
						if (capturecheckvalue == true) {
							screenshot(loopnum, TScrowcount, TScname);
						}
					}
				} else if ((ActionVal).equalsIgnoreCase("storevalue"))

				{
					varname = ObjectSetValCh;

					if (actval.equalsIgnoreCase("Invalid syntax")) {
						Update_Report("missing", ObjectSetCh, ObjectSetValCh);

					} else {
						if (map.containsKey(varname)) {
							map.remove(varname);
							map.put(varname, actval);

							if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
								Update_Report("executed");
							} else {
								Update_Report("executed", ObjectSetCh,
										ObjectSetValCh);
							}
							System.out
									.println("Overwriting the value of the variable "
											+ varname
											+ " to store the value as mentioned in the test case row number"
											+ rowcnt);
						} else {
							map.put(varname, actval);
							if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
								Update_Report("executed");
							} else {
								Update_Report("executed", ObjectSetCh,
										ObjectSetValCh);
							}
							System.out
									.println("Overwriting the value of the variable "
											+ varname
											+ " to store the value as mentioned in the test case row number"
											+ rowcnt);
							if (ObjectSetValCh == "nodatarow") {
								Update_Report("missing");
							} else {

							}
						}
					}
					if (capturestorevalue == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
				}

			} else {
				Update_Report("objNotFound");
			}

		} catch (Exception e) {
			if (locatorError == 1) {

			} else {

				Update_Report("failed", e);
			}
		}
	}

	@After
	public void close() throws Exception

	{
		try {
			System.out.println("Test end.");

			D8.quit();
		} catch (UnhandledAlertException e) {
			System.out.println(e);
			System.out
					.println("Because of specification of SeleniumWebDriver, downloading may be failed.");
			System.out
					.println("Please confirm the report file and screenshot about test result.");
		}
	}

	private static void Func_FindObj(String strObjtype, String strObjpath)
			throws Exception

	{

		if (strObjpath.contains("@@@")) {
			String eCellDataVal = TScsheet.getCell(4, j).getContents();
			if (map.containsKey(eCellDataVal)) {
				eCellDataVal = map.get(eCellDataVal);
			}
			strObjpath = strObjpath.replace("@@@", eCellDataVal);

		}

		try {
			if (strObjtype.length() > 0 && strObjpath.length() > 0) {

				if (strObjtype.equalsIgnoreCase("id")) {
					elem = D8.findElement(By.id(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("name")) {
					elem = D8.findElement(By.name(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("xpath")) {
					elem = D8.findElement(By.xpath(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("link")) {
					elem = D8.findElement(By.linkText(strObjpath.toString()));
				} else if (strObjtype.equalsIgnoreCase("css")) {
					elem = D8.findElement(By.cssSelector(strObjpath));
				}

			}
		} catch (Exception ex2) {
			Update_Report("objectmiss");
			System.out.println(ex2.toString());
			elem = null;
		}

	}

	public static int ifContidionSkipper(String strifConditionStatus)
			throws Exception {
		try {
			String strKeyword;
			int intIfEndConditionCount, intIfConditionCount;
			// String strKeyWord;
			intIfConditionCount = 1;
			intIfEndConditionCount = 0;

			if (strifConditionStatus.equalsIgnoreCase("false")) {
				// intLogicStartRow = j;

				do {
					j = j + 1;
					strKeyword = TScsheet.getCell(1, j).getContents();

					if (strKeyword.equalsIgnoreCase("Condition")) {
						intIfConditionCount = intIfConditionCount + 1;
					}

					if (strKeyword.equalsIgnoreCase("Endcondition")) {
						intIfEndConditionCount = intIfEndConditionCount + 1;

						if (intIfConditionCount == intIfEndConditionCount) {
							j = j + 1;
							break;
						}
					}

				} while (true);
			}
		} catch (Exception e) {
			System.out.println(e);
		}
		return j;
	}

	public String Func_IfCondition(String strConditionArgs) throws Exception {
		int iFlag = 1;
		String str[] = strConditionArgs.split(";");
		String value1 = str[0];
		String value2 = str[2];
		String strOperation = str[1];
		strOperation = strOperation.toLowerCase().trim();
		switch (strOperation.toLowerCase()) {
		case "equals":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				System.out
						.println("Variable used in condition statement has value: "
								+ value1);
				if (value1.trim().equalsIgnoreCase(value2.trim())) {

					iFlag = 0;
				}
			} else if (value1.trim().equalsIgnoreCase(value2.trim())) {
				iFlag = 0;
			}
			break;

		case "notequals":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				System.out
						.println("Variable used in condition statement has values: "
								+ value1);
				if (!value1.trim().equalsIgnoreCase(value2.trim())) {

					iFlag = 0;
				}

			} else if (!value1.trim().equals(value2.trim())) {
				iFlag = 0;
			}
			break;
		case "greaterthan":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				if (isInteger(value1) && isInteger(value2)) {
					if (Integer.parseInt(value1) > Integer.parseInt(value2)) {
						iFlag = 0;
					}
				}
			}

			else if (isInteger(value1) && isInteger(value2)) {
				if (Integer.parseInt(value1) > Integer.parseInt(value2)) {
					iFlag = 0;
				}
			}

			else {
				System.out.println("Give Only Integers for Compare ");
			}
			break;
		case "lessthan":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				if (isInteger(value1) && isInteger(value2)) {
					if (Integer.parseInt(value1) < Integer.parseInt(value2)) {
						iFlag = 0;
					}
				}
			}

			else if (isInteger(value1) && isInteger(value2)) {
				if (Integer.parseInt(value1) < Integer.parseInt(value2)) {
					iFlag = 0;
				}
			}

			else {
				System.out.println("Give Only Integers for Compare ");
			}
			break;
		default:
			Update_Report("missing");
			break;
		}
		if (iFlag == 0) {
			return "true";

		} else {
			return "false";
		}
	}

	public void arrayresize() {
		if (loopstart.length <= loopsize) {
			loopstart = Arrays.copyOf(loopstart, loopstart.length + 1);
			loopend = Arrays.copyOf(loopend, loopend.length + 1);
			loopcnt = Arrays.copyOf(loopcnt, loopcnt.length + 1);
			dtrownumloop = Arrays.copyOf(dtrownumloop, dtrownumloop.length + 1);
			loopcount = Arrays.copyOf(loopcount, loopcount.length + 1);
		}
	}

	public void bcellAction(String scriptName) throws Exception {
		try {
			switch (ActionVal.toLowerCase()) {
			case "loop":
				startrow = j;
				dtrownum = 1;
				loopsize = loopsize + 1;
				if (loopsize >= 1) {
					arrayresize();

				}
				loopflag = 1;
				loopcount[loopsize] = Integer.parseInt(cCellData);
				loopstart[loopsize] = j;
				loopcnt[loopsize] = 0;
				dtrownumloop[loopsize] = dtrownum;
				Update_Report("loop : " + "Start of loop : " + (loopsize + 1));
				Update_Report("executed");
				break;
			case "endloop":
				loopcnt[loopsize] = loopcnt[loopsize] + 1;
				loopnum = loopnum + 1;
				if (loopcnt[loopsize] == loopcount[loopsize]) {
					Update_Report("loop" + " End of Loop : " + (loopsize + 1)
							+ " : Loop count : " + loopcnt[loopsize]);
					loopsize = loopsize - 1;
					if (loopsize >= 0)
						dtrownum = dtrownumloop[loopsize];
					else {
						dtrownum = 1;
						loopflag = 0;
					}
					Update_Report("executed");
				} else {
					j = loopstart[loopsize];
					dtrownum = dtrownum + 1;
					dtrownumloop[loopsize] = dtrownum;
					Update_Report("loop" + " End of Loop : " + (loopsize + 1)
							+ " : Loop count : " + loopcnt[loopsize]);
				}
				rowcnt = 1;
				break;
			case "launchapp":
				D8.get(cCellData);
				Thread.sleep(1000);
				boolean continueLinkDisplayed =D8.findElement(By.id("overridelink")).isDisplayed();
				if(continueLinkDisplayed) {
					D8.findElement(By.id("overridelink")).click();
				}
				Update_Report("executed");
				break;

			case "wait":
				Thread.sleep(Long.parseLong(cCellData));
				Update_Report("executed");
				break;

			case "condition":

				String strConditionStatus = Func_IfCondition(cCellData);
				if (strConditionStatus.equalsIgnoreCase("false")) {
					conditionline = j;
					j = ifContidionSkipper(strConditionStatus);
					j = j - 1;
				}
				if (strConditionStatus.equalsIgnoreCase("false")) {

					Update_Report("condFailed");
				} else {

					Update_Report("executed");
				}
				break;

			case "endcondition":
				Update_Report("executed");
				break;

			case "screencaptureoption":
				String[] sco = cCellData.split(";");

				for (int s = 0; s < sco.length; s++) {
					if (sco[s].equalsIgnoreCase("perform")) {
						captureperform = true;
					}
					if (sco[s].equalsIgnoreCase("storevalue")) {
						capturestorevalue = true;
					}
					if (sco[s].equalsIgnoreCase("check")) {
						capturecheckvalue = true;
					}

				}
				Update_Report("executed");
				break;

			case "importdata":
				try {
					String xcelpath = cCellData;
					FileInputStream fs3 = null;
					WorkbookSettings ws3 = null;
					fs3 = new FileInputStream(new File(xcelpath));
					ws3 = new WorkbookSettings();
					ws3.setLocale(new Locale("en", "EN"));
					Workbook DTworkbook = Workbook.getWorkbook(fs3, ws3);
					DTsheet = DTworkbook.getSheet(0);
					// int DTrowcount = DTsheet.getRows();
					// String DTName;
					Update_Report("executed");
				} catch (Exception e) {
					Update_Report("nodatatable");
				}
				break;

			case "screencapture":
				ScreenshotTypeFlag = 0;				
				screenshot(loopnum, TScrowcount, TScname);
				Update_Report("executed");
				break;
			case "context":// Updated on 16.12.2013 for page support
				cCellObjName = "";// Updated on 16.12.2013 for page support
				cCellObjType = "";// Updated on 16.12.2013 for page support
				dCellDataVal = null;
				int frameindex = 0;
				if (cCellData.contains(";")) // Updated on 16.12.2013 for page
												// and Dialog Support
				{

					if (cCellData.endsWith(";")) {
						cCellObjType = cCellData.substring(0,
								cCellData.length() - 1);
						cCellObjName = cCellData.substring(0,
								cCellData.length() - 1);

					} else {
						cCellDataVal = cCellData.split(";");
						cCellObjType = cCellDataVal[0];
						cCellObjName = cCellDataVal[1];
					}
				}

				else {
					cCellObjType = cCellData;
					if (cCellObjType.equalsIgnoreCase("frame")
							|| cCellObjType.equalsIgnoreCase("iframe")) {
						frameindex = 1;
					} else {
						cCellObjName = cCellData;
					}
				}
				if (cCellObjType.equalsIgnoreCase("frame")
						|| cCellObjType.equalsIgnoreCase("iframe")) {
					if (frameindex == 1) {
						D8.switchTo().parentFrame();
						Update_Report("executed");
						frameindex = 0;
					// 2015-06-17 add <---
					} else if("default".equals(cCellObjName)) {
						D8.switchTo().defaultContent();	
						Update_Report("executed");
					} else if(cCellObjName.matches("^[0-9]+") == true){
						int test = new Integer(cCellObjName);
						D8.switchTo().frame(test);
						Update_Report("executed");
					// 2015-06-17 add --->
					} else {
						for (int k = 0; k < ORrowcount; k++) {
							if (((ORsheet.getCell(1, k).getContents())
									.equalsIgnoreCase(cCellObjName) == true)) {
								String[] ORcellData = new String[3];
								ORcellData = (ORsheet.getCell(4, k)
										.getContents()).split("=");
								ORvalname = ORcellData[0];
								ORvalue = ORsheet.getCell(4, k).getContents()
										.substring(ORcellData[0].length() + 1);
								k = ORrowcount + 1;
							}

						}
						Func_FindObj(ORvalname, ORvalue);
						if (elem == null) {
							return;
						} else {
							D8.switchTo().frame(elem);
							Update_Report("executed");
						}
					}
				} else {

					String windowType = null;
					try {
						dCellData.toString();
						if (dCellData.contains("::")) {
							String tempDCellData = dCellData;
							dCellDataVal = dCellData.split(";");
							windowType = dCellDataVal[0];
							ObjectSetVal = dCellDataVal[1];
							dCellData = dCellData.substring(
									dCellData.indexOf("::") + 2,
									dCellData.length());
							if (dCellData.contains(";")) // Updated on
															// 16.12.2013 for
															// page
							// and Dialog Support
							{
								if (dCellData.endsWith(";")) {
									cCellObjType = dCellData.substring(0,
											dCellData.length() - 1);
									cCellObjName = dCellData.substring(0,
											dCellData.length() - 1);

								} else {
									dCellDataVal = dCellData.split(";");
									cCellObjType = dCellDataVal[0];
									cCellObjName = dCellDataVal[1];

								}
							}

							else {
								cCellObjType = dCellData;

								if (cCellObjType.equalsIgnoreCase("frame")
										|| cCellObjType
												.equalsIgnoreCase("iframe")) {
									frameindex = 1;
								} else {
									cCellObjName = dCellData;
								}
							}
							if (cCellObjType.equalsIgnoreCase("frame")
									|| cCellObjType.equalsIgnoreCase("iframe")) {
								if (frameindex == 1) {
									D8.switchTo().parentFrame();
									dCellData = tempDCellData;
									Update_Report("executed");
									frameindex = 0;

								} else {
									for (int k = 0; k < ORrowcount; k++) {
										if (((ORsheet.getCell(1, k)
												.getContents())

										.equalsIgnoreCase(cCellObjName) == true)) {
											String[] ORcellData = new String[3];
											ORcellData = (ORsheet.getCell(4, k)
													.getContents()).split("=");
											ORvalname = ORcellData[0];
											ORvalue = ORsheet
													.getCell(4, k)
													.getContents()
													.substring(
															ORcellData[0]
																	.length() + 1);
											k = ORrowcount + 1;

										}

									}
									Func_FindObj(ORvalname, ORvalue);
									if (elem == null) {
										return;
									} else {
										D8.switchTo().frame(elem);
										dCellData = tempDCellData;

										Update_Report("executed");

									}
								}
							}
						} else if (dCellData.contains(";")) {
							if (dCellData.endsWith(";")) {

								windowType = dCellData.toString();
								ObjectSetVal = dCellData.toString();

							} else {
								dCellDataVal = dCellData.split(";");
								windowType = dCellDataVal[0];
								ObjectSetVal = dCellDataVal[1]; // 2015-06-17
								if (ObjectSetVal.substring(0, 1)
										.equalsIgnoreCase("#")) {
									if (map.get(ObjectSetVal.substring(1,
											(ObjectSetVal.length()))) != null) {
										ObjectSetVal = map
												.get(ObjectSetVal
														.substring(
																1,
																(ObjectSetVal
																		.length())));

									} else {
										ObjectSetVal = "";
									}
								// data table check Add 2015-08-19 --->
								} else if (ObjectSetVal.contains("dt_")) {
									String ObjectSetValtableheader[] = ObjectSetVal.split("_");
									int DTcolumncountCh = DTsheet.getColumns();
									int column = 0;
									String Searchtext = ObjectSetValtableheader[1];

									for (column = 0; column < DTcolumncountCh; column++) {
										if (Searchtext.equalsIgnoreCase(DTsheet.getCell(column,
												0).getContents()) == true) {
											ObjectSetVal = DTsheet.getCell(column, dtrownum)
													.getContents();
											if (ObjectSetVal.length() == 0) {
												ObjectSetVal = "";
											}
											iflag = 1;
										}
									}	
								}
								// data table check Add 2015-08-19 <---
							}
						} else {							
							windowType = dCellData.toString();
							ObjectSetVal = dCellData.toString();
						}
						// 2015-06-23  <--- condition changed（cotext|dialog;）
						if ((!windowType.equalsIgnoreCase("dialog;")) && ((windowType.equalsIgnoreCase("page")
								|| windowType.equalsIgnoreCase("page;")) && !dCellData.contains("::")
								|| !windowType.equalsIgnoreCase("page;WindowRtn;"))) {
						// 2015-06-23 --->
							int windowNums = 0;
							int windowItr = 0;
							WebDriver newWindow = null;
							Set<String> al = new HashSet<String>();
							al = D8.getWindowHandles();
							windowNums = al.size(); // get the number of window
													// handles
							Iterator<String> windowIterator = al.iterator();
							if (windowNums == 1) { 				 
								// Switch the hundle, if number of available hundle is 1.
								String handle = windowIterator.next();
								D8.switchTo().window(handle);
								// Reset Iterator
   							    windowIterator = al.iterator();
							} else {
								// save the current window handle.
								parentWindowHandle = D8.getWindowHandle();
							}
							if (D8.getTitle().toString()

							.equalsIgnoreCase(ObjectSetVal) == true) {
								Update_Report("executed"); // 2015-08-19 Add
							} else {
								if (!((ObjectSetVal.equalsIgnoreCase("page") || (ObjectSetVal
										.equalsIgnoreCase("page;"))) || (ObjectSetVal
										.toString() == ""))) {
									if (D8.getTitle().toString()
											.equalsIgnoreCase(ObjectSetVal) == false) {
										for (windowItr = 0; windowItr < windowNums; windowItr++) {

											String windowHandle = windowIterator
													.next();
											newWindow = D8.switchTo().window(
													windowHandle);

											if (newWindow
													.getTitle()
													.toString()
													.equalsIgnoreCase(
															ObjectSetVal)) {

												windowFound = 1;
												Update_Report("executed",
														windowType,
														ObjectSetVal);
												break;

											}

										}
										if (windowFound != 1) {

											Update_Report("NoWindowFound",
													windowType, ObjectSetVal);
										}

									}
								} else {
									if (windowNums == 1) {
										Update_Report("executed");
									} else {
										int winItr1 = 0;
										String windowHandle = null;

										while (winItr1 != windowNums) {
											windowHandle = windowIterator
													.next();
											System.out.println(windowHandle);
											newWindow = D8.switchTo().window(
													windowHandle);
											if (parentWindowHandle
													.equalsIgnoreCase(windowHandle)) {

												if (winItr1 != windowNums - 1) {
													windowHandle = windowIterator
															.next();
													D8.switchTo().window(
															windowHandle);
													Update_Report("executed");
													windowFound = 1;
													break;
												} else

												{
													Iterator<String> windowIterator1 = al
															.iterator();
													windowHandle = windowIterator1
															.next();
													D8.switchTo().window(
															windowHandle);
													Update_Report("executed");
													windowFound = 1;
													break;

												}
											}

											winItr1++;
										}
										if (windowFound != 1) {
											Update_Report("NoWindowFound");
										}

									}
								}

							}
						// 2015-06-18 <---
						} else if(windowType.equalsIgnoreCase("page;WindowRtn;")) {
							D8.switchTo().window(parentWindowHandle);
							Update_Report("executed");
						// 2015-06-18 --->
						}

						if ((windowType.equalsIgnoreCase("dialog") || windowType
								.equalsIgnoreCase("dialog;"))) {
							D8.switchTo().alert();
							Update_Report("executed");

						}

					} catch (Exception e) {
						Update_Report("failed", e);

					}

				}
				break;
			case "check":
				ScreenshotTypeFlag = 0;
				Func_StoreCheck();
				break;
			case "storevalue":
				ScreenshotTypeFlag = 0;
				Func_StoreCheck();
				break;
			case "upload":

				if (cCellData.toString() == "") {

					Update_Report("filePathNotFound2");
					doUploadDownload("abortupload");
				} else {

					if (new File(cCellData.toString()).exists()) {

						// System.out.println(cCellData.toString());
						doUploadDownload("upload");
					} else {

						Update_Report("filePathNotFound");
						doUploadDownload("abortupload");
					}
				}
				break;
			case "download":

				cCellObjName = "";// Updated on 16.12.2013 for page support
				cCellObjType = "";// Updated on 16.12.2013 for page support
				dCellDataVal = null;

				if (cCellData.contains(";")) // Updated on 16.12.2013 for page
												// and Dialog Support
				{
					if (cCellData.endsWith(";")) {
						cCellObjType = cCellData;
						cCellObjName = cCellData;
					} else {
						cCellDataVal = cCellData.split(";");
						cCellObjType = cCellDataVal[0];
						cCellObjName = cCellDataVal[1];

					}
				}

				else {
					cCellObjType = cCellData;
					cCellObjName = cCellData;
				}
				ObjectSet = dCellData.toString();

				if (ObjectSetVal.contains("dt_"))
					DTcolumncount = DTsheet.getColumns();
				if (!(cCellObjType.equalsIgnoreCase("page")
						|| cCellObjType.equalsIgnoreCase("page;")
						|| cCellObjType.equalsIgnoreCase("dialog")
						|| cCellObjType.equalsIgnoreCase("dialog;")
						|| cCellObjType.equalsIgnoreCase("DownloadDialog") || cCellObjType
							.equalsIgnoreCase("uploadDialog"))) // //Updated on
																// 16.12.2013
																// for
																// supporting
																// page and
																// dialog
				{
					for (int k = 0; k < ORrowcount; k++) {
						// String ORName = ORsheet.getCell(1, k).getContents();

						if (((ORsheet.getCell(1, k).getContents())
								.equalsIgnoreCase(cCellObjName) == true)) {
							String[] ORcellData = new String[3];
							ORcellData = (ORsheet.getCell(4, k).getContents())
									.split("=");

							ORvalname = ORcellData[0];
							ORvalue = ORsheet.getCell(4, k).getContents()
									.substring(ORcellData[0].length() + 1);
							k = ORrowcount + 1;
						}
					}

				}
				try {
					// String[] cCellDataValCh = cCellData.split(";");
					// String objectType = cCellDataValCh[0];
					readAttributeforPerform();
					Func_FindObj(ORvalname, ORvalue);
					if (elem == null) {
						return;
					} else {
						if (ObjectSet == "") {

							Update_Report("filePathNotFound2");
						} else {
							try {

								// System.out.println(ObjectSet);

								ObjectSet1 = ObjectSet.split("\\\\");
							} catch (Exception e2) {

								System.out.println(e2);
							}
							for (int i = 0; i < ObjectSet1.length - 1; i++) {
								ObjectSet2 = ObjectSet2 + ObjectSet1[i] + "\\";

							}
							if (new File(ObjectSet2.toString()).exists()) {

								doUploadDownload("download");

							} else {

								Update_Report("filePathNotFound1");

							}
							ObjectSet2 = "";
							ObjectSet1 = null;
						}

					}

				} catch (Exception e) {

					System.out.println(e);
				}
				break;

			case "perform":
				ScreenshotTypeFlag = 0;

				cCellObjName = "";// Updated on 16.12.2013 for page support
				cCellObjType = "";// Updated on 16.12.2013 for page support
				dCellDataVal = null;

				if (cCellData.contains(";")) // Updated on 16.12.2013 for page
												// and Dialog Support
				{
					if (cCellData.endsWith(";")) {
						cCellObjType = cCellData;
						cCellObjName = cCellData;
					} else {
						cCellDataVal = cCellData.split(";");
						cCellObjType = cCellDataVal[0];
						cCellObjName = cCellDataVal[1];

					}
				}

				else {
					cCellObjType = cCellData;
					cCellObjName = cCellData;
				}

				dCellData.toString();

				if (dCellData.contains(":")) {
					if (!dCellData.endsWith(":")) {
						dCellDataVal = dCellData.split(":");
						ObjectSet = dCellDataVal[0].toLowerCase();
						ObjectSetVal = dCellDataVal[1];

						if (dCellDataVal.length > 2) {
							for (int dCellvalcount = 2; dCellvalcount < dCellDataVal.length; dCellvalcount++) {
								if (ObjectSetVal.startsWith("dt_")
										|| ObjectSetVal.startsWith("#"))

								{
									Update_Report("tooManyArguments");
									break;
								} else {
									ObjectSetVal = ObjectSetVal + ":"
											+ dCellDataVal[dCellvalcount];
								}
							}
						}
					} else {
						ObjectSet = dCellData.replace(":", "");
						ObjectSetVal = "";

					}
				} else {
					ObjectSet = dCellData.toString();
				}
				DTcolumncount = 0;
				if (ObjectSetVal.contains("dt_"))
					DTcolumncount = DTsheet.getColumns();
				if (!(cCellObjType.equalsIgnoreCase("page")
						|| cCellObjType.equalsIgnoreCase("page;")
						|| cCellObjType.equalsIgnoreCase("dialog")
						|| cCellObjType.equalsIgnoreCase("dialog;")
						|| cCellObjType.equalsIgnoreCase("DownloadDialog") || cCellObjType
							.equalsIgnoreCase("uploadDialog"))) // //Updated on
																// 16.12.2013
																// for
																// supporting
																// page and
																// dialog
				{
					for (int k = 0; k < ORrowcount; k++) {
						// String ORName = ORsheet.getCell(1, k).getContents();

						if (((ORsheet.getCell(1, k).getContents())
								.equalsIgnoreCase(cCellObjName) == true)) {
							String[] ORcellData = new String[3];
							ORcellData = (ORsheet.getCell(4, k).getContents())
									.split("=", 2);

							ORvalname = ORcellData[0];
							ORvalue = ORsheet.getCell(4, k).getContents()
									.substring(ORcellData[0].length() + 1);
							k = ORrowcount + 1;
						}
					}

				}

				dCellAction();
				break;
			case "callfunction":

				try {
					Func_invokeFunction(cCellData, dCellData);
				} catch (Exception e) {
					Update_Report("userdefined");
				}
				break;
			case "callaction":

				reporttype = 1;
				exeStatus = "Pass";
				String ComponentPath = ReusableComponents + cCellData;
				if (cCellData.contains("xls")) {
					String ComponentName = cCellData.split(".xls")[0];
					Sheet TestScriptSheet = TScsheet;
					FileInputStream ComponentFile1 = null;
					WorkbookSettings ComponentWS1 = null;

					try {
						ComponentFile1 = new FileInputStream(new File(
								ComponentPath));
						ComponentWS1 = new WorkbookSettings();
						ComponentWS1.setLocale(new Locale("en", "EN"));
						Workbook ComponentWorkBook = Workbook.getWorkbook(
								ComponentFile1, ComponentWS1);
						Sheet ComponentSheet = ComponentWorkBook.getSheet(0);
						int ComponentRowCount = 0;
						int introwcnt = 0;
						int introwcntStore = j;
						Update_Report("executed");
						j = j + 1;
						TScsheet = ComponentSheet;
						Stack<String> ComponentStack = new Stack<String>();
						globalCompName = ComponentName;
						ComponentStack.push(ComponentName);
						Update_Report("callactionstart");
						ComponentRowCount = ComponentSheet.getRows();
						introwcnt = 0;
						for (int jloop = 0; jloop < ComponentRowCount; jloop++) {
							introwcnt = introwcnt + 1;
							j = jloop;
							String CTValidate = "r";
							if (((ComponentSheet.getCell(0, jloop)
									.getContents())
									.equalsIgnoreCase(CTValidate) == true)) {
								Action = ComponentSheet.getCell(1, jloop)
										.getContents();
								cCellData = ComponentSheet.getCell(2, jloop)
										.getContents();
								dCellData = ComponentSheet.getCell(3, jloop)
										.getContents();
								String ORPath = ObjectRepository;
								FileInputStream ComponentFile2 = null;
								WorkbookSettings ComponentWS2 = null;
								try {
									ComponentFile2 = new FileInputStream(
											new File(ORPath));
									ComponentWS2 = new WorkbookSettings();
									ComponentWS2.setLocale(new Locale("en",
											"EN"));
								} catch (Exception e) {
									System.out.println("File not found");
								}
								try {
									Workbook ORworkbook = Workbook.getWorkbook(
											ComponentFile2, ComponentWS2);
									ORsheet = ORworkbook.getSheet(0);
									ORrowcount = ORsheet.getRows();
									ActionVal = Action.toLowerCase();
									iflag = 0;
								} catch (Exception e) {

									fail("Excel file of Open2Test is not correct.");
								}
								System.out.println(Action + "||" + cCellData
										+ "||" + dCellData);
								bcellAction(scriptName);

								jloop = j;
							}// End of Execution
						}// End of If that get all rows in Test Script
						globalCompName = ComponentStack.pop();
						Update_Report("callactionend");
						j = introwcntStore;
						reporttype = 0;
						TScsheet = TestScriptSheet;
					} catch (FileNotFoundException FNF1) {
						Update_Report("callactionfnf");
					}
				} else {
					Update_Report("fail");
					Update_Report("callactionff");
				}
				break;
			default:
				Update_Report("keyword");
				break;

			}
		}

		catch (Exception ex) {
			Update_Report("failed", ex);
			System.out.println(ex);
			System.out.println("------Error Information : Open2Test-------");
			System.out.println("Current Script:" + scriptName);
			System.out.println("Current ScriptPath:" + TestScript);
			System.out
					.println("Using ObjectRepositoryPath:" + ObjectRepository);
			System.out.println("Current Keyword:" + Action);
			System.out.println("Current ObjectDetails:" + cCellData);
			System.out.println("Current ObjectDetailsPath:" + ORvalue);
			System.out.println("Current Action:" + dCellData);
			System.out.println("------Error Information : Open2Test-------");
			fail("Cannot test normally by Open2Test.");
			// return;
		}
	}

	public void readAttributeforPerform() throws Exception {

		try {
			if (ObjectSetVal.length() > 0)

			{
				// //Function was updated on 16.12.2013 for supporting page,
				if (ObjectSetVal.substring(0, 1).equalsIgnoreCase("#")) {

					if (map.get(ObjectSetVal.substring(1,
							(ObjectSetVal.length()))) != null) {
						ObjectSetVal = map.get(ObjectSetVal.substring(1,
								(ObjectSetVal.length())));
					} else {
						ObjectSetVal = "";
					}

				} else if (ObjectSetVal.contains("dt_")) {
					iflag = 0;

					String ObjectSetValtableheader[] = ObjectSetVal.split("_");
					int column = 0;
					String Searchtext = ObjectSetValtableheader[1];

					for (column = 0; column < DTcolumncount; column++) {
						if (Searchtext.equalsIgnoreCase(DTsheet.getCell(column,
								0).getContents()) == true) {
							ObjectSetVal = DTsheet.getCell(column, dtrownum)
									.getContents();
							if (ObjectSetVal.length() == 0) {
								ObjectSetVal = "";
							}
							iflag = 1;

						}

					}

					if (iflag != 1) {
						ObjectSetVal = "nodatarow";
					}

					else {
						Update_Report("toomanyarguments");
					}
				}

			}
		} catch (Exception e) {

		}
	}

	public void dCellAction() throws Exception {

		if (!(cCellObjType.equalsIgnoreCase("page")
				|| cCellObjType.equalsIgnoreCase("page;")
				|| cCellObjType.equalsIgnoreCase("dialog")
				|| cCellObjType.equalsIgnoreCase("dialog;")
				|| cCellObjType.equalsIgnoreCase("downloaddialog")
				|| cCellObjType.equalsIgnoreCase("downloaddialog;")
				|| cCellObjType.equalsIgnoreCase("uploaddialog") || cCellObjType
					.equalsIgnoreCase("uploaddialog;")))// Updated on 16.12.2013
														// for page and dialog
														// support
		{
			try {
				String[] cCellDataValCh = cCellData.split(";");
				String objectType = cCellDataValCh[0];
				readAttributeforPerform();
				Func_FindObj(ORvalname, ORvalue);
				if (elem == null) {
					return;
				} else {
					switch (ObjectSet.toLowerCase()) {
					case "set":
						if (objectType.equalsIgnoreCase("textbox")
								|| objectType.equalsIgnoreCase("textarea")) {

							elem.clear();
							StringBuffer inputvalue = new StringBuffer();
							if (ObjectSetVal == "nodatarow") {
								Update_Report("missing");
							} else {

								inputvalue.append(ObjectSetVal);

								D8.executeScript(
										"arguments[0].value=arguments[0].value + '"
												+ inputvalue.toString() + "';",
										elem);
								Update_Report("executed", ObjectSet,
										ObjectSetVal);
							}

							if (captureperform == true) {
								screenshot(loopnum, TScrowcount, TScname);
							}
						} else {
							Update_Report("action1");
						}
						break;
					case "listselect":
						if (objectType.equalsIgnoreCase("listbox")) {
							int foundCount = 0;
							String[] listvalues = ObjectSetVal.split(":");
							List<WebElement> listboxitems = elem
									.findElements(By.tagName("option"));
							Select chooseoptn = new Select(elem);
							chooseoptn.deselectAll();
							if (ObjectSetVal == "nodatarow") {
								Update_Report("missing");
							} else {
								for (WebElement opt : listboxitems) {
									for (int i = 0; i < listvalues.length; i++) {
										if (opt.getText().equalsIgnoreCase(
												listvalues[i])) {
											chooseoptn.selectByVisibleText(opt
													.getText());
											foundCount = foundCount + 1;
										}

									}
								}
								if (foundCount == 0
										&& ObjectSetVal.contains(""))

								{
									Update_Report("NoBlankAvailable");

								} else {
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
								}

								if (captureperform == true) {
									screenshot(loopnum, TScrowcount, TScname);
								}
							}
						} else {
							Update_Report("action1");
						}

						break;
						
					case "listselectcontains": //Added by Santosh 11/23/2017
						if (objectType.equalsIgnoreCase("listbox")) {
							int foundCount = 0;							
							List<WebElement> listboxitems = elem
									.findElements(By.tagName("option"));
							Select chooseoptn = new Select(elem);							
							if (ObjectSetVal == "nodatarow") {
								Update_Report("missing");
							} else {
								for (WebElement opt : listboxitems) {									
										if (opt.getText().contains(
												ObjectSetVal)) {
											chooseoptn.selectByVisibleText(opt
													.getText());
											foundCount = foundCount + 1;
									}
								}
								if (foundCount == 0
										&& ObjectSetVal.contains(""))

								{
									Update_Report("NoBlankAvailable");

								} else {
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
								}

								if (captureperform == true) {
									screenshot(loopnum, TScrowcount, TScname);
								}
							}
						} else {
							Update_Report("action1");
						}

						break;
						
					case "select":
						if (objectType.equalsIgnoreCase("combobox")) {
							if (ObjectSetVal != "") {
								new Select(elem)
										.selectByVisibleText(ObjectSetVal);
								Update_Report("executed", ObjectSet,
										ObjectSetVal);
							} else if (ObjectSetVal == "nodatarow") {
								Update_Report("missing");
							}

							else {
								if (new Select(elem).getOptions().toString()
										.contains("") == true) {
									try {
										new Select(elem)
												.selectByVisibleText("");
										Update_Report("executed", ObjectSet,
												ObjectSetVal);
									} catch (Exception e) {
										Update_Report("NoBlankAvailable");
									}
								} else {
									Update_Report("NoBlankAvailable");
								}
							}
							if (captureperform == true) {
								screenshot(loopnum, TScrowcount, TScname);

							}
						} else {
							Update_Report("action1");
						}

						break;
						
					case "selectcontains":
						System.out.println("inside selectcontains");
						if (objectType.equalsIgnoreCase("combobox")) {
							if (ObjectSetVal != "") {
								List<WebElement> options = new Select(elem)
										.getOptions();
								for(WebElement opt: options) {
									if(opt.getText().trim().contains(ObjectSetVal)) {
									new Select(elem)
									.selectByVisibleText(opt.getText());
									}
								}
								
								Update_Report("executed", ObjectSet,
										ObjectSetVal);
							} else if (ObjectSetVal == "nodatarow") {
								Update_Report("missing");
							}

							else {
								if (new Select(elem).getOptions().toString()
										.contains("") == true) {
									try {
										new Select(elem)
												.selectByVisibleText("");
										Update_Report("executed", ObjectSet,
												ObjectSetVal);
									} catch (Exception e) {
										Update_Report("NoBlankAvailable");
									}
								} else {
									Update_Report("NoBlankAvailable");
								}
							}
							if (captureperform == true) {
								screenshot(loopnum, TScrowcount, TScname);
							}
						} else {
							Update_Report("action1");
						}

						break;
					
					case "selectbyvalue": //Added by Santosh 11/23/2017
						if (objectType.equalsIgnoreCase("combobox")) {
							if (ObjectSetVal != "") {
								Select options = new Select(elem);
								options.selectByValue(ObjectSetVal);			
								
								Update_Report("executed", ObjectSet,ObjectSetVal);
								
							} else if (ObjectSetVal == "nodatarow") {
								Update_Report("missing");
							}

						}else {
							if (new Select(elem).getOptions().toString()
									.contains("") == true) {
								try {
									new Select(elem)
											.selectByVisibleText("");
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
								} catch (Exception e) {
									Update_Report("NoBlankAvailable");
								}
							} else {
									Update_Report("NoBlankAvailable");
							}
						}
						if (captureperform == true) {
							screenshot(loopnum, TScrowcount, TScname);
						}else {
						Update_Report("action1");
						}
						break;
	

					case "check":

						if (ObjectSetVal == "nodatarow") {
							Update_Report("missing");
						} else {

							if (objectType.equalsIgnoreCase("checkbox")) {
								if (elem.isSelected()
										&& ObjectSetVal.equalsIgnoreCase("On")) {
									Update_Report("executed", ObjectSet,
											ObjectSetVal);

								} else if ((elem.isSelected() && ObjectSetVal
										.equalsIgnoreCase("off"))) {
									elem.click();
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
								} else if (ObjectSetVal.equalsIgnoreCase("on")) {
									elem.click();
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
								} else if ((ObjectSetVal
										.equalsIgnoreCase("off"))) {
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
								} else {
									Update_Report("failed", ObjectSet,
											ObjectSetVal);
								}

								if (captureperform == true) {

									screenshot(loopnum, TScrowcount, TScname);
								}
							} else {
								Update_Report("action1");
							}
						}
						break;

					case "click":
						parentWindowHandle1 = D8.getWindowHandle();
						
						
						try {
							System.out.println("elem type is: "+elem.getAttribute("type").toLowerCase());

							if (elem.getAttribute("type").toLowerCase()
									.equals("file")
									&& BrowserName.equalsIgnoreCase("ff")) {

								JavascriptExecutor executor = (JavascriptExecutor) D8;
								executor.executeScript("arguments[0].click();",
										elem);

								Update_Report("executed");
								break;

							} else if (elem.getAttribute("type").toLowerCase()
									.equals("file")
									&& BrowserName.equalsIgnoreCase("ie")
									&& Integer.parseInt(browserver) == 8) {
								Update_Report("executed");
								
							} else if(elem.getAttribute("type").toLowerCase() //Added by Santosh 11/23/2017
									.equals("radio")) {									
								List<WebElement> allRadios = D8.findElementsByTagName("input");
								for(WebElement radio: allRadios) {									
									if (radio.getAttribute("value").trim().equalsIgnoreCase(ObjectSetVal)){
										JavascriptExecutor executor = (JavascriptExecutor) D8;
										executor.executeScript("arguments[0].click();",
												radio);
										break;
									}
								}
							} else if(elem.getAttribute("type").toLowerCase() //Added by Santosh  to handle link click based on link text 11/23/2017
									.equals("text/css")) {								
								WebElement reqLink = D8.findElementByLinkText(ObjectSetVal);
								JavascriptExecutor executor = (JavascriptExecutor) D8;
								executor.executeScript("arguments[0].click();",reqLink);
										break;								
							} else {		
								//System.out.println("inside else click");								
								elem.click();
								Update_Report("executed");
							}
						} catch (Exception exp1) {
							//System.out.println("inside exception click");							
							elem.click();
							Update_Report("executed");
						}

						if (captureperform == true) {
							screenshot(loopnum, TScrowcount, TScname);
						}
						break;
					
					case "hover":
						JavascriptExecutor js = (JavascriptExecutor) D8;
						String mouseOverScript = "if(document.createEvent){var evObj = document.createEvent('MouseEvents');evObj.initEvent('mouseover', true, false); arguments[0].dispatchEvent(evObj);} else if(document.createEventObject) { arguments[0].fireEvent('onmouseover');}";
						js.executeScript(mouseOverScript, elem);
						Update_Report("executed");
						break;

					case "altclick":
						JavascriptExecutor executor = (JavascriptExecutor) D8;
						executor.executeScript("arguments[0].click();", elem);
						Update_Report("executed");
						break;
					case "enter":
						elem.sendKeys(org.openqa.selenium.Keys.ENTER);
						Update_Report("executed");
						if (captureperform == true) {
							screenshot(loopnum, TScrowcount, TScname);
						}
						break;
					case "setdate":
						Robot robot1 = new Robot();
						ScreenshotTypeFlag = 1;
						String calstring = cCellObjName.toLowerCase();
						if (cCellObjType.equalsIgnoreCase("calendar")
								&& calstring.startsWith("cal_")) {
							try {

								String[] datearray = ObjectSetVal.split("-");
								mm = datearray[0];
								dd = datearray[1];
								yyyy = datearray[2];
								if (Integer.parseInt(mm) > 12
										|| Integer.parseInt(mm) < 1
										|| Integer.parseInt(yyyy) < 1000
										|| Integer.parseInt(yyyy) > 2999
										|| Integer.parseInt(yyyy) < 1700
										|| ((Integer.parseInt(dd) > 30) && (Integer
												.parseInt(dd) == 2
												|| Integer.parseInt(dd) == 4
												|| Integer.parseInt(dd) == 6
												|| Integer.parseInt(dd) == 9 || Integer
												.parseInt(dd) == 11))
										|| (Integer.parseInt(dd) > 28
												&& Integer.parseInt(mm) == 2 && (Integer
												.parseInt(yyyy) % 4 != 0))) {
									Update_Report("invaliddate1");
								} else {
									selectDate(mm, dd, yyyy);
								}

							} catch (Exception e) {
								Update_Report("invaliddate1");
								robot1.keyPress(KeyEvent.VK_ESCAPE);
								robot1.keyRelease(KeyEvent.VK_ESCAPE);

							}
						} else {

							Update_Report("calendaraction");
						}
						break;

					default:
						Update_Report("action");
						break;
					}
				}

			} catch (Exception e) {
				System.out.println(e.toString());
				Update_Report("failed", e);
			}
		} else {
			try {
				switch (ObjectSet.toLowerCase()) {

				case "ok":// Updated on 16.12.2013 for dialog support
					ScreenshotTypeFlag = 1;
					if (cCellObjType.equalsIgnoreCase("dialog")
							|| cCellObjType.equalsIgnoreCase("dialog;")) {
						dialogSwitch = D8.switchTo().alert();
						dialogSwitch.accept();
						Update_Report("executed");
					}
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}

					break;
				case "cancel":// Updated on 16.12.2013 for dialog support
					ScreenshotTypeFlag = 1;

					if (cCellObjType.equalsIgnoreCase("dialog")
							|| cCellObjType.equalsIgnoreCase("dialog;")) {
						dialogSwitch = D8.switchTo().alert();
						dialogSwitch.dismiss();
						Update_Report("executed");
					}

					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					break;

				case "close":// Updated on 16.12.2013 for dialog and page
								// support
					ScreenshotTypeFlag = 1;

					if (cCellObjType.equalsIgnoreCase("dialog")
							|| cCellObjType.equalsIgnoreCase("dialog;")) {

						dialogSwitch = D8.switchTo().alert();
						dialogSwitch.dismiss();
						Update_Report("executed");
						if (captureperform == true) {
							screenshot(loopnum, TScrowcount, TScname);
						}

					}

					else {
						windowFound = 0;
						int windowNums = 0;
						int windowItr = 0;
						String currentWindowHandle = D8.getWindowHandle();
						WebDriver newWindow = null;
						Set<String> al = new HashSet<String>();
						al = D8.getWindowHandles();
						windowNums = al.size();
						Iterator<String> windowIterator = al.iterator();
						if (cCellObjName.equalsIgnoreCase("page;")
								|| cCellObjName.equalsIgnoreCase("page")) {
							if (windowNums == 1) {
								if (captureperform == true) {
									screenshot(loopnum, TScrowcount, TScname);
								}
								D8.close();
								Update_Report("executed");
								windowFound = 1;
							} else {
								int winItr1 = 0;
								String windowHandle = null;
								String tempWindowHandle = null;
								while (winItr1 != windowNums) {
									tempWindowHandle = windowHandle;
									windowHandle = windowIterator.next();
									newWindow = D8.switchTo().window(
											windowHandle);
									if (currentWindowHandle
											.equalsIgnoreCase(windowHandle)) {
										if (winItr1 == 0) {
											if (captureperform == true) {
												screenshot(loopnum,
														TScrowcount, TScname);
											}
											D8.close();
											windowHandle = windowIterator
													.next();
											D8.switchTo().window(windowHandle);
											Update_Report("executed");
											windowFound = 1;
											break;
										} else {
											if (captureperform == true) {
												screenshot(loopnum,
														TScrowcount, TScname);
											}
											D8.close();
											D8.switchTo().window(
													tempWindowHandle);
											Update_Report("executed");
											windowFound = 1;
											break;

										}
									}

									winItr1++;
								}
							}
						} else {
							if (windowNums == 1) {
								if (D8.getTitle().toString()
										.equalsIgnoreCase(ObjectSetVal) == true) {
									if (captureperform == true) {
										screenshot(loopnum, TScrowcount,
												TScname);
									}
									D8.close();
									Update_Report("executed");
									windowFound = 1;
								}

							} else {
								if (cCellObjType.equalsIgnoreCase("page")
										&& D8.getTitle().equalsIgnoreCase(
												cCellObjName) == false) {
									for (windowItr = 0; windowItr < windowNums; windowItr++) {
										windowHandle = windowIterator.next();
										newWindow = D8.switchTo().window(
												windowHandle);
										if (newWindow.getTitle()
												.equalsIgnoreCase(cCellObjName)) {
											if (captureperform == true) {
												screenshot(loopnum,
														TScrowcount, TScname);
											}
											newWindow.close();
											Update_Report("executed");
											D8.switchTo().window(
													currentWindowHandle);
											windowFound = 1;
											break;
										}

									}

								} else {
									if (cCellObjType.equalsIgnoreCase("page")
											&& D8.getTitle()
													.toString()
													.equalsIgnoreCase(
															cCellObjName) == true) {
										// D8.close();
										int winItr1 = 0;
										String windowHandle = null;
										String tempWindowHandle = null;
										while (winItr1 != windowNums) {
											tempWindowHandle = windowHandle;
											windowHandle = windowIterator
													.next();
											newWindow = D8.switchTo().window(
													windowHandle);
											if (currentWindowHandle
													.equalsIgnoreCase(windowHandle)) {
												if (winItr1 == 0) {
													if (captureperform == true) {
														screenshot(loopnum,
																TScrowcount,
																TScname);
													}
													D8.close();
													windowHandle = windowIterator
															.next();
													D8.switchTo().window(
															windowHandle);
													Update_Report("executed");
													windowFound = 1;
													break;
												} else {
													D8.close();
													D8.switchTo().window(
															tempWindowHandle);
													Update_Report("executed");
													windowFound = 1;
													break;

												}
											}

											winItr1++;
										}

									}
								}
							}
						}
						if (windowFound != 1) {
							Update_Report("NoWindowFound");
						}
					}
					break;
				case "closeupload":
					ScreenshotTypeFlag = 1;
					doUploadDownload("closeupload");
					break;

				case "cancelupload":
					ScreenshotTypeFlag = 1;
					doUploadDownload("cancelupload");
					break;

				default:
					System.out.println("Action not supported");
					break;

				}
			}

			catch (NoAlertPresentException e) {

			}

		}

	}

	private void doUploadDownload(String action1) throws Exception {
		// Robot robot = new Robot();
		if (BrowserName.equalsIgnoreCase("ff")) {
			switch (action1) {
			case "upload":
				try {

					Thread.sleep(2000);
					Runtime.getRuntime().exec(
							execpath + " 2 upload " + cCellData + " "
									+ BrowserName.toLowerCase());
					Update_Report("executed");
				} catch (IOException e) {

					// TODO Auto-generated catch block
					System.out.println(e);

				}
				break;

			case "closeupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 closeupload " + cCellData + " "
									+ BrowserName.toLowerCase());
					Update_Report("executed");

				} catch (IOException e) {

					// TODO Auto-generated catch block
					System.out.println(e);

				}
				break;
			case "abortupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 closeupload " + "abort" + " "
									+ BrowserName.toLowerCase());

				} catch (IOException e) {

					// TODO Auto-generated catch block
					System.out.println(e);

				}
				break;

			case "cancelupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 cancelupload " + cCellData + " "
									+ BrowserName.toLowerCase());
					Update_Report("executed");
				} catch (IOException e) {

					// TODO Auto-generated catch block
					System.out.println(e);

				}
				break;

			case "download":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 download " + ObjectSet + " "
									+ BrowserName.toLowerCase() + " "
									+ elem.getAttribute("href"));
					Thread.sleep(4000);
					Update_Report("executed");
				} catch (IOException e) {

					e.printStackTrace();

				}
				break;

			default:
				System.out.println("Action not supported");
				break;
			}
		} else if (BrowserName.equalsIgnoreCase("ie")) {
			switch (action1) {
			case "upload":
				if (Integer.parseInt(browserver) != 8) {
					try {

						Runtime.getRuntime().exec(
								execpath + " 2 upload " + cCellData + " "
										+ BrowserName.toLowerCase());
						Update_Report("executed");
					} catch (IOException e) {

						// TODO Auto-generated catch block
						System.out.println(e);

					}

				} else {
					elem.sendKeys(cCellData);
				}
				break;

			case "closeupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 closeupload " + cCellData + " "
									+ BrowserName.toLowerCase());
					Update_Report("executed");

				} catch (IOException e) {

					// TODO Auto-generated catch block
					System.out.println(e);

				}
				break;

			case "cancelupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 cancelupload " + cCellData + " "
									+ BrowserName.toLowerCase());
					Update_Report("executed");
				} catch (IOException e) {

					// TODO Auto-generated catch block
					System.out.println(e);

				}
				break;

			case "download":

				try {
					Runtime.getRuntime().exec(
							execpath + " 2 download " + ObjectSet + " "
									+ BrowserName.toLowerCase() + " "
									+ elem.getAttribute("href"));
					Update_Report("executed");
				} catch (IOException e) {

					e.printStackTrace();

				}
				break;

			}

		}

		if (captureperform == true) {
			screenshot(loopnum, TScrowcount, TScname);
		}
		// TODO Auto-generated method stub

	}

	public void selectDate(String mm2, String dd2, String yyyy2)
			throws Exception {
		String dateClass = null;
		int objectfound = 0;
		int monthnum1 = 0;
		String usedTitleMonth = null;
		String usedTitleYear = null;
		String usedTitletag = null;
		WebElement prevMonth = null;
		WebElement nextMonth = null;
		WebElement titleYear = null;
		WebElement daytoClick = null;
		int titleYearnum = 0;
		String titleMonthString = "";
		WebElement titleMonth = null;
		String datePickerType = "";
		String locatorNext = "";
		String locatorPrev = "";
		String outerHTMLCalendar = "";
		outerHTMLCalendar = elem.getAttribute("outerHTML");
		elem.click();
		int setmm = Integer.parseInt(mm2);
		int setyyyy = Integer.parseInt(yyyy2);
		int setdd = Integer.parseInt(dd2);
		WebElement titleDefault = null;
		String titleDefaultText;
		// List<WebElement> titlebootYearTags;
		List<WebElement> realCategoryDeciders = new ArrayList<WebElement>();
		// int titlebootYearNum;
		Robot robot1 = new Robot();
		int isDateNotSelected = 0;
		int isDateSelected = 0;

		if (outerHTMLCalendar.toLowerCase().contains("datepicker")) {
			datePickerType = "jquery";
		} else if (outerHTMLCalendar.toLowerCase().contains("type=")
				&& outerHTMLCalendar.toLowerCase().contains("date")) {
			datePickerType = "html5";
		} else {
			datePickerType = "boot";
		}

		switch (setmm) {

		case 1:
			monthpart = "Jan";
			monthpartjap = "1月";
			break;
		case 2:
			monthpart = "Feb";
			monthpartjap = "2月";
			break;
		case 3:
			monthpart = "Mar";
			monthpartjap = "3月";
			break;
		case 4:
			monthpart = "Apr";
			monthpartjap = "4月";
			break;
		case 5:
			monthpart = "May";
			monthpartjap = "5月";
			break;
		case 6:
			monthpart = "Jun";
			monthpartjap = "6月";
			break;
		case 7:
			monthpart = "Jul";
			monthpartjap = "7月";
			break;
		case 8:
			monthpart = "Aug";
			monthpartjap = "8月";
			break;
		case 9:
			monthpart = "Sep";
			monthpartjap = "9月";
			break;
		case 10:
			monthpart = "Oct";
			monthpartjap = "10月";
			break;
		case 11:
			monthpart = "Nov";
			monthpartjap = "11月";
			break;
		case 12:
			monthpart = "Dec";
			monthpartjap = "12月";
			break;
		default:
			Update_Report("invalidmonth");
			break;
		}
		int titleLength;
		switch (datePickerType) {

		case "jquery":
			if (!(lastSelecteddateJQ == setdd && lastSelectedmonthJQ == setmm && lastSelectedyearJQ == setyyyy)) {

				for (String strPrevMonth : setObj.envprevMonth1) {
					try {
						prevMonth = D8
								.findElementByXPath("//a[contains(@class,'"
										+ strPrevMonth + "')]");
						locatorPrev = prevMonth.getAttribute("class")
								.toString();
						objectfound = 1;
						break;
					} catch (Exception e) {
						continue;
					}
				}
				if (objectfound == 0) {
					Update_Report("prevmonth");
					break;
				} else {
					objectfound = 0;
				}
				for (String strNextMonth : setObj.envnextMonth1) {
					try {
						nextMonth = D8
								.findElementByXPath("//a[contains(@class,'"
										+ strNextMonth + "')]");
						locatorNext = nextMonth.getAttribute("class")
								.toString();
						objectfound = 1;
						break;
					} catch (Exception e) {
						continue;
					}
				}

				if (objectfound == 0) {
					Update_Report("nextmonth");

					break;
				} else {
					objectfound = 0;

				}

				for (String strtitleMonth : setObj.envtitleMonth) {

					try {
						titleMonth = D8
								.findElementByXPath("//span[contains(@class,'"
										+ strtitleMonth + "')]");
						titleMonthString = titleMonth.getText();

						usedTitleMonth = strtitleMonth;
						usedTitletag = "span";
						objectfound = 1;
						break;
					} catch (Exception e) {
						try {
							titleMonth = D8
									.findElementByXPath("//select[contains(@class,'"
											+ strtitleMonth + "')]");
							titleMonthString = new Select(titleMonth)
									.getFirstSelectedOption().getText()
									.toString();
							usedTitleMonth = strtitleMonth;
							usedTitletag = "select";
							objectfound = 1;
							break;
						} catch (Exception e4) {

						}
						continue;
					}

				}

				if (objectfound == 0) {
					Update_Report("titlemonth");

					break;
				} else {
					objectfound = 0;
				}
				for (String strtitleYear : setObj.envtitleYear) {

					try {
						titleYear = D8
								.findElementByXPath("//span[contains(@class,'"
										+ strtitleYear + "')]");
						titleYearnum = Integer.parseInt(titleYear.getText());

						usedTitleYear = strtitleYear;
						usedTitletag = "span";
						objectfound = 1;
						break;
					} catch (Exception e) {
						try {
							titleYear = D8
									.findElementByXPath("//select[contains(@class,'"
											+ strtitleYear + "')]");
							titleYearnum = Integer.parseInt(new Select(
									titleYear).getFirstSelectedOption()
									.getText());
							usedTitleYear = strtitleYear;
							usedTitletag = "select";
							objectfound = 1;

							break;
						} catch (Exception e5) {

						}
						continue;
					}
				}
				if (objectfound == 0) {
					Update_Report("titleyear");
					break;
				} else {
					objectfound = 0;

				}

				switch (titleMonthString.toLowerCase()) {

				case "jan":
					monthnum1 = 1;
					break;
				case "1月":
					monthnum1 = 1;
					break;
				case "january":
					monthnum1 = 1;
					break;
				case "feb":
					monthnum1 = 2;
					break;
				case "2月":
					monthnum1 = 2;
					break;
				case "february":
					monthnum1 = 2;
					break;
				case "mar":
					monthnum1 = 3;
					break;
				case "3月":
					monthnum1 = 3;
					break;
				case "march":
					monthnum1 = 3;
					break;
				case "apr":
					monthnum1 = 4;
					break;
				case "4月":
					monthnum1 = 4;
					break;
				case "april":
					monthnum1 = 4;
					break;
				case "may":
					monthnum1 = 5;
					break;
				case "5月":
					monthnum1 = 5;
					break;
				case "june":
					monthnum1 = 6;
					break;
				case "jun":
					monthnum1 = 6;
					break;
				case "6月":
					monthnum1 = 6;
					break;
				case "july":
					monthnum1 = 7;
					break;
				case "jul":
					monthnum1 = 7;
					break;
				case "7月":
					monthnum1 = 7;
					break;
				case "aug":
					monthnum1 = 8;
					break;
				case "august":
					monthnum1 = 8;
					break;
				case "8月":
					monthnum1 = 8;
					break;
				case "sep":
					monthnum1 = 9;
					break;
				case "september":
					monthnum1 = 9;
					break;
				case "9月":
					monthnum1 = 9;
					break;
				case "oct":
					monthnum1 = 10;
					break;
				case "october":
					monthnum1 = 10;
					break;
				case "10月":
					monthnum1 = 10;
					break;
				case "nov":
					monthnum1 = 11;
					break;
				case "november":
					monthnum1 = 11;
					break;
				case "11月":
					monthnum1 = 11;
					break;
				case "dec":
					monthnum1 = 12;
					break;
				case "december":
					monthnum1 = 12;
					break;
				case "12月":
					monthnum1 = 12;
					break;
				default:
					Update_Report("monthnotidentified");
					break;
				}

				if (setyyyy > titleYearnum && setmm > monthnum1) {
					try {
						do {

							nextMonth.click();
							monthnum1++;
							if (monthnum1 == 13)
								monthnum1 = 1;
							nextMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorNext + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
					}

				} else if (setyyyy < titleYearnum && setmm < monthnum1) {

					try {
						do {

							prevMonth.click();
							monthnum1--;
							if (monthnum1 == 0)
								monthnum1 = 12;
							prevMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorPrev + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
					}
				}

				else if (setyyyy == titleYearnum && setmm < monthnum1) {
					try {
						do {
							prevMonth.click();
							monthnum1--;
							if (monthnum1 == 0)
								monthnum1 = 12;
							prevMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorPrev + "')]");
							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");
							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}
						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
					}
				} else if (setyyyy == titleYearnum && setmm > monthnum1) {
					try {
						do {
							nextMonth.click();

							monthnum1++;
							if (monthnum1 == 13)
								monthnum1 = 1;
							nextMonth = D8

							.findElementByXPath("//a[contains(@class,'"
									+ locatorNext + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
					}
				}

				else if (setyyyy > titleYearnum && setmm == monthnum1) {
					try {
						do {

							nextMonth.click();
							monthnum1++;
							if (monthnum1 == 13)
								monthnum1 = 1;
							nextMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorNext + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
					}
				} else if (setyyyy < titleYearnum && setmm == monthnum1) {
					try {
						do {
							prevMonth.click();
							monthnum1--;
							if (monthnum1 == 0)
								monthnum1 = 12;
							prevMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorPrev + "')]");
							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");
							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();

								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
					}

				} else if (setyyyy > titleYearnum && setmm < monthnum1) {
					try {
						do {

							nextMonth.click();
							monthnum1++;
							if (monthnum1 == 13)
								monthnum1 = 1;
							nextMonth = D8

							.findElementByXPath("//a[contains(@class,'"
									+ locatorNext + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}
						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
					}
				} else if (setyyyy < titleYearnum && setmm > monthnum1) {
					try {
						do {
							prevMonth.click();
							monthnum1--;
							if (monthnum1 == 0)
								monthnum1 = 12;
							prevMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorPrev + "')]");
							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");
							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}
						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
					}
				}
				if (isDateNotSelected == 0) {
					if (setyyyy == titleYearnum && setmm == monthnum1) {

						try {

							List<WebElement> daystoClick = D8
									.findElementsByXPath("//a[contains(@class,'"
											+ locatorPrev
											+ "')]"
											+ "/following::a[contains(text(),'"
											+ setdd + "')]");

							List<WebElement> categoryDeciders = D8
									.findElementsByXPath("//a[contains(@class,'"
											+ locatorPrev
											+ "')]"
											+ "/following::a[contains(text(),'"
											+ 15 + "')]");

							for (WebElement categoryDecider : categoryDeciders) {
								if (categoryDecider.getAttribute("href")
										.endsWith("#")
										&& categoryDecider.getText().equals(
												"15")) {
									realCategoryDeciders.add(categoryDecider);
									dateClass = categoryDecider
											.getAttribute("class");

								}
							}

							for (int i2 = 0; i2 < daystoClick.size(); i2++) {
								// Date date = new Date();
								if (today == setdd
										&& Integer.parseInt(daystoClick.get(i2)
												.getText()) == setdd
										&& daystoClick.get(i2)
												.getAttribute("href")
												.endsWith("#")) {

									daytoClick = daystoClick.get(i2);
									daytoClick.click();
									lastSelecteddateJQ = setdd;
									lastSelectedmonthJQ = setmm;
									lastSelectedyearJQ = setyyyy;
									isDateSelected = 1;
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
									break;

								} else if (Integer.parseInt(daystoClick.get(i2)
										.getText()) == setdd
										&& (daystoClick.get(i2)
												.getAttribute("class")
												.equalsIgnoreCase(dateClass) || (daystoClick
												.get(i2).getAttribute("class")
												.equalsIgnoreCase(selectedDateClass))
												&& daystoClick.get(i2)
														.getAttribute("href")
														.endsWith("#"))) {

									daytoClick = daystoClick.get(i2);
									daytoClick.click();
									lastSelecteddateJQ = setdd;
									lastSelectedmonthJQ = setmm;
									lastSelectedyearJQ = setyyyy;
									isDateSelected = 1;
									selectedDateClass = daytoClick
											.getAttribute("class");
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
									break;
								}
							}

						} catch (Exception e1) {
							isDateNotSelected = 1;

						}

					}
					// }
					if (captureperform == true) {
						screenshot(loopnum, TScrowcount, TScname);
					}
					if (isDateSelected == 0) {

						Update_Report("invaliddate");
						robot1.keyPress(KeyEvent.VK_ESCAPE);
						robot1.keyRelease(KeyEvent.VK_ESCAPE);
					} else {
						isDateSelected = 0;
					}

				} else {

					isDateNotSelected = 0;
					Update_Report("invaliddate");
					robot1.keyPress(KeyEvent.VK_ESCAPE);
					robot1.keyRelease(KeyEvent.VK_ESCAPE);

				}
			} else {

				Update_Report("executed", ObjectSet, ObjectSetVal);
				robot1.keyPress(KeyEvent.VK_ESCAPE);
				robot1.keyRelease(KeyEvent.VK_ESCAPE);
				if (captureperform == true) {
					screenshot(loopnum, TScrowcount, TScname);
				}
			}

			break;

		case "boot":

			for (String strPrevMonth : setObj.envprevMonth1) {
				try {
					prevMonth = D8.findElementByXPath("//th[contains(@class,'"
							+ strPrevMonth + "')]");
					locatorPrev = strPrevMonth;
					objectfound = 1;
					break;
				} catch (Exception e) {
					continue;
				}
			}
			if (objectfound == 0) {
				Update_Report("prevmonth");
				break;
			} else {
				objectfound = 0;
			}

			for (String strNextMonth : setObj.envnextMonth1) {
				try {
					nextMonth = D8.findElementByXPath("//th[contains(@class,'"
							+ strNextMonth + "')]");
					locatorNext = strNextMonth;
					objectfound = 1;
					break;
				} catch (Exception e) {
					continue;
				}
			}
			if (objectfound == 0) {
				Update_Report("nextmonth");
				break;
			} else {
				objectfound = 0;
			}

			try {
				titleDefault = D8.findElementByXPath("//th[contains(@class,'"
						+ "switch" + "')]");

			}

			catch (Exception e) {
				Update_Report("titledefault");
				break;
			}
			titleDefaultText = titleDefault.getText().trim();
			titleLength = titleDefaultText.length();
			titleYearnum = Integer.parseInt(titleDefaultText
					.substring(titleLength - 4));
			titleMonthString = titleDefaultText.substring(0, titleLength - 5)
					.trim();
			switch (titleMonthString.toLowerCase()) {

			case "jan":
				monthnum1 = 1;
				break;
			case "1月":
				monthnum1 = 1;
				break;
			case "january":
				monthnum1 = 1;
				break;
			case "feb":
				monthnum1 = 2;
				break;
			case "2月":
				monthnum1 = 2;
				break;
			case "february":
				monthnum1 = 2;
				break;
			case "mar":
				monthnum1 = 3;
				break;
			case "3月":
				monthnum1 = 3;
				break;
			case "march":
				monthnum1 = 3;
				break;
			case "apr":
				monthnum1 = 4;
				break;
			case "4月":
				monthnum1 = 4;
				break;
			case "april":
				monthnum1 = 4;
				break;
			case "may":
				monthnum1 = 5;
				break;
			case "5月":
				monthnum1 = 5;
				break;
			case "june":
				monthnum1 = 6;
				break;
			case "jun":
				monthnum1 = 6;
				break;
			case "6月":
				monthnum1 = 6;
				break;
			case "july":
				monthnum1 = 7;
				break;
			case "jul":
				monthnum1 = 7;
				break;
			case "7月":
				monthnum1 = 7;
				break;
			case "aug":
				monthnum1 = 8;
				break;
			case "august":
				monthnum1 = 8;
				break;
			case "8月":
				monthnum1 = 8;
				break;
			case "sep":
				monthnum1 = 9;
				break;
			case "september":
				monthnum1 = 9;
				break;
			case "9月":
				monthnum1 = 9;
				break;
			case "oct":
				monthnum1 = 10;
				break;
			case "october":
				monthnum1 = 10;
				break;
			case "10月":
				monthnum1 = 10;
				break;
			case "nov":
				monthnum1 = 11;
				break;
			case "november":
				monthnum1 = 11;
				break;
			case "11月":
				monthnum1 = 11;
				break;
			case "dec":
				monthnum1 = 12;
				break;
			case "december":
				monthnum1 = 12;
				break;
			case "12月":
				monthnum1 = 12;
				break;
			default:
				Update_Report("monthnotidentified");
				break;
			}
			try {
				nextMonth = D8.findElementByXPath("//th[contains(@class,'"
						+ locatorNext + "')]");
			} catch (Exception e) {
				Update_Report("nextmonth");
				break;
			}
			try {
				prevMonth = D8.findElementByXPath("//th[contains(@class,'"
						+ locatorPrev + "')]");
			} catch (Exception e) {
				Update_Report("prevmonth");
				break;
			}
			if (setyyyy > titleYearnum && setmm > monthnum1) {
				try {
					do {

						nextMonth.click();
						monthnum1++;
						if (monthnum1 == 13)
							monthnum1 = 1;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();

					}

					while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					Update_Report("invaliddate");
					break;
				}
			} else if (setyyyy < titleYearnum && setmm < monthnum1) {
				try {
					do {

						prevMonth.click();
						monthnum1--;
						if (monthnum1 == 0)
							monthnum1 = 12;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					Update_Report("invaliddate");
					break;
				}
			}

			else if (setyyyy == titleYearnum && setmm < monthnum1) {
				try {
					do {
						prevMonth.click();
						monthnum1--;
						if (monthnum1 == 0)
							monthnum1 = 12;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					Update_Report("invaliddate");
					break;
				}
			} else if (setyyyy == titleYearnum && setmm > monthnum1) {
				try {
					do {
						nextMonth.click();
						monthnum1++;
						if (monthnum1 == 13)
							monthnum1 = 1;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					Update_Report("invaliddate");
					break;
				}

			}

			else if (setyyyy > titleYearnum && setmm == monthnum1) {
				try {
					do {
						nextMonth.click();
						monthnum1++;
						if (monthnum1 == 13)
							monthnum1 = 1;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					Update_Report("invaliddate");
					break;
				}
			} else if (setyyyy < titleYearnum && setmm == monthnum1) {
				try {
					do {
						prevMonth.click();
						monthnum1--;
						if (monthnum1 == 0)
							monthnum1 = 12;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();

					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					Update_Report("invaliddate");
					break;
				}
			} else if (setyyyy > titleYearnum && setmm < monthnum1) {
				try {
					do {
						nextMonth.click();
						monthnum1++;
						if (monthnum1 == 13)
							monthnum1 = 1;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					Update_Report("invaliddate");
					break;
				}
			} else if (setyyyy < titleYearnum && setmm > monthnum1) {
				try {
					do {
						prevMonth.click();
						monthnum1--;
						if (monthnum1 == 0)
							monthnum1 = 12;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					Update_Report("invaliddate");
					break;
				}
			}
			if (setyyyy == titleYearnum && setmm == monthnum1) {
				try {
					List<WebElement> daystoClick = D8
							.findElementsByXPath("//td[contains(text(),'"
									+ setdd + "')]");

					if (daystoClick.size() == 1) {
						daytoClick = daystoClick.get(0);
						daytoClick.click();
						Update_Report("executed", ObjectSet, ObjectSetVal);
					} else if (daystoClick.size() > 1) {
						if (setdd <= 7) {
							for (int i2 = 0; i2 < daystoClick.size(); i2++) {

								if (Integer.parseInt(daystoClick.get(i2)
										.getText()) == setdd) {

									daytoClick = daystoClick.get(i2);
									daytoClick.click();
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
									break;
								}
							}
							// daytoClick = daystoClick.get(0);

						} else {
							for (int i2 = daystoClick.size() - 1; i2 >= 0; i2--) {

								if (Integer.parseInt(daystoClick.get(i2)
										.getText()) == setdd) {

									daytoClick = daystoClick.get(i2);
									daytoClick.click();
									Update_Report("executed", ObjectSet,
											ObjectSetVal);
									break;
								}
							}

						}
					} else {
						Update_Report("invaliddate");
						break;
					}

				} catch (Exception ex1) {
					try {
						List<WebElement> daystoClick = D8
								.findElementsByXPath("//span[contains(text(),'"
										+ setdd + "')]");
						if (daystoClick.size() == 1) {
							daytoClick = daystoClick.get(0);
							daytoClick.click();
							Update_Report("executed", ObjectSet, ObjectSetVal);
						} else if (daystoClick.size() > 1) {
							if (setdd <= 7) {
								for (int i2 = 0; i2 < daystoClick.size(); i2++) {

									if (Integer.parseInt(daystoClick.get(i2)
											.getText()) == setdd) {

										daytoClick = daystoClick.get(i2);
										daytoClick.click();
										Update_Report("executed", ObjectSet,
												ObjectSetVal);
										break;
									}
								}
								// daytoClick = daystoClick.get(0);

							} else {
								for (int i2 = daystoClick.size() - 1; i2 >= 0; i2--) {

									if (Integer.parseInt(daystoClick.get(i2)
											.getText()) == setdd) {

										daytoClick = daystoClick.get(i2);
										daytoClick.click();
										Update_Report("executed", ObjectSet,
												ObjectSetVal);
										break;
									}
								}

							}
						} else {
							Update_Report("invaliddate");
						}

						Update_Report("executed", ObjectSet, ObjectSetVal);
					}

					catch (Exception e) {
						Update_Report("invaliddate");
					}
				}

			}
			if (captureperform == true) {
				screenshot(loopnum, TScrowcount, TScname);
			}
			elem.click();

			break;

		default:
			System.out.println("Notdefined");
			break;

		}

	}

	public static void Func_invokeFunction(String functionName,
			String argumentlist) {
		Object[] argument_list = null;
		int checkNULL = 0;
		int CheckONE = 0;
		Class[] parameterTypes = null;
		if (argumentlist == "") {
			checkNULL = 1;

		} else if (argumentlist.contains(":")) {
			argument_list = argumentlist.split(":");
		} else {
			CheckONE = 1;

		}
		String function_name = functionName;
		try {
			SeleniumWD s1 = new SeleniumWD();
			Method[] declaredMethods = SeleniumWD.class.getDeclaredMethods();
			for (Method m : declaredMethods) {
				System.out.println(m.getName());
				if (checkNULL != 1) {
					parameterTypes = m.getParameterTypes();
				}
				if (checkNULL == 0) {
					if ((m.getName()).equals(function_name)) {
						if (parameterTypes.length > 1) {
							if (parameterTypes.length == argument_list.length) {
								try {
									m.invoke(s1, argument_list);
									Update_Report("executed");
								} catch (Exception e) {
									Update_Report("userdefined");
								}
							}
							break;
						} else if ((m.getName()).equals(function_name)
								&& CheckONE == 1 && parameterTypes.length == 1) {

							try {
								m.invoke(s1, argumentlist);
								Update_Report("executed");
							} catch (Exception e) {
								Update_Report("userdefined");
							}

							break;
						}
					}
				} else if (m.getName().equals(function_name) && checkNULL == 1) {
					try {
						m.invoke(s1);
						Update_Report("executed");
					} catch (Exception e) {
						Update_Report("userdefined");
					}
					break;
				}

			}

		} catch (Exception e) {
			System.out.println(e);
		}
	}

	public static boolean isInteger(String input) {
		try {
			Integer.parseInt(input);
			return true;
		} catch (Exception e) {
			return false;
		}
	}

//--- 2015-06-15 sample method for "callfunction" --- 
	// Single Argument
	public static void Display(String Name)
	{

				System.out.println(Name);
	}

// three
	public static void Display(String Name,String Name2,String Name3)
	{

				System.out.println(Name+ " "+Name2+ " "+Name3+ " ");
	}
	
	public static void CalculationTest(String value1, String value2, String value3, String value4)
	{
		int convert_value1 = Integer.parseInt(value1);
		int convert_value2 = Integer.parseInt(value2);
		int convert_value3 = Integer.parseInt(value3);
		int convert_value4 = Integer.parseInt(value4);
		int Add = 0;
		int Subtraction = 0;
		int Multiplication = 0;
		int Division = 0;

		Add = convert_value1 + convert_value2;
		Subtraction = convert_value3 - convert_value4;
		Multiplication = convert_value1 * convert_value3;
		Division = convert_value2 / convert_value4;

		System.out.println("value1=" + value1);
		System.out.println("value2=" + value2);
		System.out.println("value3=" + value3);
		System.out.println("value4=" + value4);
		System.out.println("Add = value1 + value2");
		System.out.println("Add="+ Add);
		System.out.println("Subtraction = value3 - value4");
		System.out.println("Subtraction=" + Subtraction);
		System.out.println("Multiplication = value1 * value3");
		System.out.println("Multiplication=" + Multiplication);
		System.out.println("Division = value2 / value4");
		System.out.println("Division=" + Division);
		
	}
	
public static void SelectDate(String ExpDate) {
		
		try {
		//Get Expected Date
		String dispExpDate = ExpDate;//dt_Afiliacion.getValue("Exp_Date");
						
		//Convert Expected Date in to Calendar format
		//Expected Date					
		String expDate = dispExpDate.substring(0, 2);
						
		//Expected Month				
		String expMonth =  dispExpDate.substring(3, 5);				
		if(expMonth.trim().equals("01")) {
			expMonth = "Enero";
		}else if(expMonth.trim().equals("02")) {
			expMonth = "Febrero";
		}else if(expMonth.trim().equals("03")) {
			expMonth = "Marzo";
		}else if(expMonth.trim().equals("04")) {
			expMonth = "Abril";
		}else if(expMonth.trim().equals("05")) {
			expMonth = "Mayo";
		}else if(expMonth.trim().equals("06")) {
			expMonth = "Junio";
		}else if(expMonth.trim().equals("07")) {
			expMonth = "Julio";
		}else if(expMonth.trim().equals("08")) {
			expMonth = "Agosto";
		}else if(expMonth.trim().equals("09")) {
			expMonth = "Septiembre";
		}else if(expMonth.trim().equals("10")) {
			expMonth = "Octubre";
		}else if(expMonth.trim().equals("11")) {
			expMonth = "Noviembre";
		}else if(expMonth.trim().equals("12")) {
			expMonth = "Diciembre";
		}	
			
		//Expected Year
		String expYear = dispExpDate.substring(dispExpDate.length()- 4);				
		
		//Construct Calendar Month and Year String
		String expMonthYear = expMonth+","+" "+expYear;
		System.out.println("expMonthYear: "+expMonthYear);
		
		//Get the current Month and Year
		WebElement disp_Date = D8.findElement(By.cssSelector("//.header > td:nth-child(1)"));
		String currMonthYear = disp_Date.getText();
		System.out.println("currMonthYear: "+currMonthYear);
		Thread.sleep(2000);			
		
		//Get the Current Year
		String currYear = currMonthYear.substring(currMonthYear.length()- 4);
		
		System.out.println("currYear: "+currYear);
		System.out.println("expYear: "+expYear);
		
		//Get Month Name				
		String[] exMonth = expMonthYear.trim().split(",");
		String[] cuMonth = currMonthYear.trim().split(",");
		
		//Expected curInt Month
		int curIntMonth = 0;							
		if(cuMonth[0].trim().equals("Enero")) {
			curIntMonth = 1;
		}else if(cuMonth[0].trim().equals("Febrero")) {
			curIntMonth = 2;
		}else if(cuMonth[0].trim().equals("Marzo")) {
			curIntMonth = 3;
		}else if(cuMonth[0].trim().equals("Abril")) {
			curIntMonth = 4;
		}else if(cuMonth[0].trim().equals("Mayo")) {
			curIntMonth = 5;
		}else if(cuMonth[0].trim().equals("Junio")) {
			curIntMonth = 6;
		}else if(cuMonth[0].trim().equals("Julio")) {
			curIntMonth = 7;
		}else if(cuMonth[0].trim().equals("Agosto")) {
			curIntMonth = 8;
		}else if(cuMonth[0].trim().equals("Septiembre")) {
			curIntMonth = 9;
		}else if(cuMonth[0].trim().equals("Octubre")) {
			curIntMonth = 10;
		}else if(cuMonth[0].trim().equals("Noviembre")) {
			curIntMonth = 11;
		}else if(cuMonth[0].trim().equals("Diciembre")) {
			curIntMonth = 12;
		}
		
		//Expected exIntMonth
		int exIntMonth = 0;							
		if(exMonth[0].trim().equals("Enero")) {
			exIntMonth = 1;
		}else if(exMonth[0].trim().equals("Febrero")) {
			exIntMonth = 2;
		}else if(exMonth[0].trim().equals("Marzo")) {
			exIntMonth = 3;
		}else if(exMonth[0].trim().equals("Abril")) {
			exIntMonth = 4;
		}else if(exMonth[0].trim().equals("Mayo")) {
			exIntMonth = 5;
		}else if(exMonth[0].trim().equals("Junio")) {
			exIntMonth = 6;
		}else if(exMonth[0].trim().equals("Julio")) {
			exIntMonth = 7;
		}else if(exMonth[0].trim().equals("Agosto")) {
			exIntMonth = 8;
		}else if(exMonth[0].trim().equals("Septiembre")) {
			exIntMonth = 9;
		}else if(exMonth[0].trim().equals("Octubre")) {
			exIntMonth = 10;
		}else if(exMonth[0].trim().equals("Noviembre")) {
			exIntMonth = 11;
		}else if(exMonth[0].trim().equals("Diciembre")) {
			exIntMonth = 12;
		}
		
		//Iterate through Year				
		while (!expYear.trim().equals(currYear.trim())) {					
			//Click on Previous Year
			WebElement prev_Year =D8.findElement(By.xpath("//*[@id=\"prevYear\"]"));
			prev_Year.click();					
			Thread.sleep(100);
			disp_Date = D8.findElement(By.cssSelector("//.header > td:nth-child(1)"));
			currMonthYear = disp_Date.getText();
			currYear = currMonthYear.substring(currMonthYear.length()- 4);
			prev_Year = null;
		}				
		Thread.sleep(1000);
		
		//Iterate through Month	
		WebElement req_Month = null;				
		if(curIntMonth>exIntMonth) {
			req_Month =D8.findElement(By.xpath("//*[@id=\"prevMonth\"]"));
			
		}else if(curIntMonth<exIntMonth) {
			req_Month =D8.findElement(By.xpath("//*[@id=\"nextMonth\"]"));					
		}
		while (!currMonthYear.trim().contains(expMonth.trim())) {					
			//Click on Previous or Next Month					
			req_Month.click();					
			Thread.sleep(100);
			disp_Date = D8.findElement(By.cssSelector("//.header > td:nth-child(1)"));
			currMonthYear = disp_Date.getText();					
			req_Month = null;
		}				
		Thread.sleep(2000);
		
		//Select the required date
		if(expDate.trim().substring(0,1).equalsIgnoreCase("0")) {
			expDate = expDate.trim().substring(1,2);					
		}
		List<WebElement> all_Days =D8.findElements(By.className("days"));				
		for(WebElement date: all_Days) {
			String reqDate = date.getText();					
			if(reqDate.trim().equals(expDate.trim())) {
				date.click();
				break;
			}				
		}
		}catch(Exception e) {
			
		}

	}


}
