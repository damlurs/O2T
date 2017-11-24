package SeleniumWD;

import java.util.List;
import org.openqa.selenium.WebElement;

public class Setting {
	public String utilityFilePath = "C:/SeleniumProjects/O2T/SeleniumUtility.xls";
	String[] envprevMonth1 = { "prev", "Prev" }; // Specify a class name through which the framework can identify the
													// image representing the previous month
	String[] envnextMonth1 = { "next", "Next" };// Specify a class name through which the framework can identify the
												// image representing the next month
	String[] envtitleMonth = { "month" };// Specify a class name through which we can identify the title month element
											// in calendar control element
	String[] envtitleYear = { "year" };// Specify a class name through which we can identify the title year element in
										// calendar control element
	
	public static void main(String[] args) {
		try {
			SeleniumWD o2t = new SeleniumWD();		
			o2t.ReadUtilFile();	
		}catch(Exception e) {
			e.printStackTrace();
		}
		
	}
}
