package OutlookTests;

import java.io.File;

import org.testng.annotations.Test;

import OutlookProvider.CalendarItems;
import OutlookProvider.CreateAndSaveOutlookMessageFile;
import OutlookProvider.OutlookCommon;
import OutlookProvider.ReadOffice365OSTFile;
import OutlookProvider.ReadOutlookOSTFile;
import OutlookProvider.ReadOutlookPSTFile;
import OutlookProvider.SearchMessagesAndFoldersInAPST;
import OutlookProvider.SendEmail;

public class OutlookTests {
	// Message
	static String from ="sandeep.dhamale@outlook.com";
	static String to = "sandeep.dhamale@outlook.com";
	static String toMultiple = "sandeep.dhamale@outlook.com; dhamale_sandeep@yahoo.com";
	static String subject = "Test Subject.";
	static String body = "This is a body of message.";
	
	// Calendar Items
	static String location = "LAKE ARGYLE WA 6743";
	static int startYear = 2020;
	static int startMonth = 10;
	static int startDate = 3;

	static int endYear = 2020;
	static int endMonth = 10;
	static int endDate = 3;
	
	static int timeHours = 0;
	static int timeMinutes = 0;
	static int timeSeconds = 0;
	
	static String attendeeEmail = "attendee_address@domain.com";
	static String organizerEmail = "organizer@domain.com";
	
	static String appointmentFileName = "CalendarItem_Test.ics";
	static String displayReminderFileName = "CalendarDisplayReminder_Test.ics";
	
	static String filePath = OutlookCommon.getSharedDataDir();
	
	static String pstFilePath = "C:\\Users\\sande\\AppData\\Local\\Microsoft\\Outlook";
	static String pstFileName = "sandeep.dhamale@hotmail.com.pst";
	static String ostFileName = "sandeep.dhamale@hotmail.com.ost";
	
	static String userName = "Sandeep.Dhamale@hotmail.com";
	static String password = "MyPass";
	
	@Test
	public void Outlook_CreateAndSaveOutlookFile()
	{	
		CreateAndSaveOutlookMessageFile.createAndSaveOutlookMessageFile( filePath + File.separator, from, to, subject, body, "Test_Mail.msg");
	}
	
	@Test
	public void Outlook_CalendarItems()
	{
		CalendarItems.creatAndSaveCalendarItems( filePath + File.separator, location, subject, 
					body, startYear, startMonth, startDate,	endYear, endMonth, endDate, appointmentFileName);
	}
	
	@Test 
	public void Outlook_SendEmailSMTP()
	{
		SendEmail.SendEmailSMTP();
	}
	
	@Test 
	public void Outlook_ReadOutlookPSTFile()
	{
		ReadOutlookPSTFile.loadAPSTFile(pstFilePath + File.separator, pstFileName);
	}
	
	@Test 
	public void Outlook_ReadOutlookPSTFileFolders()
	{
		ReadOutlookPSTFile.displayFolderAndMessageInformationForPSTFile(pstFilePath + File.separator, pstFileName);
	}
	
	@Test 
	public void Outlook_ReadOutlookPSTFileSubFolders()
	{
		// Root - Mailbox: "AAAAABUA00HkrkZJjZBWkqQYFK6iIAAA"
		// Root - Public: "AAAAABUA00HkrkZJjZBWkqQYFK4CIAAA"
		ReadOutlookPSTFile.parseSearchableFolders(pstFilePath + File.separator, pstFileName, "AAAAABUA00HkrkZJjZBWkqQYFK4CIAAA");
	}
	
	@Test 
	public void Outlook_ReadOutlookOSTFileFolders()
	{
		ReadOutlookOSTFile.readAnOSTFile(pstFilePath + File.separator, ostFileName);
	}
	
	// NOt able to search subfolders e.g. Inbox with following method
	@Test 
	public void Outlook_ReadOutlookOSTFileFoldersAndMessages()
	{
		SearchMessagesAndFoldersInAPST.ReadOstFoldersAndMessages(pstFilePath + File.separator, ostFileName);
	}
	
	@Test 
	public void Outlook_ReadOutlookOSTFile()
	{
		ReadOffice365OSTFile.LoadOffice365Outlook(userName, password);
		ReadOffice365OSTFile.OutLookReader_imaps("INBOX"); //pstFilePath + File.separator, ostFileName
	}
	
	@Test 
	public void Outlook_ReadOutlookOSTRootSubFolders()
	{
		ReadOffice365OSTFile.LoadOffice365Outlook(userName, password);
		ReadOffice365OSTFile.GetRootSubFolders(); //pstFilePath + File.separator, ostFileName
	}
	
	@Test 
	public void Outlook_ReadOutlookOSTRootNestedSubFolders()
	{
		ReadOffice365OSTFile.LoadOffice365Outlook(userName, password);
		ReadOffice365OSTFile.GetRootNestedSubFolders(); //pstFilePath + File.separator, ostFileName
	}
	
	@Test 
	public void Outlook_ReadOutlookOSTSubFolders()
	{
		ReadOffice365OSTFile.LoadOffice365Outlook(userName, password);
		ReadOffice365OSTFile.GetSubFolders("TestFolder"); //pstFilePath + File.separator, ostFileName
	}
	
	@Test 
	public void Outlook_ReadOutlookOSTSearchBySubject()
	{
		ReadOffice365OSTFile.LoadOffice365Outlook(userName, password);
		ReadOffice365OSTFile.SearchEmail("Inbox", "Lenovo", null, null, "", ""); //pstFilePath + File.separator, ostFileName
	}
}
