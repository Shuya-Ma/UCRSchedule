package schedule;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.time.LocalDate;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session; 
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Schedule {
	
	public static void writeToExcel(XSSFSheet sheet_out, XSSFSheet sheet_cal, int i, int colNum_out, CellStyle dateStyle){
		int index = 1;
		String[] day_of_the_week = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
		while (index < colNum_out) {
			if (sheet_out.getRow(index).getCell(0).toString().equals(sheet_cal.getRow(i).getCell(0).toString()) && sheet_out.getRow(index).getCell(1).toString().equals(sheet_cal.getRow(i).getCell(1).toString())) {
				
				Date date = sheet_cal.getRow(i).getCell(2).getDateCellValue();
				sheet_out.getRow(index).createCell(colNum_out).setCellValue(date);
				sheet_out.getRow(index).createCell(colNum_out+1).setCellValue(day_of_the_week[date.getDay()]);
				sheet_out.getRow(index).createCell(colNum_out+2).setCellValue(sheet_cal.getRow(i).getCell(3).getStringCellValue());
				
				sheet_out.getRow(index).getCell(colNum_out).setCellStyle(dateStyle);
				
				Date today = new Date();
				System.out.println(today);
				
				Calendar calendar = Calendar.getInstance();
				calendar.setTime(today);
		        calendar.set(Calendar.HOUR_OF_DAY, 0);
		        calendar.set(Calendar.MINUTE, 0);
		        calendar.set(Calendar.SECOND, 0);
		        calendar.set(Calendar.MILLISECOND, 0);
		        calendar.add(Calendar.DATE, 1);
		        Date tomorrow = calendar.getTime();
		        System.out.println(date);
		        System.out.println(tomorrow);
		        
				System.out.println(date.compareTo(tomorrow));

				return;
			}
			else {
				index++;
			}
		}
	}
	
		
		
	public static void sendEmail(Session session, String sender, String receiver, String name, String date, String time) {
		try {

	        Message message = new MimeMessage(session);
	        message.setFrom(new InternetAddress(sender));
	        message.setRecipients(Message.RecipientType.TO,
	            InternetAddress.parse(receiver)); 
	        
	        message.setSubject("**Important** Profile Study Confirmed Schedule");
//	        
//	        message.setText("HI you have done sending mail with outlook");
	        
	        // This mail has 2 part, the BODY and the embedded image
	        MimeMultipart multipart = new MimeMultipart("related");

	        // first part (the html)
	        BodyPart messageBodyPart = new MimeBodyPart();
	        
//	        String htmlText = "<H1>Hello "+ name +"</H1><img src=\"cid:image\"><H1>End of the test</H1>";
	        
	     
	        String htmlText = "<p> Hello " + name + ",</p>" + "<p>This is Bernice from the Social Neuroscience Lab. Thank you again for being interested in our study. Below is your assigned schedule:</p> <p><b>Session 1: " + date + " " + time + "</b></p> </p> <p><b>Session 2: " + date + " " + time + "</b></p><p> The study will be taken place at <b>Olmsted Hall 2133</b>. I've attached a map to help you find the building.</p>"+"<img src=\"cid:image\">"+"<p><b><u>Please reply to this email to confirm your appointment.</u></b> If you need to reschedule for either session 1 or session 2, please let me know as well!</p><p>Best,<p><p>Bernice</p>";	        
	        messageBodyPart.setContent(htmlText, "text/html");
	        // add it
	        multipart.addBodyPart(messageBodyPart);

	        // second part (the image)
	        messageBodyPart = new MimeBodyPart();
	        DataSource fds = new FileDataSource("/Users/ShuyaMa/eclipse-workspace/UCRSchedule/src/schedule/Picture1.jpg");

	        messageBodyPart.setDataHandler(new DataHandler(fds));
	        messageBodyPart.setHeader("Content-ID", "<image>");

	        // add image to the multipart
	        multipart.addBodyPart(messageBodyPart);

	        // put everything together
	        message.setContent(multipart);
	        
	        

	        Transport.send(message);

	        System.out.println("Done");

	    } catch (MessagingException e) {
	        throw new RuntimeException(e);
	    }
		
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		
		
		//Excel!!!!!!!!!!!!!!!
		// Change the filePath to your excel file path
		String filePath = ""; // info file
		String calFile = ""; // calendar file
		
		try {
			FileInputStream inputStream_out = new FileInputStream(new File(filePath));
			XSSFWorkbook workbook_out = new XSSFWorkbook(inputStream_out);
			XSSFSheet sheet_out = workbook_out.getSheetAt(0);
			
			FileInputStream inputStream_cal = new FileInputStream(new File(calFile));
			XSSFWorkbook workbook_cal = new XSSFWorkbook(inputStream_cal);
			XSSFSheet sheet_cal = workbook_cal.getSheetAt(0);
			
			int rowNum_cal = sheet_cal.getLastRowNum()+1;			
			int colNum_out = sheet_out.getRow(0).getLastCellNum();			
			
			sheet_out.getRow(0).createCell(colNum_out).setCellValue("date");
			sheet_out.getRow(0).createCell(colNum_out+1).setCellValue("day of the week");
			sheet_out.getRow(0).createCell(colNum_out+2).setCellValue("time");
		
			CreationHelper creationHelper = workbook_out.getCreationHelper();
			CellStyle dateStyle = workbook_out.createCellStyle();
			dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("mm/dd/yyyy"));
			
			for (int i = 1; i < rowNum_cal; i++) {
				writeToExcel(sheet_out, sheet_cal, i, colNum_out, dateStyle);
			}
						
			inputStream_out.close();
			inputStream_cal.close();
			
			FileOutputStream outputStream = new FileOutputStream(new File(filePath));
			workbook_out.write(outputStream);
			outputStream.close();
			workbook_out.close();
			System.out.println("done");
			
		}
		catch(FileNotFoundException e) {
			e.printStackTrace();
		}
		catch(IOException e) {
			e.printStackTrace();
		}
		
		
		//Outlook!!!!!!!!!!!!!!!
//		String sender = ""; sender's email
//		String password = ""; sender's email account password
//		
//		Properties props = new Properties();
//	    props.put("mail.smtp.auth", "true");
//	    props.put("mail.smtp.starttls.enable", "true");
//	    props.put("mail.smtp.host", "smtp-mail.outlook.com");
//	    props.put("mail.smtp.port", "587");
//
//	    Session session = Session.getInstance(props,
//	      new javax.mail.Authenticator() {
//	    	@Override
//	        protected PasswordAuthentication getPasswordAuthentication() {
//	            return new PasswordAuthentication(sender, password);
//	        }
//	      });
//	    session.setDebug(true);
//	    
//	    String receiver = ""; receiver's account
//		String name = "Shuya";
//        String date = "6/29(Friday)";
//        String time = "9:00-10:00AM";
//		
//	    receiver = "shuyama@shuyama.me";
//		name = "Shuya";
//        date = "6/29(Friday)";
//        time = "9:00-10:00AM";
//        sendEmail(session, sender, receiver, name, date, time);
//        receiver = "s.ma@wustl.edu";
//		name = "Shuya2";
//        date = "6/30(Saturday)";
//        time = "9:30-11:00AM";
//        sendEmail(session, sender, receiver, name, date, time);
//

	    

	}

}
