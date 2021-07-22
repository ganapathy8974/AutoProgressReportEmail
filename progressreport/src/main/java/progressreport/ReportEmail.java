package progressreport;

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class ReportEmail {

	Session session;    
    private final String from = "ganapathy8974@gmail.com";
    private final String password = "********";
    //Authentication Method this is create a session once successfully authenticated 
    public void MailAuthentication() {
    	
        String host = "smtp.gmail.com";    
        Properties properties = System.getProperties();

        // Setup mail server
        properties.put("mail.smtp.host", host);
        properties.put("mail.smtp.port", "465");
        properties.put("mail.smtp.ssl.enable", "true");
        properties.put("mail.smtp.auth", "true");

        // Get the Session object.// and pass username and password
       this.session = Session.getInstance(properties, new javax.mail.Authenticator() {

            protected PasswordAuthentication getPasswordAuthentication() {

                return new PasswordAuthentication(from, password);

            }

        });
       
        session.setDebug(false);
    }
    //This Method Contain Email Messages
    public void EmailMessage(String to, String name,DataSource excelReport) {

        try {            
            MimeMessage message = new MimeMessage(session);            
            message.setFrom(new InternetAddress(from));           
            message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));           
            message.setSubject(name+"'s Report Card");    
          
            //Creating body part objects
            MimeBodyPart messageBody = new MimeBodyPart();
            MimeBodyPart attachmentBody = new MimeBodyPart();
            
            //set the message to body one.
            messageBody.setText("Dear Parent "+name+",\n\nPFA of your son's report card. \n\n"            		
            		+ "Thanks & Regards, \r\n"
            		+ "S.Ganapathy Ramasubramanian,\r\n"
            		+ "Mobile: 1123456789,");
            
            //Set the attachment here
            attachmentBody.setDataHandler(new DataHandler(excelReport));
            attachmentBody.setFileName("Report Card.xlsx");
            
            //Create a Multi Part object to set the Message Body objects
            Multipart mp = new MimeMultipart();
            mp.addBodyPart(messageBody);
            mp.addBodyPart(attachmentBody);
            
            //set a Multi Part to Mine Message Object.
            message.setContent(mp);
            message.saveChanges();
            
            System.out.print("Email Sending to "+to);
            // Send message
            Transport.send(message);            
        } catch (MessagingException mex) {
            mex.printStackTrace();
        }

    } 

}
