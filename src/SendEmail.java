import java.util.*;  
import javax.mail.*;  
import javax.mail.internet.*;  
  
public class SendEmail 
{
	public static void main(String [] args)
	{  
	      // email ID of Recipient. 
	      String to = "ushakhedkaruck@gmail.com"; 
	  
	      // email ID of  Sender. 
	      String from = "ushakhedkaruck@gmail.com"; 
	  
	      // using host as localhost 
	      String host = "localhost"; 
	  
	      // Get system properties
	      Properties properties = System.getProperties();

	      // Setup mail server
	      properties.setProperty("mail.smtp.host", host);

	      // Get the default Session object.
	      Session session = Session.getDefaultInstance(properties);

	      try {
	         // Create a default MimeMessage object.
	         MimeMessage message = new MimeMessage(session);

	         // Set From: header field of the header.
	         message.setFrom(new InternetAddress(from));

	         // Set To: header field of the header.
	         message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));

	         // Set Subject: header field
	         message.setSubject("This is the Subject Line!");

	         // Send the actual HTML message, as big as you like
	         message.setContent("<h1>This is actual message</h1>", "text/html");

	         // Send message
	         Transport.send(message);
	         System.out.println("Sent message successfully....");
	      } catch (MessagingException mex) {
	         mex.printStackTrace();
	      }

	   
	}  

}
