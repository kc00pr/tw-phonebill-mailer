package com.thoughtworks;

import java.util.Properties;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

public class Mail {

    private static Session session;
    private static Mail instance = null;
    private static Properties smtpProperties = new Properties();

    private Mail()
    {
        final String username = "khushroo.cooper@gmail.com";
        final String password = "xx";

        smtpProperties.put("mail.smtp.auth", "true");
        smtpProperties.put("mail.smtp.starttls.enable", "true");
        smtpProperties.put("mail.smtp.host", "smtp.gmail.com");
        smtpProperties.put("mail.smtp.port", "587");

        session = Session.getInstance(smtpProperties,
                new javax.mail.Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });
    }

    public static void Send(String payload, PhoneBill phoneBill)
    {
        if (instance == null) instance = new Mail();
        try {
            System.out.println("Sending Mail to " + phoneBill.employeeName);
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress("khushroo.cooper@gmail.com"));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(phoneBill.employeeEmail));
            message.setSubject("Mobile Phone Bill for " + phoneBill.month);
//            message.setText(payload);
            message.setContent(payload, "text/html");

            Transport.send(message);
            System.out.println("Mail Sent");

        } catch (MessagingException e) {
            throw new RuntimeException(e);
        }
    }

}
