using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace CourseReportEmailer.Workers
{
    internal class EnrollmentDetailReportEmailSender
    {
        public void Send(string fileName)
        {
            SmtpClient client = new SmtpClient("smtp-mail.outlook.com"); //using System.Net.Mail
                                                                         //smtp server is the machine that takes care of
                                                                         //the entire email delivery process
                                                                         //in our case we're gonna use smtp-mail.outlook.com in order to send mail

            //need to specify the port 587 - default port for mail submission
            client.Port = 587;
            client.DeliveryMethod = SmtpDeliveryMethod.Network; //that's how we're gonna to deliver an email
                                                                //there are 3 options - Network is default = it will be sent via network to smpt server
            client.UseDefaultCredentials = false; //we want to create our own credentials instead



            NetworkCredential credentials = new NetworkCredential("m.vikhrenko+test@gmail.com", "password_placeholder"); //using System.Net
            client.EnableSsl = true; //Secure Sockets Layer = this will ensure our info goes to the server in the secure manner
                                     //it will create a secure connection between our client and smtp server '
                                     //and while they are communication will go securely through
            client.Credentials = credentials;



            //next create an actual message 
            MailMessage message = new MailMessage("m.vikhrenko+test@gmail.com", "m.vikhrenko+test2@gmail.com"); //mail FROM and mail TO
            message.Subject = "Enrollment Details Report"; //comes from requirements
            message.IsBodyHtml = true; //because we're gonna put some html code => it allows more flexibility than just a text
            message.Body = "Hi,<br><br>Attached please find the enrollemnt details report.<br><br>Please let me know if there are any questions.<br><br>Best,<br><br>Mira";

            Attachment attachment = new Attachment(fileName);
            message.Attachments.Add(attachment); //you can have more than 1
            client.Send(message);

        }
    }
}
