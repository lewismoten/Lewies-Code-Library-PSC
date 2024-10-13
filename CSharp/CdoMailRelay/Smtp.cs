using System;
using System.Web.Mail;

namespace SendMail
{
	public class Smtp
	{

		public Smtp()
		{
		}
		public static void Send(System.Web.Mail.MailMessage message, string server, string username, string password)
		{
			const long cdoBasic = 1;
			System.Web.Mail.SmtpMail.SmtpServer = server;
			string schema = "http://schemas.microsoft.com/cdo/configuration/";
			message.Fields[schema + "smtpauthenticate"] = (int)cdoBasic;
			message.Fields[schema + "sendusername"] = username;
			message.Fields[schema + "sendpassword"] = password;
			System.Web.Mail.SmtpMail.Send(message);
		}
		public static void Send(System.Web.Mail.MailMessage message, string server)
		{
			System.Web.Mail.SmtpMail.SmtpServer = server;
			System.Web.Mail.SmtpMail.Send(message);
		}
		public static void Send(System.Web.Mail.MailMessage message)
		{
			// user SMTP service on local machine
			System.Web.Mail.SmtpMail.SmtpServer = Environment.MachineName;
			System.Web.Mail.SmtpMail.Send(message);
		}
	}
}
