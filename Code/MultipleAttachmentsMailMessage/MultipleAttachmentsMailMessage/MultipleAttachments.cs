using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Activities;
using System.ComponentModel;


namespace MultipleAttachmentsMailMessage
{ 
    public class MultipleAttachments : CodeActivity
{
    [Category("AttachmentFolderPath")]
    [RequiredArgument]
    public InArgument<string> InputDirPath { get; set; }

    [Category("Email")]
    public InArgument<string> Body { get; set; }

    [Category("Email")]
    [RequiredArgument]
    public InArgument<string> Subject { get; set; }

    [Category("Host")]
    public InArgument<int> Port { get; set; }

    [Category("Host")]
    public InArgument<string> Server { get; set; }

    [Category("Logon")]
    public InArgument<string> Email { get; set; }

    [Category("Logon")]
    public InArgument<string> Password { get; set; }

    [Category("Receiver")]
    public InArgument<string> Bcc { get; set; }

    [Category("Receiver")]
    public InArgument<string> Cc { get; set; }

    [Category("Receiver")]
    public InArgument<string> To { get; set; }

    [Category("Sender")]
    public InArgument<string> From { get; set; }

    [Category("Sender")]
    public InArgument<string> Name { get; set; }

    //[Category("Output")]
    //public OutArgument<double> ResultNumber { get; set; }
    

    protected override void Execute(CodeActivityContext context)
    {
        try
        {

            var body = Body.Get(context);
            var subject = Subject.Get(context);
            var port = Port.Get(context);
            var server = Server.Get(context);
            var email = Email.Get(context);
            var password = Password.Get(context);
            var bcc = Bcc.Get(context);
            var cc = Cc.Get(context);
            var to = To.Get(context);
            var from = From.Get(context);
            var name = Name.Get(context);
            var inputdirpath = InputDirPath.Get(context);

            String[] files = Directory.GetFiles(inputdirpath);

            MailMessage mail = new MailMessage(from,to);

                mail.Subject = subject;
                mail.Body = body;
                

            foreach (string file in files)
            {
                mail.Attachments.Add(new System.Net.Mail.Attachment(file));
            }

            SmtpClient smtp = new SmtpClient();
            smtp.Host = server;
            smtp.Port = port;

            smtp.Credentials = new NetworkCredential(email, password);

            smtp.EnableSsl = true;
            
            smtp.Send(mail);
        }
        catch (Exception e)
        {
            Console.WriteLine("Exception Caught at: "+e.Source);
            Console.WriteLine("Exception Message: " + e.Message);
        }
    }
}
}