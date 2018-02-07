using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace HelloWorld_EWS_app
{
    class HelloWorld_EWS_app
    {
        static void Main(string[] args)
        {
            /* instantiate the ExchangeService object with the service version you intend to target. 
             * This example targets the earliest version of the EWS schema.
             */
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            /* pass explicit credentials and set the credentials for your mailbox account. 
             * The user name should be the user principal name
             */
            service.Credentials = new WebCredentials("miguel.gonzalez@aps-holding.com", "Maria2018");

            /* If you want to see the calls being made, add the following two lines of code before the AutodiscoverUrl method is called.
             * Then press F5.This will trace out the EWS requests and responses to the console window.
             */
            //service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;

            /* The AutodiscoverUrl method on the ExchangeService object performs a series of calls to the Autodiscover service to get the service URL.
             * If this method call is successful, the URL property on the ExchangeService object will be set with the service URL. Pass the user’s email address 
             * and the RedirectionUrlValidationCallback to the AutodiscoverUrl method
             */
            service.AutodiscoverUrl("miguel.gonzalez@aps-holding.com", RedirectionUrlValidationCallback);

            /* instantiate a new EmailMessage object and pass in the service object you created 
             * to have an email message on which the service binding is set.Any calls initiated on the EmailMessage object will be targeted at the service.
             */
           EmailMessage email = new EmailMessage(service);

            /* set the To: line recipient of the email message. To do this, use your SMTP address.
             */
            email.ToRecipients.Add("miguel.gonzalez@aps-holding.com");

            /* Set the subject and the body of the email message.
             */
            email.Subject = "HelloWorld1";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API.");

            /* The Send method will call the service and submit the email message for delivery.
             */
            email.Send();            

        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
