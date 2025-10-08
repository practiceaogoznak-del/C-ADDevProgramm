using System;
using Microsoft.Office.Interop.Outlook;

namespace AccessManager.Services
{
    public class EmailService
    {
        public void SendOutlookEmail(string to, string subject, string body)
        {
            try
            {
                Application outlookApp = new Application();
                MailItem mail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
                mail.To = to;
                mail.Subject = subject;
                mail.Body = body;
                mail.Send(); // сразу отправить
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show("Ошибка при отправке письма: " + ex.Message);
            }
        }
    }
}
