using System;
using Microsoft.Office.Interop.Outlook;

namespace AccessManager.Services
{
    public class EmailService
    {
        public void CreateOutlookEmail(string to, string subject, string body)
        {
            try
            {
                Application outlookApp = new Application();
                MailItem mail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                mail.To = to;
                mail.Subject = subject;
                mail.Body = body;

                // ✅ Показываем окно письма пользователю
                mail.Display(false);
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show("Ошибка при создании письма: " + ex.Message);
            }
        }
    }
}

