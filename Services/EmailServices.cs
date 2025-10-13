using Microsoft.Office.Interop.Outlook;
using System;
using System.Runtime.InteropServices;

namespace AccessManager.Services
{
    public class EmailService
    {
        public void CreateOutlookEmail(string to, string subject, string body)
        {
            try
            {
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    throw new InvalidOperationException("Outlook не установлен или недоступен.");

                dynamic outlookApp = Activator.CreateInstance(outlookType);
                dynamic mailItem = outlookApp.CreateItem(0); // 0 = olMailItem

                mailItem.To = to;
                mailItem.Subject = subject;
                mailItem.Body = body;
                mailItem.Display(true); // false = не редактировать, true = открыть для пользователя

                Marshal.ReleaseComObject(mailItem);
                Marshal.ReleaseComObject(outlookApp);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Ошибка при создании письма Outlook: {ex.Message}", ex);
            }
        }
    }
 }

