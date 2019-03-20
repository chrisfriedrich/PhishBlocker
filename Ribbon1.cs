using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PhishTest
{

    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        public static string HOST = "http://scopt97.pythonanywhere.com/api/";
        // This would in the future targ
        public static string PHISHING_EMAIL = "cnf@uoregon.edu";

        private Office.IRibbonUI ribbon;

        public MyRibbon()
        {

        }

        public Bitmap GetSendImage(Office.IRibbonControl control)
        {
            return Properties.Resources.phish_icon;
        }

        public Bitmap GetAddImage(Office.IRibbonControl control)
        {
            return Properties.Resources.unsafe_email;
        }

        public Bitmap GetRemoveImage(Office.IRibbonControl control)
        {
            return Properties.Resources.safe_email;
        }

        public void OnAddButton(Office.IRibbonControl control)
        {
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            MailItem mailItem = null;

            if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
            {
                object item = explorer.Selection[1];
                if (item is MailItem)
                {
                    mailItem = item as MailItem;
                }
            }

            if (mailItem != null)
            {
                string senderEmail = "";

                if (mailItem.SenderEmailType == "EX")
                {
                    // Exchange email address

                    string[] directoryParts = mailItem.SenderEmailAddress.Split('=');
                    if (directoryParts.Length > 1)
                    {
                        senderEmail = directoryParts[directoryParts.Length - 1].ToLower();
                    }
                }
                else
                {
                    senderEmail = mailItem.SenderEmailAddress;


                    int successCode = AddPhishingEmail(senderEmail);

                    if (successCode > 0)
                    {
                        MessageBox.Show("'" + senderEmail + "' was successfully added to the list of phishing email addresses.");

                        mailItem.Categories = "Dangerous sender!  Do not open links click images or open attachments.";
                        mailItem.Save();
                    }
                }
            }
        }


        public void OnRemoveButton(Office.IRibbonControl control)
        {
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            MailItem mailItem = null;

            if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
            {
                object item = explorer.Selection[1];
                if (item is MailItem)
                {
                    mailItem = item as MailItem;
                }
            }

            if (mailItem != null)
            {
                string senderEmail = "";

                if (mailItem.SenderEmailType == "EX")
                {
                    // Exchange email address

                    string[] directoryParts = mailItem.SenderEmailAddress.Split('=');
                    if (directoryParts.Length > 1)
                    {
                        senderEmail = directoryParts[directoryParts.Length - 1].ToLower();
                    }
                }
                else
                {
                    senderEmail = mailItem.SenderEmailAddress;


                    int successCode = DeletePhishingEmail(senderEmail);

                    if (successCode > 0)
                    {
                        MessageBox.Show("'" + senderEmail + "' was successfully removed from the list of phishing email addresses.");

                        if (mailItem.Categories != null)
                        {
                            mailItem.Categories = mailItem.Categories.Replace("Dangerous sender!  Do not open links click images or open attachments.", "");
                            mailItem.Save();
                        }
                    }
                }
            }
        }

        public void OnSendButton(Office.IRibbonControl control)
        {
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            Microsoft.Office.Interop.Outlook.AddressEntry addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            string currentUserEmail = "";
            string userFullName = "";

            if (addrEntry.Type == "EX")
            {
                Microsoft.Office.Interop.Outlook.ExchangeUser currentUser = Globals.ThisAddIn.Application.Session.CurrentUser.
                    AddressEntry.GetExchangeUser();
                currentUserEmail = currentUser.PrimarySmtpAddress;
                userFullName = currentUser.FirstName + " " + currentUser.LastName + " (" + currentUser.PrimarySmtpAddress + ")";
            }
            else
            {
                currentUserEmail = addrEntry.Address;
                userFullName = addrEntry.Name + " (" + addrEntry.Address + ")";
            }

            MailItem mailItem = null;

            if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
            {
                object item = explorer.Selection[1];
                if (item is MailItem)
                {
                    mailItem = item as MailItem;
                }
            }

            if (mailItem != null)
            {
                string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

                Microsoft.Office.Interop.Outlook.PropertyAccessor olPA = mailItem.PropertyAccessor;
                string headers = olPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS);

                string senderEmail = "";

                if (mailItem.SenderEmailType == "EX")
                {
                    // Exchange email address

                    string[] directoryParts = mailItem.SenderEmailAddress.Split('=');
                    if (directoryParts.Length > 1)
                    {
                        senderEmail = directoryParts[directoryParts.Length - 1].ToLower();
                    }
                }
                else
                {
                    senderEmail = mailItem.SenderEmailAddress;


                    int successCode = SubmitPhishingEmail(senderEmail);
                    
                    if(successCode > 0)
                    {
                        MessageBox.Show("'" + senderEmail + "' was successfully added to the list of phishing email addresses and forwarded to central IS (phishing.uoregon.edu).");

                        string subject = "Phishing Email added to Phishing Email List";

                        string messageBody = "Phishing Email added to Phishing Email List '" + senderEmail + "'.<br />";
                        messageBody += "Added By: " + userFullName + " at " + DateTime.Now.ToString() + "<br /><br />"; 
                        messageBody += "Headers: <br />" + headers;


                        Microsoft.Office.Interop.Outlook.MailItem eMail = (Microsoft.Office.Interop.Outlook.MailItem)
                            Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                        eMail.Subject = subject;
                        eMail.To = PHISHING_EMAIL;
                        eMail.Body = messageBody;
                        eMail.HTMLBody = messageBody;
                        eMail.Attachments.Add(mailItem);

                        eMail.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh;
                        ((Microsoft.Office.Interop.Outlook._MailItem)eMail).Send();
                    }
                }
            }
        }

        protected int CheckEmailAddress(string email)
        {
            int status = -1;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(HOST + "add-phish?email=" + email);
            request.AllowAutoRedirect = true;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            string rawStatusCode = response.StatusCode.ToString();

            if (rawStatusCode == "OK")
            {
                Stream receiveStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(receiveStream, Encoding.UTF8);

                string results = reader.ReadToEnd();

                if (results == "confirmed")
                {
                    status = 1;
                }
                else if (results == "unconfirmed")
                {
                    status = 0;
                }
                else
                {
                    status = -1;
                }
            }
            return status;
        }


        protected int SubmitPhishingEmail(string email)
        {
            int status;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(HOST + "add-phish?email=" + email);
            request.AllowAutoRedirect = true;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            string rawStatusCode = response.StatusCode.ToString();

            if(rawStatusCode == "OK")
            {
                status = 1;
            }
            else
            {
                status = -1;
            }

            Stream receiveStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(receiveStream, Encoding.UTF8);

            string results = reader.ReadToEnd();

            return status;
        }

        protected int AddPhishingEmail(string email)
        {
            int status = -1;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(HOST + "add-phish?email=" + email);
            request.AllowAutoRedirect = true;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            string rawStatusCode = response.StatusCode.ToString();

            if (rawStatusCode == "OK")
            {
                status = 1;

                Stream receiveStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(receiveStream, Encoding.UTF8);

                string results = reader.ReadToEnd();

                if (results == "success")
                {
                    status = 1;
                }
                else
                {
                    status = 1;
                }
            }
            else
            {
                status = -1;
            }

            return status;
        }


        protected int DeletePhishingEmail(string email)
        {
            int status = -1;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(HOST + "del-phish?email=" + email);
            request.AllowAutoRedirect = true;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            string rawStatusCode = response.StatusCode.ToString();

            if (rawStatusCode == "OK")
            {
                status = 1;

                Stream receiveStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(receiveStream, Encoding.UTF8);

                string results = reader.ReadToEnd();

                if (results == "success")
                {
                    status = 1;
                }
                else
                {
                    status = 1;
                }
            }
            else
            {
                status = -1;
            }

            return status;
        }


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PhishTest.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
