using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Net;
using System.IO;
using System.Windows.Forms;

namespace PhishTest
{
    public partial class ThisAddIn
    {
        public static string HOST = "http://scopt97.pythonanywhere.com/api/";
        private Outlook.MailItem currentItem;
        public Outlook.MailItem CurrentItem
        {
            get
            {
                return currentItem;
            }
            set
            {
                currentItem = value;
            }
        }

        Outlook.Inspectors inspectors;
        Outlook.Explorer currentExplorer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook
                .ExplorerEvents_10_SelectionChangeEventHandler
                (CurrentExplorer_Event);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon();
        }

        private void ThisAddInFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
        {
            Outlook.ContactItem myItem = (Outlook.ContactItem)e.OutlookItem;

            if (myItem != null)
            {
                if ((myItem.BusinessAddress != null &&
                        myItem.BusinessAddress.Trim().Length > 0) ||
                    (myItem.HomeAddress != null &&
                        myItem.HomeAddress.Trim().Length > 0) ||
                    (myItem.OtherAddress != null &&
                        myItem.OtherAddress.Trim().Length > 0))
                {
                    return;
                }
            }

            e.Cancel = true;
        }


        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder = this.Application.ActiveExplorer().CurrentFolder;
            
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];

                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem = (selObject as Outlook.MailItem);

                        if(!VerifyEmailAddress(mailItem.SenderEmailAddress) && (CurrentItem == null || mailItem.EntryID != CurrentItem.EntryID))
                        {
                            MessageBox.Show("This sender has been marked as dangerous by another user.\n\rDo not follow any links or open any attachments.", "Dangerous Sender Detected", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                        }
                        CurrentItem = mailItem;
                }
            }
           
        }

        #region

        private bool VerifyEmailAddress(string email)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(HOST + "check-phish?email=" + email);
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
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return true;
            }
        }
        #endregion

        private void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector inspector)
        {
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;

            if(mailItem != null)
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

                    string fullName = mailItem.SenderName;

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(HOST + "?name=" + fullName.Replace(" ", "%20") + "&email=" + senderEmail);
                    request.AllowAutoRedirect = true;

                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                    Stream receiveStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(receiveStream, Encoding.UTF8);

                    string results = reader.ReadToEnd();

                    MessageBox.Show(results);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
