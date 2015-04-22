using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Win32;

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

namespace OutlookPlugin
{
    #region Win32 import and definitions
    
    public class Win32
    {
        [DllImport("user32.dll")]
        public static extern int FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        public static extern bool PostMessage(int hWnd, uint Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern int SendMessage(int hWnd, uint Msg, int wParam, ref COPYDATASTRUCT lParam);

        [DllImport("shell32. dll")]
        private static extern long ShellExecute(Int32 hWnd, string lpOperation, string lpFile, string lpParameters, string lpDirectory, long nShowCmd);

        public struct COPYDATASTRUCT
        {
            public IntPtr dwData;
            public int cbData;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpData;            
        }
    }

    #endregion

    [ComVisible(true)]
    public class OutlookPlugin : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private Outlook.Application olApplication;

        String appName = "OutCALL";//"BSOC Main Application Window";
        String appBinary = "OutCALL.exe";
        #region Windows Messages

        public const int WM_USER =                          0x0400;
        public const int WM_COPYDATA =                      0x004A;

        #endregion

        //Override of constructor to pass 
        // a trusted Outlook.Application object
        public OutlookPlugin(Outlook.Application outlookApplication)
        {
            olApplication = outlookApplication as Outlook.Application;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string customUI = string.Empty;

            //RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\BicomSystems\\outcall");

            //Return the appropriate Ribbon XML for ribbonID
            //if ((string)key.GetValue("licensed") == "1")
            //{
            customUI = GetResourceText("OutlookPlugin.UIManifest.xml");
            //}
            return customUI;
        }

        #endregion

        #region Ribbon Callbacks
        // RibbonX callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ThisAddIn.m_Ribbon = ribbonUI;
        }

        // Only show MyTab when Explorer Selection is 
        // a received mail or when Inspector is a read note
        public bool MyTab_GetVisible(Office.IRibbonControl control)
        {
            if (control.Context is Outlook.Explorer)
            {
                Outlook.Explorer explorer = control.Context as Outlook.Explorer;
                Outlook.Selection selection = explorer.Selection;
                if (selection.Count == 1)
                {
                    if (selection[1] is Outlook.MailItem)
                    {
                        Outlook.MailItem oMail = selection[1] as Outlook.MailItem;
                        if (oMail.Sent == true)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else if (control.Context is Outlook.Inspector)
            {
                Outlook.Inspector oInsp = control.Context as Outlook.Inspector;
                if (oInsp.CurrentItem is Outlook.MailItem)
                {
                    Outlook.MailItem oMail = oInsp.CurrentItem as Outlook.MailItem;
                    if (oMail.Sent == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        public bool MyTabInspector_GetVisible(Office.IRibbonControl control)
        {
            if (control.Context is Outlook.Inspector)
            {
                Outlook.Inspector oInsp = control.Context as Outlook.Inspector;
                if (oInsp.CurrentItem is Outlook.MailItem)
                {
                    Outlook.MailItem oMail = oInsp.CurrentItem as Outlook.MailItem;
                    if (oMail.Sent == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        //MyBackstageTab_GetVisible hides the place in an Inspector window
        public bool MyBackstageTab_GetVisible(Office.IRibbonControl control)
        {
            if (control.Context is Outlook.Explorer)
                return true;
            else
                return false;
        }

        public stdole.IPictureDisp GetCurrentUserImage(Office.IRibbonControl control)
        {
            //stdole.IPictureDisp pictureDisp = null;
            Outlook.AddressEntry addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
                if(addrEntry.Type=="EX")
                {
                    if (Globals.ThisAddIn.m_pictdisp != null)
                    {
                        return Globals.ThisAddIn.m_pictdisp;
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
        }

        public stdole.IPictureDisp GetImage(string imageName)
        {
            return PictureConverter.IconToPictureDisp(Properties.Resources.Icon);
        }
        
        public string GetLabel(Office.IRibbonControl control)
        {
            //RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\BicomSystems\\outcall");
            //return (string)key.GetValue("app_name");
            return appName;
        }

        // OnMyButtonClick routine handles all button click events
        // and displays IRibbonControl.Context in message box
        //public void OnMyButtonClick(Office.IRibbonControl control) 

        public void ButtonClicked(Office.IRibbonControl control)    
        {
            String ContactName = "";
            String numbers = "";
            Outlook.ContactItem contactItem = null;

            //int handle = Win32.FindWindow(null, appName);

            if (true) // handle != 0
            {
				try
				{
					if (control.Id == "outcallContacts")
					{
						Outlook.Explorer explorer = control.Context as Outlook.Explorer;
						Object item = explorer.Selection[1];
						contactItem = item as Outlook.ContactItem;
					}
					else if (control.Id == "outcallMenuContact")
					{
						Outlook.Selection selection = control.Context as Outlook.Selection;
						OutlookItem olItem = new OutlookItem(selection[1]);
						Object item = new Object();
						item = selection[1];
						contactItem = item as Outlook.ContactItem;
					}
					else if (control.Id == "outcallMenuMail" || control.Id == "outcallMail")
					{
						Outlook.MailItem mailItem = null;
						Object item = null;

                        List<string> accounts = new List<string>();

						if (control.Id == "outcallMail")
						{
							Outlook.Explorer explorer = control.Context as Outlook.Explorer;
							item = explorer.Selection[1];
						}
						else if (control.Id == "outcallMenuMail")
						{
							Outlook.Selection selection = control.Context as Outlook.Selection;
							OutlookItem olItem = new OutlookItem(selection[1]);
							item = new Object();
							item = selection[1];
						}

						mailItem = item as Outlook.MailItem;
						Outlook.Folder folder = mailItem.Parent as Outlook.Folder;
                        Outlook.Folder contacts = (Outlook.Folder)folder.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

                        if (mailItem != null)
                        {
                            foreach (Outlook.Account account in folder.Session.Accounts)
                            {
                                accounts.Add(account.SmtpAddress);
                            }

                            bool isOut = false;

                            if (mailItem.Sender == null)
                            {
                                isOut = true;
                            }
                            else if (accounts.Contains(mailItem.Sender.Address))
                            {
                                isOut = true;
                            }

                            if (isOut)
                            {
                                Outlook.Recipients recipients = mailItem.Recipients;
                                foreach (Outlook.Recipient recipient in recipients)
                                {
                                    contactItem = recipient.AddressEntry.GetContact();

                                    if (contactItem != null)
                                        break;
                                }
                            }
                            else
                            {
                                contactItem = mailItem.Sender.GetContact();
                            }
                        }
					}
					else
					{
                        MessageBox.Show("The selected item is not supported. Please select a Contact or an Email.");
					}
				} catch (Exception e) {
                    MessageBox.Show("Please select a contact.");
					return;
				}

				if (contactItem != null)
				{
					ContactName = contactItem.FullName;
					/*if (ContactName != "")
					{
						ContactName = contactItem.FirstName + " | " + contactItem.MiddleName + " | " + contactItem.LastName;
					}
					else
					{
						ContactName = contactItem.CompanyName;
					}*/

					String strTemp = contactItem.AssistantTelephoneNumber;
					if (strTemp != null)
						numbers += "Assistant:" + strTemp + " | ";

					strTemp = contactItem.BusinessTelephoneNumber;
					if (strTemp != null)
						numbers += "Business:" + strTemp + " | ";

					strTemp = contactItem.Business2TelephoneNumber;
					if (strTemp != null)
						numbers += "Business2:" + strTemp + " | ";

					strTemp = contactItem.BusinessFaxNumber;
					if (strTemp != null)
						numbers += "Business Fax:" + strTemp + " | ";

					strTemp = contactItem.CallbackTelephoneNumber;
					if (strTemp != null)
						numbers += "Callback:" + strTemp + " | ";

					strTemp = contactItem.CarTelephoneNumber;
					if (strTemp != null)
						numbers += "Car:" + strTemp + " | ";

					strTemp = contactItem.CompanyMainTelephoneNumber;
					if (strTemp != null)
						numbers += "Company Main:" + strTemp + " | ";

					strTemp = contactItem.HomeTelephoneNumber;
					if (strTemp != null)
						numbers += "Home:" + strTemp + " | ";

					strTemp = contactItem.Home2TelephoneNumber;
					if (strTemp != null)
						numbers += "Home2:" + strTemp + " | ";

					strTemp = contactItem.HomeFaxNumber;
					if (strTemp != null)
						numbers += "Home Fax:" + strTemp + " | ";

					strTemp = contactItem.ISDNNumber;
					if (strTemp != null)
						numbers += "ISDN:" + strTemp + " | ";

					strTemp = contactItem.MobileTelephoneNumber;
					if (strTemp != null)
						numbers += "Mobile:" + strTemp + " | ";

					strTemp = contactItem.OtherTelephoneNumber;
					if (strTemp != null)
						numbers += "Other:" + strTemp + " | ";

					strTemp = contactItem.OtherFaxNumber;
					if (strTemp != null)
						numbers += "Other Fax:" + strTemp + " | ";

					strTemp = contactItem.PagerNumber;
					if (strTemp != null)
						numbers += "Pager:" + strTemp + " | ";

					strTemp = contactItem.PrimaryTelephoneNumber;
					if (strTemp != null)
						numbers += "Primary:" + strTemp + " | ";

					strTemp = contactItem.RadioTelephoneNumber;
					if (strTemp != null)
						numbers += "Radio:" + strTemp + " | ";

					strTemp = contactItem.TelexNumber;
					if (strTemp != null)
						numbers += "Telex:" + strTemp + " | ";

					strTemp = contactItem.TTYTDDTelephoneNumber;
					if (strTemp != null)
						numbers += "TTY/TDD:" + strTemp + " | ";
				}
				else
				{
					MessageBox.Show("Contact not found in your Address book.");
					return;
				}

				if (numbers == "")
				{
					MessageBox.Show("There are no phone numbers for the selected contact.");
					return;
				}

                String strDataToSend = "Dial#####" + ContactName + "#####" + numbers;

                strDataToSend = strDataToSend.Replace(" ", "&&&");

                /*Win32.COPYDATASTRUCT cpd;
                cpd.dwData = IntPtr.Zero;
                cpd.cbData = strDataToSend.Length;
                cpd.lpData = strDataToSend;

                Win32.SendMessage(handle, WM_COPYDATA, 0, ref cpd);*/

                RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Bicom Systems\\OutCALL");

                string outcall_path = (string)key.GetValue("InstallDir");

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = false;
                startInfo.FileName = outcall_path + "\\" + appBinary;
                startInfo.WindowStyle = ProcessWindowStyle.Normal;
                startInfo.Arguments = strDataToSend;

                try
                {
                    // Start the process with the info we specified.
                    // Call WaitForExit and then the using statement will close.
                    using (Process outcallProcess = Process.Start(startInfo))
                    {
                        outcallProcess.WaitForExit();
                    }
                }
                catch
                {
                    // Log error.

                    Debug.WriteLine("outcall execute error");
                }
            }
            else
            {
                /*RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\BicomSystems\\outcall");
                String applicationName = (string)key.GetValue("app_name");
                MessageBox.Show(applicationName + " is not running.");*/
                MessageBox.Show(appName + " is not running.");
            }
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

        private Outlook.ContactItem FindContactItem(String contactName, Outlook.Folder folder)
        {
            //object missing = System.Reflection.Missing.Value;
            foreach (Outlook.ContactItem contactItem in folder.Items)
            {
                if (contactItem.FullName==contactName)
                {
                    return contactItem;
                }
            }
            return null;
        }
        
        #endregion
    }

    #region PictureConverter Class
    internal class PictureConverter : AxHost
    {
        private PictureConverter() : base(String.Empty) { }

        static public stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }

        static public stdole.IPictureDisp IconToPictureDisp(Icon icon)
        {
            return ImageToPictureDisp(icon.ToBitmap());
        }

        static public Image PictureDispToImage(stdole.IPictureDisp picture)
        {
            return GetPictureFromIPicture(picture);
        }
    }

    #endregion

}
