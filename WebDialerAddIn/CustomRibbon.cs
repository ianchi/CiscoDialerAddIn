/*****************************************************************************
Copyright (c) 2015 Adrian Panella

The MIT License (MIT)

Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
 all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 THE SOFTWARE.
*****************************************************************************/

//------------------------------------------------------------------------------
// 
//    Cisco IP Phone Dialer
//    Outlook UI customization classes 
//
//    Copyright (c) 2015 Adrian Panella 
// 
//------------------------------------------------------------------------------


#region Using Directives

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

#endregion


namespace Ianchi.WebDialerAddIn
{
    [ComVisible(true)]
    public class CustomRibbon : Office.IRibbonExtensibility
    {
        #region Private Variables

        private Office.IRibbonUI ribbon;
        private string lastNumber="";

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            //customization for ContactCard and Contact
            if (ribbonID == "Microsoft.Mso.IMLayerUI" || 
                ribbonID == "Microsoft.Outlook.Contact")
                return GetResourceText("Ianchi.WebDialerAddIn.CustomRibbon.xml");
            
            //customization for Explorer
            else if(ribbonID == "Microsoft.Outlook.Explorer")
                return GetResourceText("Ianchi.WebDialerAddIn.CustomRibbonExplorer.xml");
            
            else return String.Empty;
        }

        #endregion

        #region Ribbon Callbacks
 
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public bool getEnabledCalls(Office.IRibbonControl control) {
            return Globals.WebDialerAddIn.isChecked;
        }

        public bool getEnabledCallsExplorer(Office.IRibbonControl control)
        {
            return Globals.WebDialerAddIn.isChecked && control.Context.Selection.Count == 1;
        }

        public void dialCallback(Office.IRibbonControl control)
        {
            try
            {
                lastNumber = control.Tag;
                ribbon.InvalidateControl("Ianchi_Dialer_txtPhoneNumber");
                Globals.WebDialerAddIn.call(control.Tag);
            }
            catch { }
        }

        public void Ianchi_Dialer_SendKey(Office.IRibbonControl control) {
            try
            {
                if (control.Tag.Length > 0)
                    Globals.WebDialerAddIn.sendKey(control.Tag);
            }
            catch { }

        }

        public string GetLastNumber(Office.IRibbonControl control)
        {
            return lastNumber;
        }

        public string getPhonesMenuExplorer(Office.IRibbonControl control)
        {
            string xml = "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">";
            try
            {
                var contactOutlook = control.Context.Selection(1);
                if (contactOutlook != null)
                    xml += getButtonsXml(contactOutlook, outlookPhoneProperties, "OUT");
            }
            catch { }

            xml += "</menu>";

            return xml;
        }

        public string getPhonesMenuContact(Office.IRibbonControl control)
        {
            string xml = "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">";
            try
            {
                Outlook.ContactItem contactOutlook = (Outlook.ContactItem)control.Context.CurrentItem;
                if (contactOutlook != null)
                    xml += getButtonsXml(contactOutlook, outlookPhoneProperties, "OUT");
            }
            catch { }

            xml += "</menu>";

            return xml;
        }


        public string getPhonesMenuContactCard(Office.IRibbonControl control)
        {
            string xml = "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">";
            try
            {
                Office.IMsoContactCard card = (Office.IMsoContactCard)control.Context;

                if (card.AddressType == Office.MsoContactCardAddressType.msoContactCardAddressTypeOutlook)
                {

                    Outlook.ContactItem contactOutlook = null;
                    Outlook.ExchangeUser contactExchange = null;

                    Outlook.AddressEntry ae = card.Application.Session.GetAddressEntryFromID(card.Address);

                    //contact in CAB
                    contactOutlook = ae.GetContact();

                    //contact in GAL
                    if (contactOutlook == null) {
                        contactExchange = ae.GetExchangeUser();
                        if (contactExchange != null)
                        {   //see if it also local
                            Outlook.MAPIFolder contactFolder = card.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                            contactOutlook = (Outlook.ContactItem)contactFolder.Items.Find(String.Format(
                                                "[Email1Address]='{0}' or [Email2Address]='{0}' or [Email3Address]='{0}'",
                                                contactExchange.PrimarySmtpAddress));
                        }
                    }

                    if (contactOutlook != null)
                    {
                        xml += "<menuSeparator id=\"Inachi_Dialer_separator_OUT\" title=\"Outlook\" />";
                        xml += getButtonsXml(contactOutlook, outlookPhoneProperties,"OUT");
                    }
                    if (contactExchange != null)
                    {
                        xml += "<menuSeparator id=\"Inachi_Dialer_separator_EX\" title=\"Exchange\" />";
                        xml += getButtonsXml(contactExchange, exchangePhoneProperties, "EX");
                    }
                }
            }
            catch { }
            
            xml += "</menu>";
            return xml;
        }

        public void OnPhoneChange(Office.IRibbonControl control, string number)
        {
            try { Globals.WebDialerAddIn.call(number); }
            catch {}
            
        }

        public void DialogLauncherClick(Office.IRibbonControl control)
        {
            
            frmDialerOptions dlg = new frmDialerOptions();
            dlg.webDialerOptions.setAppPhone(Globals.WebDialerAddIn.phone);

            dlg.ShowDialog();

        }

        #endregion

        #region Helpers

        /// <summary>
        /// Returns an xml formated string of buttons with the active phones stored in the contact
        /// to add to a dynamicMenu
        /// </summary>
        private static string getButtonsXml(object contact, PhoneList phoneProperties, string sufix)
        {
            string xml = "";

            string number;
            foreach (var phone in phoneProperties)
            {
                number = (string)contact.GetType().InvokeMember(phone.Value.Item1, BindingFlags.GetProperty, null, contact, null);

                if (number != null && number.Length > 0)
                {
                    xml += String.Format("<button id=\"Ianchi_Dialer_{0}_{4}\" label=\"Call {2} {3}\" {1} tag=\"{3}\" onAction=\"dialCallback\" />",
                        phone.Value.Item1, phone.Value.Item2 == null ? "" : String.Format("imageMso=\"{0}\"", phone.Value.Item2), phone.Value.Item3, number, sufix);
                }
            }
            return xml;
        }

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

        #region List of Outlook Phone Properties
        PhoneList outlookPhoneProperties = new PhoneList {
                        {"AssistantTelephoneNumber", "ContactCardCallWork", "Assistant"},
                        {"Business2TelephoneNumber", "ContactCardCallWork", "Work 2"},
                        {"BusinessFaxNumber", null, "Business Fax"},
                        {"BusinessTelephoneNumber", "ContactCardCallWork", "Work"},
                        {"CallbackTelephoneNumber", null, "Callback"},
                        {"CarTelephoneNumber", "ContactCardCallMobile", "Car"},
                        {"CompanyMainTelephoneNumber", "ContactCardCallWork", "Company"},
                        {"Home2TelephoneNumber", "ContactCardCallHome", "Home 2"},
                        {"HomeFaxNumber", null, "Home Fax"},
                        {"HomeTelephoneNumber", "ContactCardCallHome", "Home"},
                        {"ISDNNumber", null, "ISDN"},
                        {"MobileTelephoneNumber", "ContactCardCallMobile", "Mobile"},
                        {"OtherFaxNumber", null, "Other Fax"},
                        {"OtherTelephoneNumber", null, "Other"},
                        {"PagerNumber", "ContactCardCallMobile", "Pager"},
                        {"PrimaryTelephoneNumber", "ContactCardCallHome", "Primary"},
                        {"RadioTelephoneNumber", "ContactCardCallMobile", "Radio"},
                        {"TelexNumber", null, "Telex"},
                        {"TTYTDDTelephoneNumber", null, "TTY/TDD"}};

        PhoneList exchangePhoneProperties = new PhoneList {
                            {"BusinessTelephoneNumber", "ContactCardCallWork", "Work"},
                            {"MobileTelephoneNumber", "ContactCardCallMobile", "Mobile"}};

        #endregion


    }


    /// <summary>
    /// Auxiliary class to allow {} initialization of List of Phone Properties
    /// In the form of PropertyName, officeImageID, Description
    /// </summary>

    public class PhoneList : SortedList<string,Tuple<string, string, string>>
    {
        public void Add(string item1, string item2, string item3)
        {
            Add(item3, new Tuple<string, string, string>(item1, item2, item3));
        }
    }

}
