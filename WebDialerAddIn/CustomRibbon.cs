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

        #region General

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Globals.CustomRibbon = this;
        }

        public void Invalidate()
        {
            if (ribbon != null)
                ribbon.Invalidate();
        }

        #endregion

        #region Context Menu callbacks

        public bool getEnabledCalls(Office.IRibbonControl control) {
            return Globals.WebDialerAddIn.isChecked;
        }

        public bool getEnabledCallsExplorer(Office.IRibbonControl control)
        {
            return Globals.WebDialerAddIn.isChecked && control.Context.Selection.Count == 1;
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
                                                "([Email1AddressType]='SMTP' and [Email1Address]='{0}') or " +
                                                "([Email2AddressType]='SMTP' and [Email2Address]='{0}') or " +
                                                "([Email3AddressType]='SMTP' and [Email3Address]='{0}') or " +
                                                "([Email1AddressType]='EX' and [Email1Address]='{1}') or " +
                                                "([Email2AddressType]='EX' and [Email2Address]='{1}') or " +
                                                "([Email3AddressType]='EX' and [Email3Address]='{1}')",
                                                contactExchange.PrimarySmtpAddress, contactExchange.Address));
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
                        xml += getButtonsXml(contactExchange, mapiPhoneProperties, "MAPI");
                    }
                }
            }
            catch { }
            
            xml += "</menu>";
            return xml;
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
        
        #endregion

        #region Main Ribbon Callbacks

        public void OnPhoneChange(Office.IRibbonControl control, string number)
        {
            try {
                lastNumber = number;
                Globals.WebDialerAddIn.call(number); }
            catch { }

        }

        public void Ianchi_Dialer_SendKey(Office.IRibbonControl control)
        {
            try
            {
                if (control.Tag == "Hook")
                {
                    Globals.WebDialerAddIn.sendKey(Globals.WebDialerAddIn.phoneList.SelectedPhone.hook);
                }
                else if (control.Tag.Length > 0) { 
                    Globals.WebDialerAddIn.sendKey(control.Tag); }
            }
            catch { }

        }

        public void DialogLauncherClick(Office.IRibbonControl control)
        {
            
            frmDialerOptions dlg = new frmDialerOptions();

            dlg.ShowDialog();

        }

        public string GetLastNumber(Office.IRibbonControl control)
        {
            return lastNumber;
        }

        #endregion

        #region Profile Selector Callbacks 

        public int GetProfileCount(Office.IRibbonControl control)
        {
            return Globals.WebDialerAddIn.phoneList.Count;
        }

        public int GetSelectedProfileIndex(Office.IRibbonControl control)
        {
            return Globals.WebDialerAddIn.phoneList.SelectedIndex;
        }

        public string GetProfileLabel(Office.IRibbonControl control, int index)
        {
            return Globals.WebDialerAddIn.phoneList[index].description;
        }

        public string GetProfileID(Office.IRibbonControl control, int index)
        {
            return String.Format("Ianchi_Dialer_Profile_{0}", index);
        }

        public void OnSelectProfile(Office.IRibbonControl control,string selectedId, int selectedIndex)
        {
            Globals.WebDialerAddIn.phoneList.SelectedIndex = selectedIndex;
            Globals.WebDialerAddIn.phoneList.SelectedPhone.checkConnection(false);
            ribbon.Invalidate();

        }

        public bool getEnabledProfile(Office.IRibbonControl control) {
            return Globals.WebDialerAddIn.phoneList.Count > 1;
        }
       

        public string GetProfileSupertip(Office.IRibbonControl control)
        {
            return "Current profile: " + (Globals.WebDialerAddIn.phoneList.SelectedPhone != null ? Globals.WebDialerAddIn.phoneList.SelectedPhone.description : "");
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Returns an xml formated string of buttons with the active phones stored in the contact
        /// to add to a dynamicMenu
        /// </summary>
        private static string getButtonsXml(object contact, PhoneProperties phoneProperties, string sufix)
        {
            string xml = "";

            string number;
            foreach (var phone in phoneProperties)
            {
                if (sufix == "MAPI")
                {
                    try
                    {
                        number = ((Outlook.ExchangeUser)contact).PropertyAccessor.GetProperty(String.Format("http://schemas.microsoft.com/mapi/proptag/0x{0}", phone.Value.Item1));
                    } catch {
                        number = null;
                    }
                }
                else
                {
                    number = (string)contact.GetType().InvokeMember(phone.Value.Item1, BindingFlags.GetProperty, null, contact, null);
                }
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
        PhoneProperties outlookPhoneProperties = new PhoneProperties {
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

        PhoneProperties exchangePhoneProperties = new PhoneProperties {
                            {"BusinessTelephoneNumber", "ContactCardCallWork", "Work"},
                            {"MobileTelephoneNumber", "ContactCardCallMobile", "Mobile"}};

        PhoneProperties mapiPhoneProperties = new PhoneProperties {
                        {"3A2E001E", "ContactCardCallWork", "Assistant"},
                        {"3A1B001E", "ContactCardCallWork", "Work 2"},
                        {"3A24001E", null, "Business Fax"},
                        {"3A08001E", "ContactCardCallWork", "Work"},
                        {"3A02001E", null, "Callback"},
                        {"3A1E001E", "ContactCardCallMobile", "Car"},
                        {"3A57001E", "ContactCardCallWork", "Company"},
                        {"3A2F001E", "ContactCardCallHome", "Home 2"},
                        {"3A25001E", null, "Home Fax"},
                        {"3A09001E", "ContactCardCallHome", "Home"},
                        {"3A2D001E", null, "ISDN"},
                        {"3A1C001E", "ContactCardCallMobile", "Mobile"},
                        {"3A23001E", null, "Other Fax"},
                        {"3A1F001E", null, "Other"},
                        {"3A21001E", "ContactCardCallMobile", "Pager"},
                        {"3A1A001E", "ContactCardCallHome", "Primary"},
                        {"3A1D001E", "ContactCardCallMobile", "Radio"},
                        {"3A2C001E", null, "Telex"},
                        {"3A4B001E", null, "TTY/TDD"}};

        #endregion


    }


    /// <summary>
    /// Extends Globals to have easy access to customribbon
    /// </summary>
    internal sealed partial class Globals
    {
        public static CustomRibbon CustomRibbon;

    }

    /// <summary>
    /// Auxiliary class to allow {} initialization of List of Phone Properties
    /// In the form of PropertyName, officeImageID, Description
    /// </summary>

    internal class PhoneProperties : SortedList<string,Tuple<string, string, string>>
    {
        public void Add(string item1, string item2, string item3)
        {
            Add(item3, new Tuple<string, string, string>(item1, item2, item3));
        }
    }

}
