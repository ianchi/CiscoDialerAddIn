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
//    Cisco IP Phone AddIn for Outlook 
//     
//
//    Copyright (c) 2015 Adrian Panella 
// 
//------------------------------------------------------------------------------

#region Using Directives

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Security.Permissions;
using System.Runtime.Serialization;


#endregion

namespace Ianchi.WebDialerAddIn
{
    /// <summary>
    /// Class representing the Add In
    /// </summary>
    public partial class WebDialerAddIn
    {
        #region Private Variables

        WebDialerOptions optionsPane;
        PhoneList _phoneList;


        #endregion

        #region Phone API
        
        public void call(string number) { 
            if(_phoneList.SelectedPhone != null)
                _phoneList.SelectedPhone.call(number);
        }
        public void sendKey(string key) {
            if (_phoneList.SelectedPhone != null)
                _phoneList.SelectedPhone.sendCommand("Key:" + key);
        }
        public bool isChecked { get { return _phoneList.SelectedPhone != null && _phoneList.SelectedPhone.isChecked; } }
        public CiscoPhone phone { get { return _phoneList.SelectedPhone; } }
        public PhoneList phoneList { 
            get { return _phoneList; } 
            set { 
                _phoneList = value; 
                Properties.Settings.Default.PhoneList = value;
                Properties.Settings.Default.SelectedPhone = value.SelectedIndex;
            } }

        

        #endregion

        /// <summary>
        /// Initialization of the add in
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            //  Add Options Pane
            Globals.WebDialerAddIn.Application.OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_OptionsPagesAdd);      

            // initializes phone from user's settings, if none adds empty list
            _phoneList = Properties.Settings.Default.PhoneList;

            if (_phoneList == null)
            {
                _phoneList=new PhoneList();
                Properties.Settings.Default.PhoneList = _phoneList;
            }


            _phoneList.SelectedIndex = Properties.Settings.Default.SelectedPhone >= _phoneList.Count ? -1 : Properties.Settings.Default.SelectedPhone;

            if (_phoneList.Count > 0 && _phoneList.SelectedIndex < 0) _phoneList.SelectedIndex = 0;

            if(_phoneList.SelectedPhone != null)
                _phoneList.SelectedPhone.checkConnection(false);


        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }


        /// <summary>
        /// Even handler, called when Outlook is creating de AddIns Options Dialog
        /// </summary>
        void Application_OptionsPagesAdd(Outlook.PropertyPages Pages)
        {
            //creates Options Page, with referenc to phone object, for storing final preferences
            optionsPane=new WebDialerOptions();
            Pages.Add(optionsPane, "");
        }

        /// <summary>
        /// Loads de Custom UI 
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CustomRibbon();
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

    [XmlRoot("PhoneList")]
    public class PhoneList : BindingList<CiscoPhone> 
    {
        int _selectedIndex =-1;
        

        public PhoneList() { }

        public PhoneList(PhoneList from)
            : base(from)
        {
            this.SelectedIndex = from.SelectedIndex >= this.Count ? -1 : from.SelectedIndex;
        }

        public int SelectedIndex { 
            get {return _selectedIndex >= Count ? -1 : _selectedIndex;}
            set { 
                if(value<0) {
                    _selectedIndex=-1;
                }
                else
                {
                    _selectedIndex = value;
                }
            } }

        public CiscoPhone SelectedPhone { get { return _selectedIndex >= Count || _selectedIndex<0 ? null : this[_selectedIndex]; } }

    }

}
