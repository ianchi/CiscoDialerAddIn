using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace Ianchi.WebDialerAddIn
{
    [ComVisible(true)]
    public partial class WebDialerOptions : UserControl, Outlook.PropertyPage
    {
        private bool _isDirty;
        Outlook.PropertyPageSite _PropertyPageSite;

        private PhoneList phoneList;
        private BindingSource sourceProfile;


        public WebDialerOptions()
        {
            InitializeComponent();
            
            phoneList = new PhoneList(Globals.WebDialerAddIn.phoneList);

            if (phoneList.Count == 0)
                phoneList.Add(newPhone());
        }

        private void testConnection()
        {
            CiscoPhone phone = (CiscoPhone)sourceProfile.Current;
            if (phone == null)
            {
                txtStatus.Text = "No profile";
                webBrowser.Visible = false;
            }
            else try
                {
                    phone.checkConnection();
                    txtStatus.Text = "Connection successful";
                    phone.navigateScreenShot(webBrowser);
                    webBrowser.Visible = true;
                }
                catch (Exception x)
                {
                    txtStatus.Text = x.Message;
                    webBrowser.Visible = false;
                }

            tipStatus.SetToolTip(txtStatus, txtStatus.Text);
        }

        #region PropertyPage Interface

        public bool Dirty { 
            get { return _isDirty; } 
            set { 
                _isDirty = value;

                if (value)
                {
                    txtStatus.Text = "Connection not tested";
                    webBrowser.Visible = false;
                }
            
            } }

        public void Apply() {
            if (!Dirty) return;

            phoneList.SelectedIndex = cboProfile.SelectedIndex;
            if (phoneList.SelectedIndex >= 0)
                phoneList.SelectedPhone.checkConnection(false);

            //transfers options to applications phone
            //and persist 
            Globals.WebDialerAddIn.phoneList = new PhoneList(phoneList);
            Properties.Settings.Default.Save();
            Globals.CustomRibbon.Invalidate();
            
            
            Dirty = false;
            
            }
        
        public void GetPageInfo(ref string helpFile, ref int helpContext) { }
        
        [DispId(-518)]
        public string PageCaption { get { return "Cisco Dialer"; } }

        /// <summary>
        /// This function gets the parent PropertyPageSite Object using Reflection.
        /// Must be called in Load event.
        /// </summary>
        /// <returns>The parent PropertyPageSite Object</returns>
        Outlook.PropertyPageSite GetPropertyPageSite()
        {
            Type type = typeof(System.Object);
            string assembly = type.Assembly.CodeBase.Replace("mscorlib.dll", "System.Windows.Forms.dll");
            assembly = assembly.Replace("file:///", "");

            string assemblyName = System.Reflection.AssemblyName.GetAssemblyName(assembly).FullName;
            Type unsafeNativeMethods = Type.GetType(System.Reflection.Assembly.CreateQualifiedName(assemblyName, "System.Windows.Forms.UnsafeNativeMethods"));

            Type oleObj = unsafeNativeMethods.GetNestedType("IOleObject");
            System.Reflection.MethodInfo methodInfo = oleObj.GetMethod("GetClientSite");
            if (methodInfo == null) return null;
            object propertyPageSite = methodInfo.Invoke(this, null);

            return (Outlook.PropertyPageSite)propertyPageSite;
        }

        /// <summary>
        /// This Function sets our Control Drity or Clean and
        /// calls our Parent to check the control state.
        /// </summary>
        /// <param name="isDirty"></param>
        void OnDirty(bool status)
        {
            _isDirty = status;

            // When this Method is called, the PageSite checks for Dirty Flag of all Optionspages.
            if (_PropertyPageSite!=null)
                _PropertyPageSite.OnStatusChange();
        }

        #endregion

        #region Event Handlers

        private void txtUrl_TextChanged(object sender, EventArgs e) { Dirty = true;}

        private void txtUser_TextChanged(object sender, EventArgs e) { Dirty = true;}

        private void txtPassword_TextChanged(object sender, EventArgs e) { Dirty = true;}

        private void btnLogin_Click(object sender, EventArgs e)
        {
            
            //if no new parameters don't do anything
            if (!Dirty) return;
            UseWaitCursor = true;
            
            testConnection();

            
            UseWaitCursor = false;

            OnDirty(true);

        }

        private void WebDialerOptions_Load(object sender, EventArgs e)
        {
            _PropertyPageSite = GetPropertyPageSite();

            //Databinding
            sourceProfile = new BindingSource();
            sourceProfile.DataSource = phoneList;

            int selected=phoneList.SelectedIndex;
            cboProfile.DataSource = sourceProfile;
            cboProfile.DisplayMember = "description";
            cboProfile.SelectedIndex = selected;

            txtUrl.DataBindings.Add(new Binding("Text", sourceProfile, "phoneIP"));
            txtUser.DataBindings.Add(new Binding("Text", sourceProfile, "user"));
            txtPassword.DataBindings.Add(new Binding("Text", sourceProfile, "password"));

            testConnection();
            Dirty = false;
        }

        #endregion

        #region Profile Management Events

        private void cboProfile_SelectedIndexChanged(object sender, EventArgs e) { Dirty = true; }

        private void cboProfile_Validating(object sender, CancelEventArgs e)
        {
            if (sourceProfile.Current == null) return;
            CiscoPhone current = (CiscoPhone)sourceProfile.Current;

            //if there was a change
            if (current.description != cboProfile.Text)
            {
                current.description = cboProfile.Text;
                sourceProfile.ResetCurrentItem();
                _isDirty = true;
            }
        }


        private CiscoPhone newPhone()
        {
            CiscoPhone newPhone = new CiscoPhone();

            newPhone.phoneIP = txtUrl.Text;
            newPhone.user = txtUser.Text;
            newPhone.password = txtPassword.Text;
            newPhone.description = String.Format("New Profile N° {0}", phoneList.Count + 1);

            return newPhone;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            sourceProfile.Add(newPhone());
            cboProfile.SelectedIndex = cboProfile.Items.Count - 1;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            sourceProfile.RemoveCurrent();

            if (phoneList.Count == 0)
            {
                sourceProfile.Add(newPhone());
                cboProfile.SelectedIndex = cboProfile.Items.Count - 1;
            }

        }

        #endregion




    }
}
