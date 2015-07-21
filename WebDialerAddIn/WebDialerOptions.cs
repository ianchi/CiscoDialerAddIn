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
        private bool isDirty;
        Outlook.PropertyPageSite _PropertyPageSite;

        private CiscoPhone phone=new CiscoPhone();
        private CiscoPhone appPhone;


        public WebDialerOptions()
        {
            InitializeComponent();
        }

        public WebDialerOptions(CiscoPhone phoneApp)
        {
            InitializeComponent();
            setAppPhone(phoneApp);

        }

        public void setAppPhone(CiscoPhone phoneApp) {
            //stores reference to applications phone to return parameters on apply
            appPhone = phoneApp;

            if (phoneApp == null) return;

            phone.phoneIP = appPhone.phoneIP;
            phone.user = appPhone.user;
            phone.password = appPhone.password;
            phone.checkConnection(false);
        
        }

        //PropertyPage Interface
        public bool Dirty { get { return isDirty; }}
        public void Apply() {

            //persist options
            Properties.Settings.Default.Save();

            //transfers options to applications phone

            appPhone.phoneIP    = phone.phoneIP;
            appPhone.user       = phone.user;
            appPhone.password   = phone.password;
            appPhone.checkConnection();

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
            isDirty = status;

            // When this Method is called, the PageSite checks for Dirty Flag of all Optionspages.
            if (_PropertyPageSite!=null)
                _PropertyPageSite.OnStatusChange();
        }


        //Form Event Handlers

        private void txtUrl_TextChanged(object sender, EventArgs e)
        {
            isDirty = true;
        }

        private void txtUser_TextChanged(object sender, EventArgs e)
        {
            isDirty = true;
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
            isDirty = true;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            
            //if no new parameters don't do anything
            if (!isDirty) return;
            UseWaitCursor = true;
            
            try
            {
                phone.phoneIP=txtUrl.Text;
                phone.user=txtUser.Text;
                phone.password=txtPassword.Text;
                phone.checkConnection();
                txtStatus.Text = "Connection successful";
                phone.navigateScreenShot(webBrowser);
            }
            catch (Exception x)
            {
                txtStatus.Text = x.Message;
                webBrowser.Navigate("about:blank");
            }
            tipStatus.SetToolTip(txtStatus, txtStatus.Text);
            UseWaitCursor = false;

            OnDirty(true);

        }

        private void WebDialerOptions_Load(object sender, EventArgs e)
        {
            _PropertyPageSite = GetPropertyPageSite();

            if (phone.isChecked)
            {
                txtStatus.Text = "Connection successful";
                phone.navigateScreenShot(webBrowser);
            }
        }

   
    }
}
