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
//    Cisco IP Phone wrapper classes 
//     
//
//    Copyright (c) 2015 Adrian Panella 
// 
//------------------------------------------------------------------------------


#region Using Directives

using System;
using System.Web;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Security.Permissions;
using System.Security.Cryptography;
using System.Runtime.Serialization;


#endregion


namespace Ianchi.WebDialerAddIn
{
    /// <summary>
    /// Wrapper class to a Cisco IP Phone.
    /// Provides methos for sending commands directly to the phone.
    /// </summary>
    public class CiscoPhone : IXmlSerializable
    {
        #region Private Variables
        static byte[] _salt = Encoding.Unicode.GetBytes("10635AF2-96CC-40C4-8FE7-4E0A9D5C2800");

        string _phoneIP;    // IP address of the phone. (you can get it from de phone in Settings - Phone Information)
        string _user;       // user name to login to the phone, same used to enter the CM user web   
        string _password;   // password to login to the phone, same used to enter the CM user web
        string _description;

        [NonSerialized]
        bool _checked = false;// true if the above parameters have been checked to work correctly with the phone

        #endregion

        #region Public Properties

        public string phoneIP { get { return _phoneIP; } set { _phoneIP = value; _checked=false; } }
        public string user { get { return _user; } set { _user = value; _checked = false; } }
        public string password { get { return _password; } set { _password = value; _checked = false; } }
        public string description { get { return _description; } set { _description = value; } }

        public bool isChecked { get { return _checked; } }

        public enum executePriority { inmmediately = 0, whenIdle, ifIdle}

        #endregion

        #region Public Methods

        public CiscoPhone() { }

        public CiscoPhone clone()
        {
            return (CiscoPhone)this.MemberwiseClone();
        }

        /// <summary>
        /// Dials a number on the phone
        /// </summary>
        /// <param name="num">String with the number to dial. Any non valid character will be discarded before sending to the phone.
        /// Valid characters 0-9+#*,;x  (the x is the extension sepparator in outlook contacts
        /// The "," insterts a pause, though in some models all numbers after the main number are ignored
        /// </param>
        /// <exception cref="CiscoException">If the phone returned a response with status not 0 (success)</exception>
        /// <exception > Any exception other exception on the HttpWebRequest </exception>
        public void call(string num)
        {
            //delete all non valid characters
            num = Regex.Replace(num, "[^0-9+*#,;x]*", String.Empty);

            //split the string on the first pause [,;].
            var m = Regex.Match(num, "^([0-9+*#]*)(.*)$");


            //first part is sent as a DIAL command
            
            if (m.Groups[1].Length > 0)
                sendCommand("Dial:" + m.Groups[1].Value);

            //the rest as individual keys, with [,;] converted to delays.
            if (m.Groups.Count > 2 && m.Groups[2].Length > 0)
            {
                System.Threading.Thread.Sleep(6000); //allows for the call to be connected before sending keys
                sendKeys(m.Groups[2].Value);
            }

        }

        /// <summary>
        /// Check if the parameters of the connection are working
        /// </summary>
        /// <param name="except">If true exceptions are thrown on any error, otherwise status is returned in the return value
        /// </param>
        /// <returns>True if connection is successful false on any error</returns>
        /// <exception > Any exception on the HttpWebRequest if except==true</exception>
        public bool checkConnection(bool except = true)
        {

            try
            {
                if (_phoneIP.Length == 0 || _user.Length == 0 || _password.Length == 0)
                    throw new Exception("Parameters not initialized");

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(
                            String.Format("http://{0}/CGI/Screenshot", _phoneIP));

                request.Timeout = 1000; //millisenconds to wait
                request.Method = "GET";
                request.Accept = "*/*";
                request.ContentType = "application/x-www-form-urlencoded";
                request.Credentials = new NetworkCredential(_user, _password);
                request.PreAuthenticate = true;

                WebResponse response = request.GetResponse();

                response.Close();
            }
            catch
            {
                if (except) throw;
                return false;
            }

            _checked = true;
            return true;
        }

        /// <summary>
        /// Navigates the browser object to the Phone Screenshot page, setting appropiate headers for username/password
        /// </summary>
        /// <param name="browser">The WebBrowser object to navigate
        /// </param>
        /// <exception > Any exception on the HttpWebRequest </exception>
        public void navigateScreenShot(System.Windows.Forms.WebBrowser browser)
        {

            if (browser != null && checkConnection())
                browser.Navigate(String.Format("http://{0}/CGI/Screenshot", _phoneIP), null, null,
                    "Authorization: Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(_user + ":" + _password)) + "\r\n");
        }


        /// <summary>
        /// Sends keys to the phone, to complete dialing
        /// </summary>
        /// <param name="num">String with the numbers to dial. Any non valid character will be discarded before sending to the phone.
        /// Valid characters 0-9#*,
        /// The "," insterts a 2secs pause and ";" a 10secs pause
        /// </param>
        /// <exception cref="CiscoException">If the phone returned a response with status not 0 (success)</exception>
        /// <exception > Any exception other exception on the HttpWebRequest </exception>
        public void sendKeys(string numbers) {
            numbers = Regex.Replace(numbers, "[^0-9*#,;]*", String.Empty);
            string key;

            foreach (var k in numbers)
            {
                switch (k)
                {
                    case '#':
                        key = "KeyPadPound";
                        break;
                    case '*':
                        key = "KeyPadStar";
                        break;

                    case ',':
                        System.Threading.Thread.Sleep(2000);
                        continue;
                    case ';':
                        System.Threading.Thread.Sleep(10000);
                        continue;
                        
                    default:
                        key="KeyPad" + k;
                        break;
                }
                sendCommand("Key:" + key, executePriority.whenIdle);

            }
 

        }

        
        /// <summary>
        /// Sends an URI command to the phone
        /// </summary>
        /// <param name="command">Command to send to the phone formated as a valid URI for CiscoIPPhoneExecute.ExecuteItem
        /// Example: DIAL:123456 </param>
        /// <exception cref="CiscoException">If the phone returned a response with status not 0 (success)</exception>
        /// <exception > Any exception other exception on the HttpWebRequest </exception>
        public void sendCommand(string command, executePriority priority=executePriority.ifIdle)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(
                                        String.Format("http://{0}/CGI/Execute", _phoneIP));

            request.Timeout = 30 * 1000;
            request.Method = "POST";
            request.Accept = "*/*";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Credentials = new NetworkCredential(_user, _password);
            request.PreAuthenticate = true;

            Byte[] bytes = Encoding.UTF8.GetBytes("XML=" + HttpUtility.UrlEncode(String.Format(
                "<CiscoIPPhoneExecute><ExecuteItem Priority=\"{1}\" URL=\"{0}\" /></CiscoIPPhoneExecute>", command, priority)));

            Stream outStream = request.GetRequestStream();
            outStream.Write(bytes, 0, bytes.Length);
            outStream.Close();

            WebResponse response = request.GetResponse();

            var nodes = XElement.Load(response.GetResponseStream());

            if (nodes.Element("ResponseItem").Attribute("Status").Value != "0")
                throw new CiscoException(nodes.Element("ResponseItem").Attribute("Data").Value, 
                                        Int32.Parse(nodes.Element("ResponseItem").Attribute("Status").Value));

            response.Close();


        }

        #endregion


        #region XML Serialization

        public void WriteXml(XmlWriter writer)
        {
            writer.WriteAttributeString("Name", _description);
            writer.WriteAttributeString("IP", _phoneIP);
            writer.WriteAttributeString("User", _user);
            writer.WriteElementString("Psw", Encrypt(_password));
        }

        public void ReadXml(XmlReader reader)
        {
            reader.MoveToContent();
            _description = reader.GetAttribute("Name");
            _phoneIP = reader.GetAttribute("IP");
            _user = reader.GetAttribute("User");
            Boolean isEmptyElement = reader.IsEmptyElement; 
            reader.ReadStartElement();
            if (!isEmptyElement) 
            {
                _password = Decrypt(reader.ReadElementString("Psw"));
                reader.ReadEndElement();
            }







        
        }

        public System.Xml.Schema.XmlSchema GetSchema() { return null; }



        private string Encrypt(string data)
        {
            if (data == null || data == "") return data;
            byte[] encrypted = ProtectedData.Protect(Encoding.Unicode.GetBytes(data ?? ""), _salt, DataProtectionScope.CurrentUser);

            return Convert.ToBase64String(encrypted);
        }

        private string Decrypt(string data)
        {
            if (data == null || data == "") return data;
            byte[] decrypted = ProtectedData.Unprotect(Convert.FromBase64String(data ?? ""), _salt, DataProtectionScope.CurrentUser);

            return Encoding.Unicode.GetString(decrypted);

        }


        #endregion
    }


    /// <summary>
    /// Exception class for errors returned by calls to Phone.
    /// </summary>
    /// <field name="statusCode">Status code returned in the response</field>
    /// <field name="message">Description of the error returned in the response</field>

    //[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2237:MarkISerializableTypesWithSerializable")]
    [Serializable]
    public class CiscoException : Exception {
        public int statusCode = 0;

        public CiscoException(string message) : base(message) {}

        public CiscoException(string message, int code) : base(message) { statusCode = code; }

        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);

            info.AddValue("statusCode", statusCode);
        }

        
    }
}
