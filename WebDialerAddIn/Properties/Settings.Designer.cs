﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Ianchi.WebDialerAddIn.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Configuration.SettingsProviderAttribute(typeof(Ianchi.RegistrySettingsProvider.RegistrySettingsProvider))]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public global::Ianchi.WebDialerAddIn.PhoneList PhoneList {
            get {
                return ((global::Ianchi.WebDialerAddIn.PhoneList)(this["PhoneList"]));
            }
            set {
                this["PhoneList"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Configuration.SettingsProviderAttribute(typeof(Ianchi.RegistrySettingsProvider.RegistrySettingsProvider))]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("-1")]
        public int SelectedPhone {
            get {
                return ((int)(this["SelectedPhone"]));
            }
            set {
                this["SelectedPhone"] = value;
            }
        }
    }
}
