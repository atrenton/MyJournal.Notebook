﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace MyJournal.Notebook.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("MyJournal.Notebook.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Written by Art Trenton.
        /// </summary>
        internal static string About_Author {
            get {
                return ResourceManager.GetString("About_Author", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Copyright © 2012-2023 Art Trenton.
        /// </summary>
        internal static string About_Copyright {
            get {
                return ResourceManager.GetString("About_Copyright", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to urn:MyJournal-Notebook-Config:PageSettings-v1.0.
        /// </summary>
        internal static string Config_PageSettings_Namespace {
            get {
                return ResourceManager.GetString("Config_PageSettings_Namespace", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;
        ///&lt;xs:schema xmlns:xs=&quot;http://www.w3.org/2001/XMLSchema&quot;
        ///  xmlns:tns=&quot;urn:MyJournal-Notebook-Config:PageSettings-v1.0&quot;
        ///  targetNamespace=&quot;urn:MyJournal-Notebook-Config:PageSettings-v1.0&quot;
        ///  elementFormDefault=&quot;qualified&quot;&gt;
        ///  &lt;!-- PageColorEnum --&gt;
        ///  &lt;xs:simpleType name=&quot;PageColorEnum&quot;&gt;
        ///    &lt;xs:annotation&gt;
        ///      &lt;xs:documentation xml:lang=&quot;en&quot;&gt;Page body color&lt;/xs:documentation&gt;
        ///    &lt;/xs:annotation&gt;
        ///    &lt;xs:restriction base=&quot;xs:token&quot;&gt;
        ///      &lt;xs:enumeration value= [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string Config_PageSettings_v1_0_xsd {
            get {
                return ResourceManager.GetString("Config_PageSettings_v1_0_xsd", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;?xml version=&quot;1.0&quot; encoding=&quot;utf-8&quot;?&gt;
        ///&lt;customUI xmlns=&quot;http://schemas.microsoft.com/office/2009/07/customui&quot;
        ///  xmlns:addIn=&quot;urn:MyJournal-Notebook-UI:CustomUI-v16.0&quot;
        ///  onLoad=&quot;NotifyRibbonLoad&quot; loadImage=&quot;LoadImage&quot;&gt;
        ///&lt;ribbon&gt;
        ///  &lt;tabs&gt;
        ///    &lt;tab idMso=&quot;TabHome&quot;&gt;
        ///      &lt;group idQ=&quot;addIn:MyJournal&quot; label=&quot;Add-In&quot; insertAfterMso=&quot;GroupTagging&quot;&gt;
        ///        &lt;splitButton id=&quot;SplitButton&quot; size=&quot;large&quot; keytip=&quot;J&quot;&gt;
        ///          &lt;button id=&quot;NotebookButton&quot; getLabel=&quot;GetAboutLabel&quot;
        ///           imageMso=&quot;OpenNotebook [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string UI_CustomUI_v16_0_xml {
            get {
                return ResourceManager.GetString("UI_CustomUI_v16_0_xml", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Click here to open your Journal notebook to today&apos;s page..
        /// </summary>
        internal static string UI_NotebookButton_Supertip {
            get {
                return ResourceManager.GetString("UI_NotebookButton_Supertip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Click here for additional Journal notebook configuration options..
        /// </summary>
        internal static string UI_OptionsButton_Supertip {
            get {
                return ResourceManager.GetString("UI_OptionsButton_Supertip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Click here to customize your Journal notebook page settings..
        /// </summary>
        internal static string UI_SplitButton_Supertip {
            get {
                return ResourceManager.GetString("UI_SplitButton_Supertip", resourceCulture);
            }
        }
    }
}
