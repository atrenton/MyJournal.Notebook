using System;
using System.Configuration;

namespace MyJournal.Notebook.Config
{
    /// <summary>
    /// Custom indexed configuration property for the OneNote AddIn DLL config
    /// file.
    /// </summary>
    class AppSettings
    {
        readonly KeyValueConfigurationCollection _settings;

        internal AppSettings()
        {
            try
            {
                var dllConfigPath = GetType().Assembly.Location;
                var config = ConfigurationManager.OpenExeConfiguration(dllConfigPath);
                _settings = config.AppSettings.Settings;
            }
            catch (Exception e)
            {
                Utils.ExceptionHandler.HandleException(e);
            }
        }

        public string this[string key]
        {
            get
            {
                string result = null;
                if (_settings != null)
                {
                    var element = _settings[key];
                    if (element != null) result = element.Value.Trim();
                }
                return result;
            }
        }
    }
}
