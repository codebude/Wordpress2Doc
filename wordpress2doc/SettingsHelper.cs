using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace Wordpress2Doc
{
    class SettingsHelper
    {
        public static string GetAppSetting(string key)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);           
            return (config.AppSettings.Settings[key] == null) ? null : config.AppSettings.Settings[key].Value;
        }

        public static void SetAppSetting(string key, string value)
        {        
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);
          
            if (config.AppSettings.Settings[key] != null)
                config.AppSettings.Settings.Remove(key);
            config.AppSettings.Settings.Add(key, value);
            config.Save(ConfigurationSaveMode.Modified);
        }
    }
}
