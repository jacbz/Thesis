using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace Thesis
{
    public class UserSettings
    {
        public static readonly string PATH = @"thesissettings.json";
        public string FilePath { get; set; }

        // File specific settings
        public List<string> SelectedOutputFields { get; set; }

        public UserSettings()
        {
        }

        public void ResetFileSpecificSettings()
        {
            SelectedOutputFields = null;
        }

        public static UserSettings Read()
        {
            try
            {
                if (File.Exists(PATH))
                {
                    UserSettings settings = JsonConvert.DeserializeObject<UserSettings>(File.ReadAllText(PATH));
                    return settings;
                }
                else
                {
                    return new UserSettings();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, App.AppName);
                return new UserSettings();
            }
        }

        public void Save()
        {
            File.WriteAllText(PATH, JsonConvert.SerializeObject(this));
        }
    }
}
