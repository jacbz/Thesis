using Newtonsoft.Json;
using System;
using System.IO;
using System.Windows;

namespace Thesis
{
    public class UserSettings
    {
        public static readonly string PATH = @"settings.json";
        public string FilePath { get; set; }

        public UserSettings()
        {
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
