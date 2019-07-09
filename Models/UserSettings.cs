using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Newtonsoft.Json;
using Thesis.ViewModels;

namespace Thesis
{
    public class UserSettings
    {
        public static readonly string PATH = @"thesissettings.json";

        public string FilePath { get; set; }

        // File specific settings
        public string SelectedWorksheet { get; set; }
        public List<string> SelectedOutputFields { get; set; }

        public void ResetWorkbookSpecificSettings()
        {
            SelectedWorksheet = null;
            SelectedOutputFields = null;
            Save();
        }

        public static UserSettings Read()
        {
            try
            {
                if (!File.Exists(PATH)) return new UserSettings();

                var settings = JsonConvert.DeserializeObject<UserSettings>(File.ReadAllText(PATH));
                if (!string.IsNullOrEmpty(settings.FilePath))
                {
                    if (!File.Exists(settings.FilePath))
                        settings.FilePath = null;
                }
                return settings;
            }
            catch (Exception ex)
            {
                Logger.Log(LogItemType.Error, $"Error reading {PATH}, creating a new settings file. ({ex.Message})");
                return new UserSettings();
            }
        }

        public void Save()
        {
            File.WriteAllText(PATH, JsonConvert.SerializeObject(this));
        }
    }
}