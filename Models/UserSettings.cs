using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class UserSettings
    {
        public static readonly string Path = @"ThesisSettings.json";

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
                if (!File.Exists(Path)) return new UserSettings();

                var settings = JsonConvert.DeserializeObject<UserSettings>(File.ReadAllText(Path));
                if (!string.IsNullOrEmpty(settings.FilePath))
                {
                    if (!File.Exists(settings.FilePath))
                        settings.FilePath = null;
                }
                return settings;
            }
            catch (Exception ex)
            {
                Logger.Log(LogItemType.Error, $"Error reading {Path}, creating a new settings file. ({ex.Message})");
                return new UserSettings();
            }
        }

        public void Save()
        {
            File.WriteAllText(Path, JsonConvert.SerializeObject(this));
        }
    }
}