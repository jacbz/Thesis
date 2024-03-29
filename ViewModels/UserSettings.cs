﻿// Thesis - An Excel to code converter
// Copyright (C) 2019 Jacob Zhang
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace Thesis.ViewModels
{
    public class UserSettings
    {
        public static readonly string Path = @"ThesisSettings.json";

        private string _selectedFile;

        public string SelectedFile
        {
            get => _selectedFile;
            set
            {
                _selectedFile = value;
                if (_selectedFile != null && CurrentFileSettings == null)
                    _fileDataDictionary.Add(_selectedFile, new FileSettings());
            }
        }

        private string _selectedWorksheet;
        public string SelectedWorksheet
        {
            get => _selectedWorksheet;
            set
            {
                _selectedWorksheet = value;
                if (SelectedWorksheet != null && CurrentWorksheetSettings == null)
                    CurrentFileSettings.WorksheetDataDictionary
                        .Add(_selectedWorksheet, new WorksheetSettings());
            }
        }


        [JsonIgnore]
        public FileSettings CurrentFileSettings =>
            SelectedFile != null && _fileDataDictionary.TryGetValue(SelectedFile, out var fileData)
                ? fileData
                : null;

        [JsonIgnore]
        public WorksheetSettings CurrentWorksheetSettings =>
            CurrentFileSettings != null && CurrentFileSettings.WorksheetDataDictionary
                .TryGetValue(SelectedWorksheet, out var worksheetData)
                ? worksheetData
                : null;

        [JsonProperty(PropertyName = "FileSettings")]
        private readonly Dictionary<string, FileSettings> _fileDataDictionary = new Dictionary<string, FileSettings>();

        public class FileSettings
        {
            [JsonProperty(PropertyName = "WorksheetSettings")]
            public Dictionary<string, WorksheetSettings> WorksheetDataDictionary { get; }

            public FileSettings()
            {
                WorksheetDataDictionary = new Dictionary<string, WorksheetSettings>();
            }
        }

        public class WorksheetSettings
        {
            public List<string> SelectedOutputFields { get; set; }

            public WorksheetSettings()
            {
            }
        }

        public static UserSettings Read()
        {
            try
            {
                if (!File.Exists(Path)) return new UserSettings();

                var settings = JsonConvert.DeserializeObject<UserSettings>(File.ReadAllText(Path));
                if (!string.IsNullOrEmpty(settings.SelectedFile))
                {
                    if (!File.Exists(settings.SelectedFile))
                        settings.SelectedFile = null;
                }
                return settings;
            }
            catch (Exception ex)
            {
                Logger.Log(LogItemType.Error, $"Error reading {Path}, creating a new settings file. ({ex.Message})");
                return new UserSettings();
            }
        }

        public void Persist()
        {
            File.WriteAllText(Path, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }
}