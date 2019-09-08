using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Thesis.Models.VertexTypes;

namespace Thesis.Models
{
    public class Name : IComparable, IComparer<Name>
    {
        private string _customName;
        private string _attributes;
        private string _headers;
        private string _address;

        public bool IsFunction { get; set; }
        public bool IsOutputField { get; set; }

        private int _suffix = 1;
        private bool _useHeader;
        private bool _useAddress;
        private bool _useSuffix;

        public static void MakeNamesUnique(IEnumerable<Name> names, HashSet<string> blockedNames = null)
        {
            if (blockedNames == null)
                blockedNames = new HashSet<string>();

            var nameQueue = new Queue<Name>(names);
            while (nameQueue.Count > 0)
            {
                var name = nameQueue.Dequeue();
                if (blockedNames.Contains(name.ToString()))
                {
                    name.MakeMoreUnique();
                    nameQueue.Enqueue(name);
                }
                else
                {
                    blockedNames.Add(name.ToString());
                }
            }
        }

        // only copy name components, not useHeader etc
        public Name Copy()
        {
            return new Name(_customName, _address)
            {
                _attributes = _attributes,
                _headers = _headers,
                IsFunction = IsFunction,
                IsOutputField = IsOutputField
            };
        }

        public void MakeMoreUnique()
        {
            if (!_useHeader)
            {
                _useHeader = true;
                return;
            }

            if (!_useAddress)
            {
                _useAddress = true;
                return; 
            }

            _useSuffix = true;
            _suffix++;
        }

        public Name(string customName, string address)
        {
            _customName = customName != null ? MakeNameVariableConform(customName) : null;
            _address = address;
        }

        public Name(CellVertex[] attributes, CellVertex[] headers, string address)
        {
            _attributes = attributes == null
                ? null
                : string.Join("_", attributes
                    .Select(cell => ToPascalCase(cell.DisplayValue))
                    .Where(s => !string.IsNullOrWhiteSpace(s) && s != "_")
                    .ToList());
            _headers = headers == null
                ? null
                : string.Join("_", headers
                    .Select(cell => ToPascalCase(cell.DisplayValue))
                    .Where(s => !string.IsNullOrWhiteSpace(s) && s != "_")
                    .ToList());
            _address = MakeNameVariableConform(address);
        }

        public override string ToString()
        {
            if (IsFunction)
                return "Calc" + ToStringInner().RaiseFirstCharacter();
            if (IsOutputField)
                return "OUTPUT_" + ToStringInner().RaiseFirstCharacter();
            return ToStringInner();
        }
        private string ToStringInner()
        {
            if (!string.IsNullOrWhiteSpace(_customName))
                return _customName + (_useSuffix ? "_" + _suffix : "");

            var nameComponents = new List<string>();
            if (!string.IsNullOrWhiteSpace(_attributes))
                nameComponents.Add(_attributes);
            if (!string.IsNullOrWhiteSpace(_headers) && (_useHeader || string.IsNullOrWhiteSpace(_attributes)))
                nameComponents.Add(_headers);
            if (_useAddress || string.IsNullOrWhiteSpace(_attributes) && string.IsNullOrWhiteSpace(_headers))
                nameComponents.Add(_address);
            if (_useSuffix)
                nameComponents.Add(_suffix.ToString());

            return string.Join("_", nameComponents).LowerFirstCharacter();
        }

        public static implicit operator string(Name name)
        {
            return name.ToString();
        }

        public int CompareTo(object obj)
        {
            return string.Compare(ToString(), obj.ToString(), StringComparison.Ordinal);
        }

        public int Compare(Name name1, Name name2)
        {
            return name1.CompareTo(name2);
        }

        public static string ToPascalCase(string inputString)
        {
            if (inputString == "") return "";
            var textInfo = new CultureInfo("en-US", false).TextInfo;
            inputString = textInfo.ToTitleCase(inputString);
            inputString = MakeNameVariableConform(inputString);
            if (inputString == "") return "";
            return inputString;
        }

        public static string MakeNameVariableConform(string inputString)
        {
            inputString = ProcessDiacritics(inputString);
            inputString = inputString.Replace("%", "Percent");
            inputString = Regex.Replace(inputString, "[^0-9a-zA-Z_]+", "");
            if (inputString == "") return "_";
            if (inputString.Length > 0 && char.IsDigit(inputString.ToCharArray()[0]))
            {
                if (char.IsDigit(inputString.ToCharArray()[0]))
                    inputString = "_" + inputString;
            }
            return inputString;
        }

        public static string ProcessDiacritics(string inputString)
        {
            inputString = inputString
                .Replace("ö", "oe")
                .Replace("ä", "ae")
                .Replace("ü", "ue")
                .Replace("ß", "ss");
            var normalizedString = inputString.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();
            foreach (var c in normalizedString)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    stringBuilder.Append(c);
            }

            return stringBuilder.ToString();
        }
    }
}
