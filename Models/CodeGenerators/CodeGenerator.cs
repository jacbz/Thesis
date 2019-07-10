using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Thesis.Models.CodeGenerators
{
    public abstract class CodeGenerator
    {
        private protected GeneratedClass generatedClass;
        private protected Dictionary<string, Vertex> addressToVertexDictionary;

        public abstract string ClassToCode();
        public abstract string VertexToCode(Vertex vertex);

        protected CodeGenerator(GeneratedClass generatedClass, Dictionary<string, Vertex> addressToVertexDictionary)
        {
            this.generatedClass = generatedClass;
            this.addressToVertexDictionary = addressToVertexDictionary;
        }

        public static string ToCamelCase(string inputString)
        {
            var output = ToPascalCase(inputString);
            return output.First().ToString().ToLower() + output.Substring(1);
        }

        public static string ToPascalCase(string inputString)
        {
            var textInfo = new CultureInfo("en-US", false).TextInfo;
            inputString = Regex.Replace(RemoveDiacritics(inputString), "[^0-9a-zA-Z ]+", "");
            return textInfo.ToTitleCase(inputString).Replace(" ", "");
        }

        // ö => oe etc.
        public static string RemoveDiacritics(string s)
        {
            var normalizedString = s.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();
            foreach (var c in normalizedString)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    stringBuilder.Append(c);
            }

            return stringBuilder.ToString();
        }
    }

    public enum Language
    {
        CSharp
    }
}
