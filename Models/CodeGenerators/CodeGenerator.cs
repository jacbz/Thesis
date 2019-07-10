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
        private protected List<GeneratedClass> generatedClasses;
        private protected Dictionary<string, Vertex> addressToVertexDictionary;

        private protected abstract string GetMainClass();
        private protected abstract string ClassToCode(GeneratedClass generatedClass);

        protected CodeGenerator(List<GeneratedClass> generatedClasses, Dictionary<string, Vertex> addressToVertexDictionary)
        {
            this.generatedClasses = generatedClasses;
            this.addressToVertexDictionary = addressToVertexDictionary;
        }

        public string GetCode()
        {
            return GetMainClass() + "\n\n" + string.Join("\n\n", generatedClasses.Select(c => ClassToCode(c)));
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
