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

        public abstract string GenerateCode();

        protected CodeGenerator(List<GeneratedClass> generatedClasses, Dictionary<string, Vertex> addressToVertexDictionary)
        {
            this.generatedClasses = generatedClasses;
            this.addressToVertexDictionary = addressToVertexDictionary;
        }

        /// <summary>
        /// var -> var2, var2 -> var3 etc.
        /// </summary>
        protected static string GenerateNonDuplicateName(HashSet<string> usedVariableNames, string variableName)
        {
            while (usedVariableNames.Contains(variableName))
            {
                var pattern = @"\d+$";
                var matchNumberAtEnd = Regex.Match(variableName, pattern);
                if (matchNumberAtEnd.Success)
                    variableName = Regex.Replace(variableName, pattern,
                        (Convert.ToInt32(matchNumberAtEnd.Value) + 1).ToString());
                else
                    variableName += "2";
            }
            return variableName;
        }

    }

    public enum Language
    {
        CSharp
    }
}
