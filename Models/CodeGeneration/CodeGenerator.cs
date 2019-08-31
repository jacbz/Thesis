using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Thesis.Models.VertexTypes;

namespace Thesis.Models.CodeGeneration
{
    public abstract class CodeGenerator
    {
        protected ClassCollection ClassCollection;
        protected Dictionary<(string worksheet, string address), CellVertex> AddressToVertexDictionary;
        protected Dictionary<string, RangeVertex> RangeDictionary;
        protected Dictionary<string, Vertex> NameDictionary;

        public abstract Task<Code> GenerateCodeAsync(TestResults testResults = null);
        protected abstract string[] LanguageKeywords { get; }

        protected CodeGenerator(ClassCollection classCollection,
            Dictionary<(string worksheet, string address), CellVertex> addressToVertexDictionary, 
            Dictionary<string, RangeVertex> rangeDictionary, Dictionary<string, Vertex> nameDictionary)
        {
            ClassCollection = classCollection;
            AddressToVertexDictionary = addressToVertexDictionary;
            RangeDictionary = rangeDictionary;
            NameDictionary = nameDictionary;
        }

        protected void GenerateVariableNamesForAll()
        {
            // generate unique class names
            var blockedClassNames = LanguageKeywords.ToHashSet();
            foreach (var generatedClass in ClassCollection.Classes)
            {
                generatedClass.Name = GenerateUniqueName(blockedClassNames, generatedClass.Name.MakeNameVariableConform());
            }

            // output vertex variable must be unique as they are in the Main class
            var blockedOutputVertexNames = new HashSet<string>();
            // generate unique variable names
            foreach (var generatedClass in ClassCollection.Classes)
            {
                var blockedVariableNames = blockedClassNames.ToHashSet();
                blockedVariableNames.UnionWith(blockedOutputVertexNames);

                foreach (var vertex in generatedClass.Vertices)
                {
                    vertex.Name = GenerateUniqueName(blockedVariableNames, vertex.Name.MakeNameVariableConform());
                    if (vertex == generatedClass.OutputVertex)
                        blockedOutputVertexNames.Add(vertex.Name);
                }
            }
        }

        protected string GenerateUniqueName(HashSet<string> usedVariableNames, string variableName)
        {
            variableName = GenerateNonDuplicateName(usedVariableNames, variableName);
            usedVariableNames.Add(variableName);
            return variableName;
        }

        /// <summary>
        /// var -> var2, var2 -> var3 etc.
        /// </summary>
        protected string GenerateNonDuplicateName(HashSet<string> usedVariableNames, string variableName)
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
}
