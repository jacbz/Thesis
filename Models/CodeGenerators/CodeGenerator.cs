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
    }

    public enum Language
    {
        CSharp
    }
}
