using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Thesis.Models.CodeGenerators
{
    public abstract class Tester
    {
        public abstract List<TestResult> TestClasses(List<ClassCode> classesCode);
    }

    public class ClassCode
    {
        public bool IsSharedClass { get; }
        public string ClassName { get; }
        public string Code { get; }
        public string FieldsCode { get; }
        public string MethodCode { get; }

        public ClassCode(bool isSharedClass, string className, string code, string fieldsCode, string methodCode)
        {
            IsSharedClass = isSharedClass;
            ClassName = className;
            Code = code;
            FieldsCode = fieldsCode;
            MethodCode = methodCode;
        }
    }

    public class TestResult
    {
        public string ClassName { get; }
        public string VariableName { get; }
        public dynamic Value { get; }
        public Type Type { get; }

        public TestResult(string className, string variableName, dynamic value, Type type)
        {
            ClassName = className;
            VariableName = variableName;
            Value = value;
            Type = type;
        }
    }
}
