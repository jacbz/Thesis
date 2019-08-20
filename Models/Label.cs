using System.Collections.Generic;
using System.Linq;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class Label
    {
        public LabelType Type { get; set; }
        public string Text { get; set; }
        public List<Label> Attributes { get; set; }
        public List<Label> Headers { get; set; }
        public string VariableName { get; set; }

        public Label()
        {
            Attributes = new List<Label>();
            Headers = new List<Label>();
        }

        public void GenerateVariableName()
        {
            VariableName = "";
            if (Attributes.Count != 0 || Headers.Count != 0)
            {
                if (Headers.Count > 0)
                {
                    var headerStrings = Headers.Select(h => h.Text.ToPascalCase()).Reverse()
                        .Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                    VariableName += string.Join("_", headerStrings) + "_";
                }
                var attributeStrings = Attributes.Select(h => h.Text.ToPascalCase()).Reverse()
                    .Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                VariableName += string.Join("_", attributeStrings);
                VariableName = VariableName.FirstToLower();
            }
        }
    }

    public enum LabelType
    {
        Header,
        Attribute,
        Data,
        None
    }
}
