using System.Collections.Generic;
using System.Linq;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class Label
    {
        public CellVertex Vertex { get; }
        public LabelType Type { get; set; }
        public string Text { get; set; }
        public List<Label> Attributes { get; set; }
        public List<Label> Headers { get; set; }

        public Label(CellVertex vertex)
        {
            Attributes = new List<Label>();
            Headers = new List<Label>();
            Vertex = vertex;
        }

        public void GenerateVariableName()
        {
            var variableName = "";
            if (Attributes.Count != 0 || Headers.Count != 0)
            {
                if (Headers.Count > 0)
                {
                    var headerStrings = Headers.Select(h => h.Text.ToPascalCase()).Reverse()
                        .Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                    variableName += string.Join("_", headerStrings) + "_";
                }
                var attributeStrings = Attributes.Select(h => h.Text.ToPascalCase()).Reverse()
                    .Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                variableName += string.Join("_", attributeStrings);
            }

            Vertex.VariableName = variableName.LowerFirstCharacter();
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
