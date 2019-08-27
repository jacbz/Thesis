using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;

namespace Thesis.Models.VertexTypes
{
    public class RangeVertex : Vertex
    {
        public IRange[] CellsInRange { get; }

        public RangeVertex(IRange[] cellsInRange, string addressOrName)
        {
            CellsInRange = cellsInRange;
            VariableName = addressOrName.MakeNameVariableConform();
        }
    }
}
