using ExcelFormulaParser;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Thesis
{
    class Graph
    {
        public List<Vertex> Vertices { get; set; }
        public List<Edge> Edges { get; set; }
        public Graph(IRange cells)
        {
            var VerticesDict = new Dictionary<string, Vertex>();
            Vertices = new List<Vertex>();
            Edges = new List<Edge>();

            foreach(var cell in cells.Cells)
            {
                if (cell.Value == "") continue;
                var vertex = new Vertex(cell);

                VerticesDict.Add(vertex.Address, vertex);
                Vertices.Add(vertex);
            }

            var allAddresses = Vertices.Select(x => x.Address);

            foreach (var vertix in Vertices)
            {
                if (vertix.Type != CellType.Formula) continue;
                var formula = vertix.Formula.Replace(",", ".");

                try
                {
                    ExcelFormula excelFormula = new ExcelFormula(formula);

                    foreach (var formulaToken in excelFormula)
                    {
                        if (formulaToken.Type != ExcelFormulaTokenType.Operand) continue;
                        if (allAddresses.Contains(formulaToken.Value))
                        {
                            Edges.Add(new Edge(vertix, VerticesDict[formulaToken.Value], ""));
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, App.AppName);
                    continue;
                }

            }

            Vertices.RemoveAll(x => x.Parents.Count == 0 && x.Children.Count == 0);


            foreach(Vertex v in Vertices)
            {
                if (v.Parents.Count == 0) v.Parents.Add("root");
            }
            Vertex root = new Vertex("root");
            Vertices.Add(root);
        }
    }
}
