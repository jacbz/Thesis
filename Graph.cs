using Irony.Parsing;
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
                //if (cell.Value == "") continue;
                var vertex = new Vertex(cell);

                VerticesDict.Add(vertex.Address, vertex);
                Vertices.Add(vertex);
            }
            MainWindow.Log(LogItemType.Info, $"Adding {Vertices.Count} vertices...");

            var allAddresses = Vertices.Select(x => x.Address);

            foreach (var vertix in Vertices)
            {
                if (vertix.Type != CellType.Formula) continue;
                var formula = vertix.Formula.Replace(",", ".").Replace(";", ",").Replace("$", "");

                // TODO: create "external" vertices (SheetNameQuotedToken etc)
                //

                try
                {
                    var parseTree = XLParser.ExcelFormulaParser.Parse(formula);
                    foreach(var cell in GetListOfReferencedCells(parseTree))
                    {
                        Edges.Add(new Edge(vertix, VerticesDict[cell], ""));
                    }
                }
                catch (Exception ex)
                {
                    MainWindow.Log(LogItemType.Error, $"Error processing formula in {vertix.Address} ({formula}): {ex.Message}");
                    continue;
                }

            }
            MainWindow.Log(LogItemType.Info, $"Adding {Edges.Count} edges...");

            Vertices.RemoveAll(x => x.Parents.Count == 0 && x.Children.Count == 0);


            foreach(Vertex v in Vertices)
            {
                if (v.Parents.Count == 0) v.Parents.Add("root");
            }
            Vertex root = new Vertex("root");
            Vertices.Add(root);
        }

        // recursively gets list of referenced cells from parse tree
        private IEnumerable<string> GetListOfReferencedCells(ParseTreeNode parseTree)
        {
            if (parseTree.ChildNodes.Count == 0 && parseTree.Term.Name == "CellToken")
            {
                yield return parseTree.Token.Text;
            }
            foreach(var child in parseTree.ChildNodes)
            {
                foreach(var cell in GetListOfReferencedCells(child))
                {
                    yield return cell;
                }
            }
        }
    }
}
