using Syncfusion.UI.Xaml.Diagram;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using Thesis.Models;
using Color = System.Drawing.Color;
using Point = System.Windows.Point;

namespace Thesis.ViewModels
{
    // Extension methods to layout and format objects
    public static class Formatter
    {
        public static string GetColor(this Vertex vertex)
        {
            switch (vertex.NodeType)
            {
                case NodeType.OutputField:
                    return "#ff0000";
                case NodeType.Formula:
                    return "#8e44ad";
                default:
                    return "#2980b9";
            }
        }

        private static readonly int DIAGRAM_PADDING = 40;
        private static readonly int VERTEX_BOX = 60; // width and height of a vertex including spacing
        private static readonly int CLASS_PADDING = 20; // padding inside classes
        private static readonly int CLASS_SPACING = 40; // spacing between classes

        public static NodeViewModel FormatVertex(this Vertex vertex, Graph graph)
        {
            return FormatVertex(vertex,
                graph.PopulatedColumns.IndexOf(vertex.CellIndex[1]) * VERTEX_BOX + DIAGRAM_PADDING,
                graph.PopulatedRows.IndexOf(vertex.CellIndex[0]) * VERTEX_BOX + DIAGRAM_PADDING);
        }

        public static NodeViewModel FormatVertex(this Vertex vertex, double posX, double posY)
        {
            var size = vertex.NodeType == NodeType.OutputField ? 40 : Math.Min(55, vertex.Parents.Count * 4 + 25);
            var node = new NodeViewModel()
            {
                ID = vertex.Address,
                Content = vertex,
                ContentTemplate = new DataTemplate(),
                UnitWidth = size,
                UnitHeight = size,
                OffsetX = posX,
                OffsetY = posY,
                Shape = Application.Current.Resources[
                    vertex.NodeType == NodeType.Formula 
                        ? "Heptagon" 
                        : (vertex.NodeType == NodeType.OutputField ? "Trapezoid" : "Ellipse")
                        ],
                Annotations = new AnnotationCollection()
                {
                    new AnnotationEditorViewModel()
                    {
                        Offset = new Point(0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        Content = vertex.HasLabel ? vertex.Label : vertex.Address,
                        ViewTemplate = Application.Current.Resources[vertex.HasLabel ? "normalLabel" : "redLabel"] as DataTemplate,
                        UnitWidth = 200
                    }
                }
            };
            node.ShapeStyle = GetNodeShapeStyle(Application.Current.Resources[
                vertex.NodeType == NodeType.Formula
                        ? "FormulaColorBrush"
                        : (vertex.NodeType == NodeType.OutputField ? "OutputColorBrush" : "FormulaColorBrush")] as SolidColorBrush);
            SetNodeConstraints(node);
            vertex.Node = node;
            return node;
        }

        private static Style GetNodeShapeStyle(SolidColorBrush solidColorBrush)
        {
            Style shapeStyle = new Style
            {
                BasedOn = Application.Current.Resources["ShapeStyle"] as Style,
                TargetType = typeof(Path)
            };
            shapeStyle.Setters.Add(new Setter(Shape.FillProperty, solidColorBrush));
            return shapeStyle;
        }

        public static (GroupViewModel group, double nextPosX) FormatClass(this GeneratedClass generatedClass, double posX)
        {
            var graphLayout = LayoutGraph(generatedClass).Reverse().ToList();
            var numOfFormulaColumns = graphLayout.Max(l => l.Count(v => v.Type == CellType.Formula));

            GroupViewModel group = new GroupViewModel()
            {
                Nodes = new ObservableCollection<NodeViewModel>(),
            };
            var nodes = group.Nodes as ObservableCollection<NodeViewModel>;

            double width = (VERTEX_BOX * (numOfFormulaColumns + 1)) + CLASS_PADDING * 2;
            posX = posX == 0 ? DIAGRAM_PADDING : posX;
            double nextPosX = posX + width + CLASS_SPACING;
            double posY = DIAGRAM_PADDING;
            double vertexBoxCenter = VERTEX_BOX / 2;

            double lastColumnX = posX + CLASS_PADDING + (numOfFormulaColumns * VERTEX_BOX) + vertexBoxCenter;
            double currentRowY = posY + CLASS_PADDING + vertexBoxCenter;

            double smallVertexHeight = 40;

            foreach (List<Vertex> row in graphLayout)
            {
                var formulas = row.Where(v => v.Type == CellType.Formula).ToList();
                var constants = row.Where(v => v.Type != CellType.Formula).ToList();

                if (constants.Count > 0)
                {
                    double startRowY = currentRowY;
                    // layout to the right, top-to-bottom, center the rest
                    foreach (Vertex vertex in constants)
                    {
                        var node = vertex.FormatVertex(lastColumnX, currentRowY);
                        currentRowY += smallVertexHeight;
                        nodes.Add(node);
                    }
                    currentRowY += VERTEX_BOX - smallVertexHeight;

                    double middle = startRowY + (constants.Count - 1) / 2 * smallVertexHeight;
                    for (int i = 0; i < formulas.Count; i++)
                    {
                        var node = formulas[i].FormatVertex(posX + CLASS_PADDING + vertexBoxCenter + (i*VERTEX_BOX), middle);
                        nodes.Add(node);
                    }
                }
                else
                {
                    for (int i = 0; i < formulas.Count; i++)
                    {
                        var node = formulas[i].FormatVertex(posX + CLASS_PADDING + vertexBoxCenter + (i * VERTEX_BOX), currentRowY);
                        nodes.Add(node);
                    }
                    currentRowY += VERTEX_BOX;
                }
            }

            var classNode = new NodeViewModel()
            {
                ID = generatedClass.Name,
                Content = generatedClass,
                ContentTemplate = new DataTemplate(),
                UnitWidth = width,
                UnitHeight = currentRowY - VERTEX_BOX + CLASS_PADDING,
                Pivot = new Point(0, 0),
                OffsetX = posX,
                OffsetY = posY,
                Shape = Application.Current.Resources["Rectangle"],
                Annotations = new AnnotationCollection()
                {
                    new AnnotationEditorViewModel()
                    {
                        Offset = new Point(0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        Content = generatedClass.Name,
                        ViewTemplate = Application.Current.Resources["classLabel"] as DataTemplate
                    }
                },
                ZIndex = int.MinValue
            };
            classNode.ShapeStyle = GetNodeShapeStyle(new SolidColorBrush(generatedClass.Color.ToMediaColor()));
            SetNodeConstraints(classNode);

            nodes.Add(classNode);
            return (group, nextPosX);
        }

        private static IEnumerable<List<Vertex>> LayoutGraph(GeneratedClass generatedClass)
        {
            if (generatedClass.OutputVertex == null)
            {
                foreach (var vertex in generatedClass.Vertices)
                    yield return new List<Vertex> { vertex };
            }
            else
            {
                var vertexQueue = new Queue<Vertex>(generatedClass.Vertices);
                yield return new List<Vertex> { vertexQueue.Dequeue() };
                while (vertexQueue.Count > 0)
                {
                    var vertex = vertexQueue.Dequeue();
                    var entry = new List<Vertex> { vertex };
                    while (vertexQueue.Count > 0 && vertex.Children.Contains(vertexQueue.Peek()))
                    {
                        var child = vertexQueue.Dequeue();
                        entry.Add(child);
                        // TODO cleanup
                        while (vertexQueue.Count > 0 && child.Children.Contains(vertexQueue.Peek()) && vertexQueue.Peek().Type != CellType.Formula)
                        {
                            var child1 = vertexQueue.Dequeue();
                            entry.Add(child1);
                        }
                    }
                    
                    yield return entry;
                }
            }
        }

        private static void SetNodeConstraints(NodeViewModel node)
        {
            node.Constraints = node.Constraints.Remove(NodeConstraints.Delete, NodeConstraints.InheritRotatable, NodeConstraints.Rotatable, NodeConstraints.Connectable);
        }

        public static ConnectorViewModel FormatEdge(this Vertex from, Vertex to)
        {
            return new ConnectorViewModel()
            {
                SourceNode = from.Node,
                TargetNode = to.Node,
                ConnectorGeometryStyle = Application.Current.Resources["ConnectorGeometryStyle"] as Style,
                TargetDecoratorStyle = Application.Current.Resources["TargetDecoratorStyle"] as Style,
                Segments = new ObservableCollection<IConnectorSegment>()
                {
                    new StraightSegment()
                },
                Constraints = ConnectorConstraints.Default & ~ConnectorConstraints.Selectable,
                ZIndex = -1
            };
        }

        public static Color GetTextColor(this Color c)
        {
            var l = 0.2126 * c.R + 0.7152 * c.G + 0.0722 * c.B;
            return l < 128 ? Color.White : Color.Black;
        }

        public static System.Windows.Media.Color ToMediaColor(this Color color)
        {
            return System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B);
        }
    }
}
