using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using Syncfusion.UI.Xaml.Diagram;
using Thesis.Models;
using MColor = System.Windows.Media.Color;
using DColor = System.Drawing.Color;

namespace Thesis.ViewModels
{
    // Extension methods to layout and format objects
    public static class Formatter
    {
        private static readonly int DIAGRAM_PADDING = 40;
        private static readonly int VERTEX_BOX = 60; // width and height of a vertex including spacing
        private static readonly int CLASS_PADDING = 20; // padding inside classes
        private static readonly int CLASS_SPACING = 40; // spacing between classes

        public static DColor GetNodeTypeColor(this Vertex vertex)
        {
            MColor color;
            switch (vertex.NodeType)
            {
                case NodeType.OutputField:
                    color = (MColor) Application.Current.Resources["OutputColor"];
                    break;
                case NodeType.Formula:
                    color = (MColor)Application.Current.Resources["FormulaColor"];
                    break;
                case NodeType.Constant:
                    color = (MColor)Application.Current.Resources["ConstantColor"];
                    break;
                case NodeType.External:
                    color = (MColor)Application.Current.Resources["ExternalColor"];
                    break;
                default:
                    color = Colors.Transparent;
                    break;
            }
            return color.ToDColor();
        }

        public static NodeViewModel FormatVertex(this Vertex vertex, Graph graph)
        {
            return FormatVertex(vertex,
                graph.PopulatedColumns.IndexOf(vertex.Address.col) * VERTEX_BOX + DIAGRAM_PADDING,
                graph.PopulatedRows.IndexOf(vertex.Address.row) * VERTEX_BOX + DIAGRAM_PADDING);
        }

        public static NodeViewModel FormatVertex(this Vertex vertex, double posX, double posY)
        {
            var size = vertex.NodeType == NodeType.OutputField ? 40 : Math.Min(55, vertex.Parents.Count * 4 + 25);
            var node = new NodeViewModel
            {
                ID = vertex.StringAddress,
                Content = vertex,
                ContentTemplate = new DataTemplate(),
                UnitWidth = size,
                UnitHeight = size,
                OffsetX = posX,
                OffsetY = posY,
                ShapeStyle = GetNodeShapeStyle(Application.Current.Resources[
                    vertex.NodeType == NodeType.Formula
                        ? "FormulaColorBrush"
                        : vertex.NodeType == NodeType.OutputField
                            ? "OutputColorBrush"
                            : "ConstantColorBrush"] as SolidColorBrush),
                Shape = Application.Current.Resources[
                    vertex.NodeType == NodeType.Formula
                        ? "Heptagon"
                        : vertex.NodeType == NodeType.OutputField
                            ? "Trapezoid"
                            : "Ellipse"
                ],
                Annotations = new AnnotationCollection
                {
                    new AnnotationEditorViewModel
                    {
                        Offset = new Point(0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        Content = string.IsNullOrEmpty(vertex.VariableName) ? vertex.StringAddress : vertex.VariableName,
                        ViewTemplate =
                            Application.Current.Resources[!string.IsNullOrEmpty(vertex.VariableName) ? "normalLabel" : "redLabel"]
                                as DataTemplate,
                        UnitWidth = 200
                    }
                }
            };
            SetNodeConstraints(node);
            vertex.Node = node;
            return node;
        }

        private static Style GetNodeShapeStyle(SolidColorBrush solidColorBrush)
        {
            var shapeStyle = new Style
            {
                BasedOn = Application.Current.Resources["ShapeStyle"] as Style,
                TargetType = typeof(Path)
            };
            shapeStyle.Setters.Add(new Setter(Shape.FillProperty, solidColorBrush));
            return shapeStyle;
        }

        public static (GroupViewModel group, double nextPosX) FormatClass(this GeneratedClass generatedClass,
            double posX)
        {
            var graphLayout = LayoutGraph(generatedClass).ToList();
            var numOfFormulaColumns = graphLayout.Max(l => l.Count(v => v.NodeType != NodeType.Constant));

            var nodes = new ObservableCollection<NodeViewModel>();
            var group = new GroupViewModel
            {
                Nodes = nodes
            };

            double width = VERTEX_BOX * (numOfFormulaColumns + 1) + CLASS_PADDING * 2;
            posX = posX == 0 ? DIAGRAM_PADDING : posX;
            var nextPosX = posX + width + CLASS_SPACING;
            double posY = DIAGRAM_PADDING;
            double vertexBoxCenter = VERTEX_BOX / 2.0;

            var lastColumnX = posX + CLASS_PADDING + numOfFormulaColumns * VERTEX_BOX + vertexBoxCenter;
            var currentRowY = posY + CLASS_PADDING + vertexBoxCenter;

            double smallVertexHeight = 40;

            foreach (var row in graphLayout)
            {
                var formulas = row.Where(v => v.NodeType != NodeType.Constant).ToList();
                var constants = row.Where(v => v.NodeType == NodeType.Constant).ToList();

                if (constants.Count > 0)
                {
                    var startRowY = currentRowY;
                    // layout to the right, top-to-bottom, center the rest
                    foreach (var vertex in constants)
                    {
                        var node = vertex.FormatVertex(lastColumnX, currentRowY);
                        currentRowY += smallVertexHeight;
                        nodes.Add(node);
                    }

                    currentRowY += VERTEX_BOX - smallVertexHeight;

                    var middle = startRowY + (constants.Count - 1) / 2.0 * smallVertexHeight;
                    for (var i = 0; i < formulas.Count; i++)
                    {
                        var node = formulas[i].FormatVertex(posX + CLASS_PADDING + vertexBoxCenter + i * VERTEX_BOX,
                            middle);
                        nodes.Add(node);
                    }
                }
                else
                {
                    for (var i = 0; i < formulas.Count; i++)
                    {
                        var node = formulas[i].FormatVertex(posX + CLASS_PADDING + vertexBoxCenter + i * VERTEX_BOX,
                            currentRowY);
                        nodes.Add(node);
                    }

                    currentRowY += VERTEX_BOX;
                }
            }

            var classNode = new NodeViewModel
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
                Annotations = new AnnotationCollection
                {
                    new AnnotationEditorViewModel
                    {
                        Offset = new Point(0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        Content = generatedClass.Name,
                        ViewTemplate = Application.Current.Resources["classLabel"] as DataTemplate
                    }
                },
                ZIndex = int.MinValue,
                ShapeStyle = GetNodeShapeStyle(new SolidColorBrush(generatedClass.Color.ToMColor()))
            };
            SetNodeConstraints(classNode);

            nodes.Add(classNode);
            return (group, nextPosX);
        }

        // returns as list of vertex lists, each list is a row
        private static IEnumerable<List<Vertex>> LayoutGraph(GeneratedClass generatedClass)
        {
            if (generatedClass.OutputVertex == null)
            {
                foreach (var vertex in generatedClass.Vertices)
                    yield return new List<Vertex> {vertex};
            }
            else
            {
                var vertexQueue = new Queue<Vertex>(generatedClass.Vertices);
                yield return new List<Vertex> {vertexQueue.Dequeue()};
                while (vertexQueue.Count > 0)
                {
                    var vertex = vertexQueue.Dequeue();
                    var entry = new List<Vertex> {vertex};
                    while (vertexQueue.Count > 0 && vertex.Children.Contains(vertexQueue.Peek()))
                    {
                        var child = vertexQueue.Dequeue();
                        entry.Add(child);
                        // TODO cleanup
                        while (vertexQueue.Count > 0 && child.Children.Contains(vertexQueue.Peek()) &&
                               vertexQueue.Peek().NodeType == NodeType.Constant)
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
            node.Constraints = node.Constraints.Remove(NodeConstraints.Delete, NodeConstraints.InheritRotatable,
                NodeConstraints.Rotatable, NodeConstraints.Connectable);
        }

        public static ConnectorViewModel FormatEdge(this Vertex from, Vertex to, bool redirectSharedNodesToTopLeft = false)
        {
            var connector = new ConnectorViewModel
            {
                SourceNode = from.Node,
                ConnectorGeometryStyle = Application.Current.Resources["ConnectorGeometryStyle"] as Style,
                TargetDecoratorStyle = Application.Current.Resources["TargetDecoratorStyle"] as Style,
                Segments = new ObservableCollection<IConnectorSegment>
                {
                    new StraightSegment()
                },
                Constraints = ConnectorConstraints.Default & ~ConnectorConstraints.Selectable,
                ZIndex = -1
            };
            if (to.NodeType == NodeType.External && redirectSharedNodesToTopLeft)
                connector.TargetPoint = new Point(0, 0);
            else
                connector.TargetNode = to.Node;
            return connector;
        }

        public static DColor GetTextColor(this DColor c)
        {
            var l = 0.2126 * c.R + 0.7152 * c.G + 0.0722 * c.B;
            return l < 128 ? DColor.White : DColor.Black;
        }

        public static MColor ToMColor(this DColor color)
        {
            return MColor.FromArgb(color.A, color.R, color.G, color.B);
        }

        public static DColor ToDColor(this MColor color)
        {
            return DColor.FromArgb(color.A, color.R, color.G, color.B);
        }
    }
}