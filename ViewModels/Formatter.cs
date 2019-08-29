using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using Syncfusion.UI.Xaml.Diagram;
using Thesis.Models;
using Thesis.Models.VertexTypes;
using MColor = System.Windows.Media.Color;
using DColor = System.Drawing.Color;

namespace Thesis.ViewModels
{
    // Extension methods to layout and format objects
    public static class Formatter
    {
        public static CultureInfo CurrentCultureInfo;

        private static Style _formulaShapeStyle;
        private static Style _outputShapeStyle;
        private static Style _constantShapeStyle;
        private static Style _rangeShapeStyle;

        private static object _formulaShape;
        private static object _outputShape;
        private static object _constantShape;
        private static object _classShape;
        private static object _rangeShape;

        private static DataTemplate _normalLabelTemplate;
        private static DataTemplate _redLabelTemplate;
        private static DataTemplate _classLabelTemplate;
        private static DataTemplate _staticLabelTemplate;

        private static Style _connectorGeometryStyle;
        private static Style _externalConnectorGeometryStyle;
        private static Style _targetDecoratorStyle;
        private static Style _externalTargetDecoratorStyle;

        private static DColor _externalNodeColor;
        private static DColor _outputNodeColor;
        private static DColor _formulaNodeColor;
        private static DColor _constantNodeColor;

        public static void InitXamlStyles()
        {
            CurrentCultureInfo = CultureInfo.CurrentCulture;

            _formulaShapeStyle = GetNodeShapeStyle(Application.Current.Resources["FormulaColorBrush"] as SolidColorBrush);
            _outputShapeStyle = GetNodeShapeStyle(Application.Current.Resources["OutputColorBrush"] as SolidColorBrush);
            _constantShapeStyle = GetNodeShapeStyle(Application.Current.Resources["ConstantColorBrush"] as SolidColorBrush);
            _rangeShapeStyle = GetNodeShapeStyle(Application.Current.Resources["RangeColorBrush"] as SolidColorBrush);

            _formulaShape = Application.Current.Resources["Heptagon"];
            _outputShape = Application.Current.Resources["Trapezoid"];
            _constantShape = Application.Current.Resources["Ellipse"];
            _classShape = Application.Current.Resources["Rectangle"];
            _rangeShape = Application.Current.Resources["Rectangle"];

            _normalLabelTemplate = Application.Current.Resources["normalLabel"] as DataTemplate;
            _redLabelTemplate = Application.Current.Resources["redLabel"] as DataTemplate;
            _classLabelTemplate = Application.Current.Resources["classLabel"] as DataTemplate;
            _staticLabelTemplate = Application.Current.Resources["staticLabel"] as DataTemplate;

            _connectorGeometryStyle = Application.Current.Resources["ConnectorGeometryStyle"] as Style;
            _externalConnectorGeometryStyle = Application.Current.Resources["ExternalConnectorGeometryStyle"] as Style;
            _targetDecoratorStyle = Application.Current.Resources["TargetDecoratorStyle"] as Style;
            _externalTargetDecoratorStyle = Application.Current.Resources["ExternalTargetDecoratorStyle"] as Style;

            _externalNodeColor = ((MColor) Application.Current.Resources["ExternalColor"]).ToDColor();
            _outputNodeColor = ((MColor)Application.Current.Resources["OutputColor"]).ToDColor();
            _formulaNodeColor = ((MColor)Application.Current.Resources["FormulaColor"]).ToDColor();
            _constantNodeColor = ((MColor)Application.Current.Resources["ConstantColor"]).ToDColor();
        }

        private const int DIAGRAM_PADDING = 40;
        private const int VERTEX_BOX = 60; // width and height of a vertex including spacing
        private const int CLASS_PADDING = 20; // padding inside classes
        private const int CLASS_SPACING = 40; // spacing between classes
        private const double VERTEX_BOX_CENTER = VERTEX_BOX / 2.0;

        public static DColor GetNodeTypeColor(this CellVertex cellVertex)
        {
            if (cellVertex.IsExternal)
            {
                return _externalNodeColor;
            }

            switch (cellVertex.NodeType)
            {
                case NodeType.OutputField:
                    return _outputNodeColor;
                case NodeType.Formula:
                    return _formulaNodeColor;
                case NodeType.Constant:
                    return _constantNodeColor;
                default:
                    return DColor.Transparent;
            }
        }

        public static NodeViewModel FormatCellVertex(this CellVertex cellVertex, Graph graph)
        {
            return FormatCellVertex(cellVertex,
                Array.IndexOf(graph.PopulatedColumns, cellVertex.Address.col) * VERTEX_BOX + DIAGRAM_PADDING,
                Array.IndexOf(graph.PopulatedRows, cellVertex.Address.row) * VERTEX_BOX + DIAGRAM_PADDING);
        }

        public static NodeViewModel FormatCellVertex(this CellVertex cellVertex, double posX, double posY)
        {
            var size = cellVertex.NodeType == NodeType.OutputField ? 40 : Math.Min(55, cellVertex.Parents.Count * 4 + 25);
            var node = new NodeViewModel
            {
                ID = cellVertex.StringAddress,
                Content = cellVertex,
                ContentTemplate = new DataTemplate(),
                UnitWidth = size,
                UnitHeight = size,
                OffsetX = posX,
                OffsetY = posY,
                ShapeStyle =  cellVertex.NodeType == NodeType.Formula
                        ? _formulaShapeStyle
                        : cellVertex.NodeType == NodeType.OutputField
                            ? _outputShapeStyle
                            : _constantShapeStyle,
                Shape = 
                    cellVertex.NodeType == NodeType.Formula
                        ? _formulaShape
                        : cellVertex.NodeType == NodeType.OutputField
                            ? _outputShape
                            : _constantShape,
                ZIndex = 10000,
                Annotations = new AnnotationCollection
                {
                    new AnnotationEditorViewModel
                    {
                        Offset = new Point(0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        Content = string.IsNullOrEmpty(cellVertex.VariableName) ? cellVertex.StringAddress : cellVertex.VariableName,
                        ViewTemplate =!string.IsNullOrEmpty(cellVertex.VariableName) ? _normalLabelTemplate : _redLabelTemplate,
                        UnitWidth = 200
                    }
                }
            };
            SetNodeConstraints(node);
            cellVertex.Node = node;
            return node;
        }

        // formats a range vertex (large)
        public static NodeViewModel FormatRangeVertexLarge(this RangeVertex rangeVertex, Graph graph)
        {
            double width = rangeVertex.ColumnCount * VERTEX_BOX - VERTEX_BOX / 3.0;
            double height = rangeVertex.RowCount * VERTEX_BOX - VERTEX_BOX / 3.0;
            var posX = Array.IndexOf(graph.PopulatedColumns, rangeVertex.StartAddress.column) * VERTEX_BOX + DIAGRAM_PADDING 
                                                                                                           + width / 2.0 - VERTEX_BOX / 3.0;
            var posY = Array.IndexOf(graph.PopulatedRows, rangeVertex.StartAddress.row) * VERTEX_BOX + DIAGRAM_PADDING 
                                                                                                     + height / 2.0 - VERTEX_BOX / 3.0;

            return FormatRangeVertex(rangeVertex, posX, posY, width, height);
        }

        public static NodeViewModel FormatRangeVertexSmall(this RangeVertex rangeVertex, double posX, double posY)
        {
            return FormatRangeVertex(rangeVertex, posX, posY);
        }

        public static NodeViewModel FormatRangeVertex(this RangeVertex rangeVertex, 
            double posX, double posY, double width = 40, double height = 40)
        {
            var node = new NodeViewModel
            {
                ID = rangeVertex.VariableName,
                Content = rangeVertex,
                ContentTemplate = new DataTemplate(),
                UnitWidth = width,
                UnitHeight = height,
                OffsetX = posX,
                OffsetY = posY,
                ShapeStyle = _rangeShapeStyle,
                Shape = _rangeShape,
                Annotations = new AnnotationCollection
                {
                    new AnnotationEditorViewModel
                    {
                        Offset = new Point(0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        Content = rangeVertex.VariableName,
                        ViewTemplate = _normalLabelTemplate,
                        UnitWidth = 200
                    }
                }
            };
            SetNodeConstraints(node);
            rangeVertex.Node = node;
            return node;
        }

        private static Style GetNodeShapeStyle(SolidColorBrush solidColorBrush)
        {
            var shapeStyle = new Style
            {
                BasedOn = Application.Current.Resources["ShapeStyle"] as Style,
                TargetType = typeof(Path),
            };
            shapeStyle.Setters.Add(new Setter(Shape.FillProperty, solidColorBrush));
            return shapeStyle;
        }

        public static (GroupViewModel group, double nextPosX) FormatClass(this Class @class,
            double posX)
        {
            var graphLayout = LayoutGraph(@class).ToList();
            var numOfFormulaColumns = graphLayout.Count == 0 
                ? 0 
                : graphLayout.Max(l => l.Count(vertex => vertex is RangeVertex ||
                                                    vertex is CellVertex cellVertex && cellVertex.NodeType != NodeType.Constant));

            var nodes = new NodeCollection();
            var group = new GroupViewModel
            {
                Nodes = nodes
            };

            double width = VERTEX_BOX * (numOfFormulaColumns + 1) + CLASS_PADDING * 2;
            posX = posX == 0 ? DIAGRAM_PADDING : posX;
            var nextPosX = posX + width + CLASS_SPACING;
            double posY = DIAGRAM_PADDING;

            var lastColumnX = posX + CLASS_PADDING + numOfFormulaColumns * VERTEX_BOX + VERTEX_BOX_CENTER;
            var currentRowY = posY + CLASS_PADDING + VERTEX_BOX_CENTER;

            double smallVertexHeight = 40;

            foreach (var row in graphLayout)
            {
                var rightColumn = row
                    .Where(vertex =>  vertex is CellVertex cellVertex && cellVertex.NodeType == NodeType.Constant)
                    .Cast<CellVertex>()
                    .ToList();
                var leftColumns = row.Except(rightColumn).ToList();

                if (rightColumn.Count > 0)
                {
                    var startRowY = currentRowY;
                    // layout to the right, top-to-bottom, center the rest
                    foreach (var vertex in rightColumn)
                    {
                        var node = vertex.FormatCellVertex(lastColumnX, currentRowY);
                        currentRowY += smallVertexHeight;
                        nodes.Add(node);
                    }

                    currentRowY += VERTEX_BOX - smallVertexHeight;

                    var middle = startRowY + (rightColumn.Count - 1) / 2.0 * smallVertexHeight;
                    for (var i = 0; i < leftColumns.Count; i++)
                    {
                        var nodePosX = posX + CLASS_PADDING + VERTEX_BOX_CENTER + i * VERTEX_BOX;
                        var nodePosY = middle;
                        if (leftColumns[i] is CellVertex cellVertex)
                            nodes.Add(cellVertex.FormatCellVertex(nodePosX, nodePosY));
                        else if (leftColumns[i] is RangeVertex rangeVertex)
                            nodes.Add(rangeVertex.FormatRangeVertexSmall(nodePosX, nodePosY));
                    }
                }
                else
                {
                    for (var i = 0; i < leftColumns.Count; i++)
                    {
                        var nodePosX = posX + CLASS_PADDING + VERTEX_BOX_CENTER + i * VERTEX_BOX;
                        var nodePosY = currentRowY;
                        if (leftColumns[i] is CellVertex cellVertex)
                            nodes.Add(cellVertex.FormatCellVertex(nodePosX, nodePosY));
                        else if (leftColumns[i] is RangeVertex rangeVertex)
                            nodes.Add(rangeVertex.FormatRangeVertexSmall(nodePosX, nodePosY));
                    }

                    currentRowY += VERTEX_BOX;
                }
            }

            var classNode = new NodeViewModel
            {
                ID = @class.Name,
                Content = @class,
                ContentTemplate = new DataTemplate(),
                UnitWidth = width,
                UnitHeight = currentRowY - VERTEX_BOX + CLASS_PADDING,
                Pivot = new Point(0, 0),
                OffsetX = posX,
                OffsetY = posY,
                Shape = _classShape,
                Annotations = new AnnotationCollection
                {
                    new AnnotationEditorViewModel
                    {
                        Offset = new Point(0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        Content = @class.Name,
                        ViewTemplate = _classLabelTemplate
                    }
                },
                ZIndex = int.MinValue,
                ShapeStyle = GetNodeShapeStyle(new SolidColorBrush(@class.Color.ToMColor()))
            };
            if (@class.IsStaticClass)
            {
                ((AnnotationCollection)classNode.Annotations).Add(new AnnotationEditorViewModel
                {
                    Offset = new Point(1, 0),
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Bottom,
                    Content = "<<static>>",
                    ViewTemplate = _staticLabelTemplate
                });
            }

            SetNodeConstraints(classNode);

            nodes.Add(classNode);
            return (group, nextPosX);
        }

        // returns as list of vertex lists, each list is a row
        private static IEnumerable<List<Vertex>> LayoutGraph(Class @class)
        {
            var classVertices = @class.Vertices;
            if (@class.OutputVertex == null)
            {
                foreach (var vertex in classVertices)
                    yield return new List<Vertex> {vertex};
            }
            else
            {
                var vertexQueue = new Queue<Vertex>(classVertices);
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
                               vertexQueue.Peek() is CellVertex cellVertex && cellVertex.NodeType == NodeType.Constant)
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

        public static ConnectorViewModel FormatEdge(this CellVertex from, Vertex to, bool connectToExternal = true)
        {
            var connector = new ConnectorViewModel
            {
                SourceNode = from.Node,
                Segments = new ObservableCollection<IConnectorSegment>
                {
                    new StraightSegment()
                },
                Constraints = ConnectorConstraints.Default & ~ConnectorConstraints.Selectable,
                ZIndex = -1
            };

            if (to.IsExternal && !connectToExternal)
            {
                var length = Math.Min(VERTEX_BOX, from.Node.UnitWidth - 5);
                connector.TargetPoint = new Point(from.Node.OffsetX + length, from.Node.OffsetY - length);
                connector.ConnectorGeometryStyle = _externalConnectorGeometryStyle;
                connector.TargetDecoratorStyle = _externalTargetDecoratorStyle;
            }
            else
            {
                connector.TargetNode = to.Node;
                connector.ConnectorGeometryStyle = _connectorGeometryStyle;
                connector.TargetDecoratorStyle = _targetDecoratorStyle;
            }

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