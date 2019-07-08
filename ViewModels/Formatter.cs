using Syncfusion.UI.Xaml.Diagram;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Shapes;
using Point = System.Windows.Point;

namespace Thesis.ViewModels
{
    // Extension methods to layout and format objects
    public static class Formatter
    {
        public static string GetColor(this Vertex vertex)
        {
            switch (vertex.Type)
            {
                case CellType.Formula:
                    if (vertex.IsOutputField) return "#ff0000";
                    return "#8e44ad";
                default:
                    return "#2980b9";
            }
        }

        public static NodeViewModel FormatNode(this Vertex vertex, Graph graph)
        {
            return new NodeViewModel()
            {
                ID = vertex.Address,
                Content = vertex,
                ContentTemplate = new DataTemplate(),
                UnitWidth = 40,
                UnitHeight = 40,
                OffsetX = graph.PopulatedColumns.IndexOf(vertex.CellIndex[1]) * 60 + 40,
                OffsetY = graph.PopulatedRows.IndexOf(vertex.CellIndex[0]) * 60 + 40,
                Shape = Application.Current.Resources[
                    vertex.Type == CellType.Formula 
                        ? (vertex.IsOutputField ? "Trapezoid" : "Heptagon") 
                        : "Ellipse"
                        ],
                ShapeStyle = Application.Current.Resources[
                    vertex.Type == CellType.Formula
                        ? (vertex.IsOutputField ? "OutputStyle" : "FormulaStyle")
                        : "ValueStyle"] as Style,
                Annotations = new AnnotationCollection()
                {
                    new AnnotationEditorViewModel()
                    {
                        Offset = new Point(0, 0),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        Content = vertex.Address
                    }
                },
                Constraints = NodeConstraints.Default & ~NodeConstraints.Rotatable & ~NodeConstraints.Delete & ~NodeConstraints.Connectable
            };
        }

        public static ConnectorViewModel FormatConnector(this Vertex from, Vertex to)
        {
            return new ConnectorViewModel()
            {
                SourceNodeID = from.Address,
                TargetNodeID = to.Address,
                ConnectorGeometryStyle = Application.Current.Resources["ConnectorGeometryStyle"] as Style,
                TargetDecoratorStyle = Application.Current.Resources["TargetDecoratorStyle"] as Style,
                Segments = new ObservableCollection<IConnectorSegment>()
                {
                    new StraightSegment()
                },
                Constraints = ConnectorConstraints.Default & ~ConnectorConstraints.Selectable
            };
        }
    }
}
