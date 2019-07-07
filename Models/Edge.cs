using Syncfusion.UI.Xaml.Diagram;
using System.Windows;

namespace Thesis
{
    public class Edge
    {
        public Vertex From { get; set; }
        public Vertex To { get; set; }
        public string Label { get; set; }
        
        public Edge(Vertex from, Vertex to, string label)
        {
            to.Parents.Add(from.Address);
            from.Children.Add(to.Address);

            From = from;
            To = to;
            Label = label;
        }
        public override string ToString()
        {
            return $"{From.Address}->{To.Address} ({Label})";
        }
    }
}
