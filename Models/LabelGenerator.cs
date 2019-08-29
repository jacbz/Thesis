using System.Collections.Generic;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class LabelGenerator
    {
        public static void GenerateLabels(List<CellVertex> cellVertices)
        {
            var logItem = Logger.Log(LogItemType.Info, "Generating labels...", true);
            // Create labels
            Dictionary<(int row, int col), Label> labelDictionary = new Dictionary<(int row, int col), Label>();

            foreach (var cell in cellVertices)
            {
                Label label = new Label(cell);
                if (cell.CellType == CellType.Unknown && cell.NodeType == NodeType.None)
                {
                    label.Type = LabelType.None;
                }
                else if (cell.Children.Count == 0 && cell.Parents.Count == 0 && cell.CellType == CellType.Text)
                {
                    if (labelDictionary.TryGetValue((cell.Address.row - 1, cell.Address.col), out Label labelAbove)
                        && (labelAbove.Type == LabelType.Attribute || labelAbove.Type == LabelType.Header))
                    {
                        label.Type = LabelType.Attribute;
                        label.Text = cell.DisplayValue;
                        labelAbove.Type = LabelType.Attribute;
                    }
                    else
                    {
                        label.Type = LabelType.Header;
                        label.Text = cell.DisplayValue;
                    }
                }
                else
                {
                    label.Type = LabelType.Data;
                }

                cell.Label = label;
                labelDictionary.Add((cell.Address.row, cell.Address.col), label);
            }

            // assign attributes and headers for each data type
            foreach (var cell in cellVertices)
            {
                if (cell.Label.Type != LabelType.Data) continue;

                (int row, int col) currentPos = cell.Address;

                // add attributes
                bool foundAttribute = false;
                int distanceToAttribute = 0;
                // a list that stores how far all attributes are to the vertex. e.g. attribute in 2,3, vertex in 8: [5,6]
                List<int> distancesToAttribute = new List<int>();
                while (currentPos.col-- > 1)
                {
                    var currentLabel = labelDictionary[currentPos];
                    if (foundAttribute && currentLabel.Type != LabelType.Attribute) break;

                    distanceToAttribute++;
                    if (currentLabel.Type == LabelType.Attribute)
                    {
                        foundAttribute = true;
                        cell.Label.Attributes.Add(currentLabel);
                        distancesToAttribute.Add(distanceToAttribute);
                    }
                }

                // add headers
                currentPos = cell.Address;
                if (!foundAttribute)
                {
                    // no attributes, use first header on the top
                    while (currentPos.row-- > 1)
                    {
                        var currentLabel = labelDictionary[currentPos];
                        if (currentLabel.Type == LabelType.Header)
                        {
                            cell.Label.Headers.Add(currentLabel);
                            break;
                        }
                    }
                }
                else
                {
                    // keep adding headers, until there is no attribute to the left or left bottom with the exact distance
                    bool foundHeader = false;
                    while (currentPos.row-- > 1)
                    {
                        var currentLabel = labelDictionary[currentPos];

                        bool anyAttributeDistanceMatch = false;
                        foreach (int dist in distancesToAttribute)
                        {
                            if (labelDictionary[(currentPos.row, currentPos.col - dist)].Type == LabelType.Attribute ||
                                labelDictionary[(currentPos.row + 1, currentPos.col - dist)].Type == LabelType.Attribute)
                                anyAttributeDistanceMatch = true;
                        }

                        if (!anyAttributeDistanceMatch)
                            break;
                        if (foundHeader && currentLabel.Type != LabelType.Header)
                            break;

                        if (currentLabel.Type == LabelType.Header)
                        {
                            foundHeader = true;
                            cell.Label.Headers.Add(currentLabel);
                        }
                    }
                }

                // do not override name if name was already assigned, e.g. per named range
                if (string.IsNullOrEmpty(cell.VariableName))
                    cell.Label.GenerateVariableName();
            }
            logItem.AppendElapsedTime();
        }
    }
}
