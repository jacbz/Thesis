using System;
using System.Collections.Generic;
using System.Linq;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class LabelGenerator
    {
        private static Dictionary<(int row, int col), Region> _regionDictionary;
        public static int HorizontalMergingRange { get; set; }
        public static int VerticalMergingRange { get; set; }
        public static bool MergeOutputFieldLabels { get; set; }
        public static int HeaderAssociationRange { get; set; }
        public static int AttributeAssociationRange { get; set; }


        public static List<Region> CreateRegions(List<CellVertex> cellVertices)
        {
            _regionDictionary = new Dictionary<(int row, int col), Region>();
            foreach (var cellVertex in cellVertices)
            {
                if (cellVertex.NodeType != NodeType.None)
                    _regionDictionary.Add(cellVertex.Address, new DataRegion(cellVertex));
                else if (cellVertex.NodeType == NodeType.None && cellVertex.CellType == CellType.Text)
                    _regionDictionary.Add(cellVertex.Address, new LabelRegion(cellVertex));
            }

            var regions = new List<Region>();

            var dataRegionsList = _regionDictionary.Values.OfType<DataRegion>().ToList();

            // merge data regions
            bool didSomething = true;
            int iterations = 0;
            while (didSomething)
            {
                iterations++;
                didSomething = false;

                for (int i = 0; i < dataRegionsList.Count; i++)
                {
                    for(int j = i + 1; j < dataRegionsList.Count; j++)
                    {
                        var regionsToRemove = dataRegionsList[i].MergeIfAllowed(dataRegionsList[j]);
                        if (regionsToRemove.Count > 0)
                        {
                            didSomething = true;
                            dataRegionsList.RemoveAll(dr => regionsToRemove.Contains(dr));
                        }
                    }
                }
            }

            Logger.Log(LogItemType.Info,
                $"Discovered {dataRegionsList.Count} data regions after {iterations} iterations");

            regions.AddRange(dataRegionsList);

            var labelRegionList = _regionDictionary.Values.OfType<LabelRegion>().ToList();

            // merge data regions
            didSomething = true;
            iterations = 0;
            while (didSomething)
            {
                iterations++;
                didSomething = false;

                for (int i = 0; i < labelRegionList.Count; i++)
                {
                    for (int j = i + 1; j < labelRegionList.Count; j++)
                    {
                        var regionsToRemove = labelRegionList[i].MergeIfAllowed(labelRegionList[j]);
                        if (regionsToRemove.Count > 0)
                        {
                            didSomething = true;
                            labelRegionList.RemoveAll(dr => regionsToRemove.Contains(dr));
                        }
                    }
                }
            }
            Logger.Log(LogItemType.Info,
                $"Discovered {labelRegionList.Count} label regions after {iterations} iterations");

            // assign DataRegions for each LabelRegion
            foreach (var labelRegion in labelRegionList)
            {
                if (labelRegion.Type == LabelRegionType.Header)
                {
                    var initialRow = labelRegion.BottomRight.row;
                    for (int row = initialRow; row <= initialRow + HeaderAssociationRange; row++)
                    {
                        for (int column = labelRegion.TopLeft.column; column <= labelRegion.BottomRight.column; column++)
                        {
                            if (_regionDictionary.TryGetValue((row, column), out var region) &&
                                region is DataRegion dataRegion)
                                dataRegion.LabelRegions.Add(labelRegion);
                        }
                    }
                }
                else if (labelRegion.Type == LabelRegionType.Attribute)
                {
                    var initialColumn = labelRegion.BottomRight.column;
                    for (int column = initialColumn; column <= initialColumn + AttributeAssociationRange; column++)
                    {
                        for (int row = labelRegion.TopLeft.row; row <= labelRegion.BottomRight.row; row++)
                        {
                            if (_regionDictionary.TryGetValue((row, column), out var region) &&
                                region is DataRegion dataRegion)
                                dataRegion.LabelRegions.Add(labelRegion);
                        }
                    }
                }
            }

            regions.AddRange(labelRegionList);

            _regionDictionary = null;
            return regions;
        }

        public static void GenerateLabelsFromRegions(IEnumerable<CellVertex> cellVertices)
        {
            foreach (var vertex in cellVertices)
            {
                // do not generate name for vertices which already have a name
                if (!string.IsNullOrEmpty(vertex.Name)) continue;

                if (!(vertex.Region is DataRegion dataRegion)) continue;
                var headers = dataRegion.LabelRegions
                    .Where(lr => lr.Type == LabelRegionType.Header)
                    .SelectMany(lr => lr.Cells)
                    .Where(cell => cell.Address.col == vertex.Address.col)
                    .OrderBy(cell => cell.Address.row)
                    .Select(cell => cell.DisplayValue.ToPascalCase())
                    .Where(s => !string.IsNullOrWhiteSpace(s) && s != "_")
                    .ToList();
                var attributes = dataRegion.LabelRegions
                    .Where(lr => lr.Type == LabelRegionType.Attribute)
                    .SelectMany(lr => lr.Cells)
                    .Where(cell => cell.Address.row == vertex.Address.row)
                    .OrderBy(cell => cell.Address.col)
                    .Select(cell => cell.DisplayValue.ToPascalCase())
                    .Where(s => !string.IsNullOrWhiteSpace(s) && s != "_")
                    .ToList();
                var name = string.Join("_", headers.Concat(attributes));
                if (!string.IsNullOrWhiteSpace(name))
                    vertex.Name = name.LowerFirstCharacter();
            }
        }

        public abstract class Region
        {
            public CellVertex[] Cells { get; private set; }

            public (int row, int column) TopLeft { get; private set; }
            public (int row, int column) BottomRight { get; private set; }

            protected RegionOrientation Orientation => BottomRight.row - TopLeft.row >= BottomRight.column - TopLeft.column
                ? RegionOrientation.Vertical
                : RegionOrientation.Horizontal;

            public enum RegionOrientation
            {
                Horizontal, Vertical
            }


            protected Region(CellVertex cellVertex)
            {
                cellVertex.Region = this;
                Cells = new CellVertex[1];
                Cells[0] = cellVertex;
                TopLeft = BottomRight = cellVertex.Address;
            }

            // is not commutative
            private bool LiesWithin((int row, int column) topLeft, (int row, int column) bottomRight)
            {
                return TopLeft.row >= topLeft.row && TopLeft.column >= topLeft.column
                                                   && BottomRight.row <= bottomRight.row
                                                   && BottomRight.column <= bottomRight.column;
            }

            // is commutative
            private int HorizontalDistanceTo(Region otherRegion)
            {
                if (BottomRight.column <= otherRegion.TopLeft.column)
                    return otherRegion.TopLeft.column - BottomRight.column;
                if (otherRegion.BottomRight.column <= TopLeft.column)
                    return TopLeft.column - otherRegion.BottomRight.column;
                return -1;
            }

            // is commutative
            private int VerticalDistanceTo(Region otherRegion)
            {
                if (BottomRight.row <= otherRegion.TopLeft.row)
                    return otherRegion.TopLeft.row - BottomRight.row;
                if (otherRegion.BottomRight.row <= TopLeft.row)
                    return TopLeft.row - otherRegion.BottomRight.row;
                return -1;
            }

            // returns empty list if can't merge
            public List<Region> MergeIfAllowed(Region otherRegion)
            {
                var regionsToRemove = new List<Region>();
                if (GetType() != otherRegion.GetType())
                    return regionsToRemove;

                if (VerticalDistanceTo(otherRegion) > VerticalMergingRange || HorizontalDistanceTo(otherRegion) > HorizontalMergingRange)
                    return regionsToRemove;

                var newTopLeftRow = Math.Min(TopLeft.row, otherRegion.TopLeft.row);
                var newTopLeftColumn = Math.Min(TopLeft.column, otherRegion.TopLeft.column);
                var newBottomRightRow = Math.Max(BottomRight.row, otherRegion.BottomRight.row);
                var newBottomRightColumn = Math.Max(BottomRight.column, otherRegion.BottomRight.column);

                var regionsToMergeTogether = new List<Region> {this, otherRegion};

                var newDictionaryEntries = new Dictionary<(int row, int col), Region>();
                for (int i = newTopLeftRow; i <= newBottomRightRow; i++)
                {
                    for (int j = newTopLeftColumn; j <= newBottomRightColumn; j++)
                    {
                        // if we find another label region that lies within the new region, merge that too, else abort
                        if (_regionDictionary.TryGetValue((i, j), out var thirdRegion))
                        {
                            if (thirdRegion == this)
                                continue;

                            if (MergeOutputFieldLabels && this is LabelRegion && thirdRegion is DataRegion dataRegion
                                && dataRegion.Cells.All(c => c.NodeType == NodeType.OutputField))
                            {
                               regionsToMergeTogether.Add(thirdRegion);
                            }
                            // abort if would merge a region of different type
                            else if (thirdRegion.GetType() != GetType())
                            {
                                return regionsToRemove;
                            }

                            if (thirdRegion.LiesWithin((newTopLeftRow, newTopLeftColumn),
                                (newBottomRightRow, newBottomRightColumn)))
                            {
                                regionsToMergeTogether.Add(thirdRegion);
                            }
                            else
                            {
                                return regionsToRemove;
                            }
                        }
                        newDictionaryEntries.Add((i,j), this);
                    }
                }

                foreach (var kvp in newDictionaryEntries)
                {
                    if (_regionDictionary.ContainsKey(kvp.Key))
                        _regionDictionary[kvp.Key] = kvp.Value;
                    else
                        _regionDictionary.Add(kvp.Key, kvp.Value);
                }

                Cells = regionsToMergeTogether.SelectMany(r => r.Cells).Distinct().ToArray();
                Cells.ForEach(c => c.Region = this);
                TopLeft = (newTopLeftRow, newTopLeftColumn);
                BottomRight = (newBottomRightRow, newBottomRightColumn);
                
                return regionsToMergeTogether.Skip(1).ToList();
            }

            public override string ToString()
            {
                return $"[{string.Join(",", Cells.Select(c => c.StringAddress))}]";
            }
        }

        public class LabelRegion : Region
        {
            public LabelRegionType Type => Orientation == RegionOrientation.Horizontal
                ? LabelRegionType.Header
                : LabelRegionType.Attribute;
            public LabelRegion(CellVertex cellVertex) : base(cellVertex)
            {
            }
        }

        public enum LabelRegionType
        {
            Header, Attribute
        }

        public class DataRegion : Region
        {
            public HashSet<LabelRegion> LabelRegions { get; }
            public DataRegion(CellVertex cellVertex) : base(cellVertex)
            {
                LabelRegions = new HashSet<LabelRegion>();
            }
        }
    }
}
