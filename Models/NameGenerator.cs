// Thesis - An Excel to code converter
// Copyright (C) 2019 Jacob Zhang
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Linq;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class NameGenerator
    {
        private static Dictionary<(int row, int col), Region> _regionDictionary;
        public static int HorizontalMergingRange { get; set; }
        public static int VerticalMergingRange { get; set; }
        public static bool MergeOutputFieldLabels { get; set; }
        public static int HeaderAssociationRange { get; set; }
        public static int AttributeAssociationRange { get; set; }


        public static List<Region> CreateRegions(List<CellVertex> cellVertices)
        {
            var logItem = Logger.Log(LogItemType.Info, "Create label regions...", true);
            ;
            _regionDictionary = new Dictionary<(int row, int col), Region>();
            foreach (var cellVertex in cellVertices)
            {
                if (cellVertex.Classification != Classification.None)
                    _regionDictionary.Add(cellVertex.Address, new DataRegion(cellVertex));
                else if (cellVertex.Classification == Classification.None && cellVertex.CellType == CellType.Text)
                    _regionDictionary.Add(cellVertex.Address, new LabelRegion(cellVertex));
            }

            var regions = _regionDictionary.Values.ToList();

            // merge data regions

            for (int i = 0; i < regions.Count; i++)
            {
                for(int j = 0; j < regions.Count; j++)
                {
                    if (i == j) continue;
                    var mergedRegions = regions[i].Merge(regions[j]);
                    if (mergedRegions.Count > 0)
                    {;
                        regions.RemoveAll(dr => mergedRegions.Contains(dr));
                        j--;
                    }
                }
            }
            logItem.AppendElapsedTime();
            Logger.Log(LogItemType.Info,
                $"Discovered {regions.Count} regions");
            
            var dataRegionsList = _regionDictionary.Values.OfType<DataRegion>().ToList();
            var labelRegionList = _regionDictionary.Values.OfType<LabelRegion>().ToList();
            // assign DataRegions for each LabelRegion
            foreach (var labelRegion in labelRegionList)
            {
                foreach (var dataRegion in dataRegionsList)
                {
                    if (labelRegion.Type == LabelRegionType.Header)
                    {
                        if (labelRegion.BottomRight.row < dataRegion.TopLeft.row &&
                            dataRegion.VerticalDistanceTo(labelRegion) <= HeaderAssociationRange)
                        {
                            dataRegion.LabelRegions.Add(labelRegion);
                        }
                    }
                    else
                    {
                        if (labelRegion.BottomRight.column < dataRegion.TopLeft.column &&
                            dataRegion.HorizontalDistanceTo(labelRegion) <= AttributeAssociationRange)
                        {
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
                if (vertex.Name != null) continue;

                if (!(vertex.Region is DataRegion dataRegion)) continue;
                var headers = dataRegion.LabelRegions
                    .Where(lr => lr.Type == LabelRegionType.Header)
                    .SelectMany(lr => lr.Cells)
                    .Where(cell => cell.Address.col == vertex.Address.col)
                    .OrderBy(cell => cell.Address.row)
                    .ToArray();
                var attributes = dataRegion.LabelRegions
                    .Where(lr => lr.Type == LabelRegionType.Attribute)
                    .SelectMany(lr => lr.Cells)
                    .Where(cell => cell.Address.row == vertex.Address.row)
                    .OrderBy(cell => cell.Address.col)
                    .ToArray();
                vertex.Name = new Name(attributes, headers, vertex.StringAddress);
            }
        }

        public abstract class Region
        {
            public CellVertex[] Cells { get; private set; }
            public (int row, int column) TopLeft { get; private set; }
            public (int row, int column) BottomRight { get; private set; }

            protected RegionOrientation Orientation => BottomRight.row - TopLeft.row > BottomRight.column - TopLeft.column
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
            public int HorizontalDistanceTo(Region otherRegion)
            {
                if (BottomRight.column <= otherRegion.TopLeft.column)
                    return otherRegion.TopLeft.column - BottomRight.column;
                if (otherRegion.BottomRight.column <= TopLeft.column)
                    return TopLeft.column - otherRegion.BottomRight.column;
                return -1;
            }

            // is commutative
            public int VerticalDistanceTo(Region otherRegion)
            {
                if (BottomRight.row <= otherRegion.TopLeft.row)
                    return otherRegion.TopLeft.row - BottomRight.row;
                if (otherRegion.BottomRight.row <= TopLeft.row)
                    return TopLeft.row - otherRegion.BottomRight.row;
                return -1;
            }

            // returns list of regions that were merged (if aborted, list is empty)
            public List<Region> Merge(Region otherRegion)
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
                                && dataRegion.Cells.All(c => c.Classification == Classification.OutputField))
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
