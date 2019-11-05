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
using Syncfusion.XlsIO;

namespace Thesis.Models.VertexTypes
{
    public class RangeVertex : Vertex
    {
        public RangeType Type { get; }

        public IRange[] CellsInRange { get; }
        public (int row, int column) StartAddress { get; }
        public (int row, int column) EndAddress { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        protected Dictionary<(int Row, int Column), IRange> AddressToCellDictionary;

        public RangeVertex(IRange[] cellsInRange, string name, string address)
        {
            CellsInRange = cellsInRange;
            Name = new Name(name, address);
            StringAddress = address;
            AddressToCellDictionary = new Dictionary<(int Row, int Column), IRange>();

            if (cellsInRange.Length == 0)
            {
                Type = RangeType.Empty;
                return;
            }

            int minRow = int.MaxValue;
            int minColumn = int.MaxValue;
            int maxRow = int.MinValue;
            int maxColumn = int.MinValue;

            foreach (IRange iRange in CellsInRange)
            {
                minRow = Math.Min(iRange.Row, minRow);
                minColumn = Math.Min(iRange.Column, minColumn);
                maxRow = Math.Max(iRange.LastRow, maxRow);
                maxColumn = Math.Max(iRange.LastColumn, maxColumn);
                AddressToCellDictionary.Add((iRange.Row, iRange.Column), iRange);
            }

            StartAddress = (minRow, minColumn);
            EndAddress = (maxRow, maxColumn);
            RowCount = maxRow - minRow + 1;
            ColumnCount = maxColumn - minColumn + 1;
            Type = minRow == maxRow
                ? minColumn == maxColumn
                  ? RangeType.Single
                  : RangeType.Row
                : minColumn == maxColumn
                    ? RangeType.Column
                    : RangeType.Matrix;
        }

        public (int row, int column)[] GetAddressTuples()
        {
            return CellsInRange.Select(c => (c.Row, c.Column)).ToArray();
        }

        public (string sheetName, string address)[] GetAddresses()
        {
            return CellsInRange.Select(c => (IsExternal ? ExternalWorksheetName : null, c.AddressLocal)).ToArray();
        }

        public IEnumerable<int> GetPopulatedRows()
        {
            return Enumerable.Range(StartAddress.row, StartAddress.row + RowCount - 1);
        }

        public IEnumerable<int> GetPopulatedColumns()
        {
            return Enumerable.Range(StartAddress.column, StartAddress.column + RowCount - 1);
        }

        public CellVertex GetSingleElement()
        {
            return new CellVertex(AddressToCellDictionary[(StartAddress.row, StartAddress.column)]);
        }

        public CellVertex[] GetColumnArray()
        {
            CellVertex[] array = new CellVertex[RowCount];
            for (int i = 0; i < RowCount; i++)
            {
                array[i] = new CellVertex(AddressToCellDictionary[(i + StartAddress.row, StartAddress.column)]);
            }
            return array;
        }

        public CellVertex[] GetRowArray()
        {
            CellVertex[] array = new CellVertex[ColumnCount];
            for (int i = 0; i < ColumnCount; i++)
            {
                array[i] = new CellVertex(AddressToCellDictionary[(StartAddress.row, i + StartAddress.column)]);
            }
            return array;
        }
        
        public CellVertex[][] GetMatrixArray()
        {
            CellVertex[][] matrix = new CellVertex[RowCount][];
            for (int j = 0; j < RowCount; j++)
            {
                matrix[j] = new CellVertex[ColumnCount];
                for (int k = 0; k < ColumnCount; k++)
                {
                    matrix[j][k] = new CellVertex(AddressToCellDictionary[(j + StartAddress.row, k + StartAddress.column)]);
                }
            }
            return matrix;
        }

        public enum RangeType
        {
            Empty,
            Single,
            Column,
            Row,
            Matrix
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
