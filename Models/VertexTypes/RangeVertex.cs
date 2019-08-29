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
        protected Dictionary<(int Row, int Column), IRange> AddressToCellDictionary;
        protected (int row, int column) StartAddress;
        protected (int row, int column) EndAddress;
        protected int RowCount;
        protected int ColumnCount;

        public RangeVertex(IRange[] cellsInRange, string addressOrName)
        {
            CellsInRange = cellsInRange;
            VariableName = addressOrName.MakeNameVariableConform();
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
            return VariableName;
        }
    }
}
