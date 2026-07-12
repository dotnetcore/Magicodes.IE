
using System;

namespace Magicodes.IE.IO
{
    internal static class CellRefHelper
    {
        internal static int CellRefToCol(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return -1;
            int i = 0;
            while (i < cellRef.Length && IsAsciiLetter(cellRef[i])) i++;
            if (i == 0 || i > 3) return -1;
            int colNum = 0;
            for (int k = 0; k < i; k++)
                colNum = colNum * 26 + (char.ToUpperInvariant(cellRef[k]) - 'A' + 1);
            return colNum is >= 1 and <= 16384 ? colNum - 1 : -1;
        }

        internal static int CellRefToRow(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return -1;
            int i = 0;
            while (i < cellRef.Length && IsAsciiLetter(cellRef[i])) i++;
            if (i >= cellRef.Length) return -1;
            int row = 0;
            for (; i < cellRef.Length; i++)
            {
                char c = cellRef[i];
                if (c < '0' || c > '9') return -1;
                row = checked(row * 10 + c - '0');
                if (row > 1048576) return -1;
            }
            return row is >= 1 and <= 1048576 ? row - 1 : -1;
        }

        internal static bool IsCellReference(string? cellRef)
        {
            if (cellRef is null || cellRef.Length == 0) return false;
            int i = 0;
            while (i < cellRef.Length && IsAsciiLetter(cellRef[i])) i++;
            if (i == 0 || i > 3 || i == cellRef.Length) return false;
            for (; i < cellRef.Length; i++)
            {
                char c = cellRef[i];
                if (c < '0' || c > '9') return false;
            }
            return CellRefToCol(cellRef) >= 0 && CellRefToRow(cellRef) >= 0;
        }

        internal static bool IsCellRange(string? range)
        {
            if (range is null || range.Length == 0) return false;
            int separator = range.IndexOf(':');
            return separator > 0
                && separator == range.LastIndexOf(':')
                && IsCellReference(range.Substring(0, separator))
                && IsCellReference(range.Substring(separator + 1));
        }

        private static bool IsAsciiLetter(char c) => c is >= 'A' and <= 'Z' or >= 'a' and <= 'z';
    }
}