using System;
using System.Collections;
using System.ComponentModel;
// ReSharper disable InconsistentNaming
// ReSharper disable UnusedMember.Local

namespace UITests.PageModel.Shared.Comparators
{
    // Compare document sizes represented as string in a format "5 B", "2 KB".
    public class DocumentSizeComparer : IComparer
    {
        private const StringSplitOptions Options = StringSplitOptions.RemoveEmptyEntries;
        private readonly string[] _separator = { " " };

        public int Compare(object x, object y)
        {
            var xString = (string)x;
            var yString = (string)y;

            if (string.IsNullOrEmpty(xString) || string.IsNullOrEmpty(yString))
            {
                throw new ArgumentException("Input string was empty.");
            }

            var xValues = xString.Split(_separator, Options);
            var yValues = yString.Split(_separator, Options);

            if (xValues.Length != 2 || yValues.Length != 2)
            {
                throw new ArgumentException("Input string was not in a correct format.");
            }

            var xUnit = (Size)Enum.Parse(typeof(Size), xValues[1]);
            var yUnit = (Size)Enum.Parse(typeof(Size), yValues[1]);

            var unitsEquality = xUnit.CompareTo(yUnit);
            if (unitsEquality != 0)
            {
                return unitsEquality;
            }

            var xSize = Convert.ToDouble(xValues[0]);
            var ySize = Convert.ToDouble(yValues[0]);
            return xSize.CompareTo(ySize);
        }

        private enum Size
        {
            [Description("B")] B,
            [Description("KB")] KB,
            [Description("MB")] MB,
            [Description("GB")] GB,
            [Description("PB")] PB
        }
    }
}
