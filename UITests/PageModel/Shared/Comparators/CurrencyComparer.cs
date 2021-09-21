using System;
using System.Collections;
using System.ComponentModel;
// ReSharper disable InconsistentNaming
// ReSharper disable UnusedMember.Local

namespace UITests.PageModel.Shared.Comparators
{
    // Compare currency values represented as string in a format "78.00 USD", "400.00 USD".
    public class CurrencyComparer : IComparer
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

            var xUnit = (Currency)Enum.Parse(typeof(Currency), xValues[1]);
            var yUnit = (Currency)Enum.Parse(typeof(Currency), yValues[1]);

            var unitsEquality = xUnit.CompareTo(yUnit);
            if (unitsEquality != 0)
            {
                return unitsEquality;
            }

            var xSize = Convert.ToDouble(xValues[0]);
            var ySize = Convert.ToDouble(yValues[0]);
            return xSize.CompareTo(ySize);
        }

        private enum Currency
        {
            [Description("USD")] USD
        }
    }
}