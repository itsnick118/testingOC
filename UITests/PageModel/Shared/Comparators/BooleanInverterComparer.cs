using System.Collections;

namespace UITests.PageModel.Shared.Comparators
{
    // Changes boolean comparison result so <False> is bigger than <True>
    public class BooleanInverterComparer : IComparer
    {
        public int Compare(object x, object y)
        {
            var a = x != null && (bool)x;
            var b = y != null && (bool)y;

            var result = a.CompareTo(b);
            return result == 0 ? 0 : 0 - result;
        }
    }
}
