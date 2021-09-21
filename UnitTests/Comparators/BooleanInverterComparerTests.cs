using NUnit.Framework;
using System.Collections.Generic;
using UITests.PageModel.Shared.Comparators;

namespace UnitTests.Comparators
{
    public class BooleanInverterComparerTests
    {
        [TestCase(true, true)]
        [TestCase(false, false)]
        public void Compare_WhenEqualArguments_ReturnsZero(object x, object y)
        {
            var comparer = new BooleanInverterComparer();
            Assert.That(comparer.Compare(x, y), Is.EqualTo(0));
        }

        [Test]
        public void Compare_WhenFirstArgumentIsFalse_ReturnsOne()
        {
            var comparer = new BooleanInverterComparer();
            Assert.That(comparer.Compare(false, true), Is.EqualTo(1));
        }

        [Test]
        public void Compare_WhenFirstArgumentIsTrue_ReturnsMinusOne()
        {
            var comparer = new BooleanInverterComparer();
            Assert.That(comparer.Compare(true, false), Is.EqualTo(-1));
        }

        [Test]
        public void UseCase_WhenTrueGoesFirst_ListIsOrderedAscending()
        {
            var list = new List<bool> { true, true, false };
            Assert.That(list, Is.Ordered.Ascending.Using(new BooleanInverterComparer()));
        }

        [Test]
        public void UseCase_WhenFalseGoesFirst_ListIsOrderedDescending()
        {
            var list = new List<bool> { false, false, true };
            Assert.That(list, Is.Ordered.Descending.Using(new BooleanInverterComparer()));
        }
    }
}
