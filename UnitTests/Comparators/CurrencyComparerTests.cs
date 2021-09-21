// ReSharper disable ReturnValueOfPureMethodIsNotUsed

using NUnit.Framework;
using System;
using System.Collections.Generic;
using UITests.PageModel.Shared.Comparators;

namespace UnitTests.Comparators
{
    public class CurrencyComparerTests
    {
        [TestCase("10 USD", "10 USD")]
        public void Compare_WhenEqualArguments_ReturnsZero(object x, object y)
        {
            var comparer = new CurrencyComparer();
            Assert.That(comparer.Compare(x, y), Is.EqualTo(0));
        }

        [TestCase("10.10 USD", "10.00 USD")]
        public void Compare_WhenFirstArgumentIsBigger_ReturnsOne(object x, object y)
        {
            var comparer = new CurrencyComparer();
            Assert.That(comparer.Compare(x, y), Is.EqualTo(1));
        }

        [TestCase("10.00 USD", "10.10 USD")]
        public void Compare_WhenFirstArgumentIsLess_ReturnsMinusOne(object x, object y)
        {
            var comparer = new CurrencyComparer();
            Assert.That(comparer.Compare(x, y), Is.EqualTo(-1));
        }

        [TestCase(null, "10 USD")]
        [TestCase("10 USD", "")]
        public void Compare_WhenInputFormatIsInvalid_ThrowsArgumentException(object x, object y)
        {
            var comparer = new CurrencyComparer();
            Assert.Throws<ArgumentException>(() => comparer.Compare(x, y));
        }

        [TestCase("A USD", "10 USD")]
        public void Compare_WhenValueIsNotNumeric_ThrowsFormatException(object x, object y)
        {
            var comparer = new CurrencyComparer();
            Assert.Throws<FormatException>(() => comparer.Compare(x, y));
        }

        [Test]
        public void UseCase_WhenSmallValueGoesFirst_ListIsOrderedAscending()
        {
            var list = new List<string> { "1 USD", "10 USD", "10.78 USD", "500.00 USD" };
            Assert.That(list, Is.Ordered.Ascending.Using(new CurrencyComparer()));
        }

        [Test]
        public void UseCase_WhenBigValueGoesFirst_ListIsOrderedDescending()
        {
            var list = new List<string> { "500.00 USD", "10.78 USD", "10 USD", "1 USD" };
            Assert.That(list, Is.Ordered.Descending.Using(new CurrencyComparer()));
        }
    }
}