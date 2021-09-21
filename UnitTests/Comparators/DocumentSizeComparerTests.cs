using NUnit.Framework;
using System;
using System.Collections.Generic;
using UITests.PageModel.Shared.Comparators;

// ReSharper disable ReturnValueOfPureMethodIsNotUsed
namespace UnitTests.Comparators
{
    public class DocumentSizeComparerTests
    {
        [TestCase("1 B", "1 B")]
        [TestCase("5 KB", "5 KB")]
        [TestCase("10 MB", "10 MB")]
        [TestCase("2 GB", "2 GB")]
        [TestCase("1 PB", "1 PB")]
        public void Compare_WhenEqualArguments_ReturnsZero(object x, object y)
        {
            var comparer = new DocumentSizeComparer();
            Assert.That(comparer.Compare(x, y), Is.EqualTo(0));
        }

        [TestCase("2 B", "1 B")]
        [TestCase("1 KB", "10 B")]
        [TestCase("2 MB", "5 KB")]
        [TestCase("2.1 MB", "2 MB")]
        public void Compare_WhenFirstArgumentIsBigger_ReturnsOne(object x, object y)
        {
            var comparer = new DocumentSizeComparer();
            Assert.That(comparer.Compare(x, y), Is.EqualTo(1));
        }

        [TestCase("1 B", "2 B")]
        [TestCase("10 B", "1 KB")]
        [TestCase("5 KB", "2 MB")]
        public void Compare_WhenFirstArgumentIsLess_ReturnsMinusOne(object x, object y)
        {
            var comparer = new DocumentSizeComparer();
            Assert.That(comparer.Compare(x, y), Is.EqualTo(-1));
        }

        [TestCase(null, "1 B")]
        [TestCase("1 B", "")]
        [TestCase("1KB", "1 B")]
        [TestCase("1 QB", "1 B")]
        public void Compare_WhenInputFormatIsInvalid_ThrowsArgumentException(object x, object y)
        {
            var comparer = new DocumentSizeComparer();
            Assert.Throws<ArgumentException>(() => comparer.Compare(x, y));
        }

        [TestCase("A KB", "1 KB")]
        public void Compare_WhenValueIsNotNumeric_ThrowsFormatException(object x, object y)
        {
            var comparer = new DocumentSizeComparer();
            Assert.Throws<FormatException>(() => comparer.Compare(x, y));
        }

        [Test]
        public void UseCase_WhenSmallValueGoesFirst_ListIsOrderedAscending()
        {
            var list = new List<string> { "1 B", "1 KB", "5 KB", "5 MB" };
            Assert.That(list, Is.Ordered.Ascending.Using(new DocumentSizeComparer()));
        }

        [Test]
        public void UseCase_WhenBigValueGoesFirst_ListIsOrderedDescending()
        {
            var list = new List<string> { "5 MB", "5 KB", "1 KB", "1 B" };
            Assert.That(list, Is.Ordered.Descending.Using(new DocumentSizeComparer()));
        }
    }
}
