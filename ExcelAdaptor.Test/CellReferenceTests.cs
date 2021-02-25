using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace ExcelAdaptor.Test
{
    public class CellReferenceTests
    {
        public CellReferenceTests()
        {

        }
        [Fact]
        public void CtorDefaultTest()
        {
            var t = new CellReference();
            Assert.Equal(CellReference.Null, t);
        }
        [Fact]
        public void Ctor1ArgTest()
        {
            Assert.Throws<ArgumentException>(() => new CellReference(""));
            Assert.Throws<ArgumentException>(() => new CellReference("3"));
            Assert.Throws<ArgumentNullException>(() => new CellReference(null));
            var t = new CellReference("A2");
            Assert.Equal(1, t.ColumnIndex);
            Assert.Equal(2, t.Row);
        }
        [Fact]
        public void Ctor2ArgTest()
        {
            Assert.Throws<ArgumentException>(() => new CellReference(-1, -1));
            Assert.Throws<ArgumentException>(() => new CellReference(0, "A"));
            Assert.Throws<ArgumentException>(() => new CellReference(1, "3"));
            Assert.Throws<ArgumentException>(() => new CellReference(-1, -1));
            var t = new CellReference(2, 1);
            Assert.Equal(1, t.ColumnIndex);
            Assert.Equal(2, t.Row);
            Assert.Equal("A", t.Column);
        }
        [Fact]
        public void IsNullTest()
        {
            Assert.True((new CellReference()).IsNull);
            Assert.False((new CellReference("A1").IsNull));
        }
        [Fact]
        public void IsValidColumnTest()
        {
            Assert.True(CellReference.IsValidColumn(1));
            Assert.False(CellReference.IsValidColumn(0));
            Assert.True(CellReference.IsValidColumn(1000));
            Assert.False(CellReference.IsValidColumn(-1));
        }
        [Fact]
        public void IsValidRowTest()
        {
            Assert.True(CellReference.IsValidRow(1));
            Assert.False(CellReference.IsValidRow(0));
            Assert.True(CellReference.IsValidRow(1000));
            Assert.False(CellReference.IsValidRow(-1));
        }
        [Fact]
        public void GetColumnIndexTest()
        {
            Assert.Equal(1, CellReference.GetColumnIndex("A"));
            Assert.Equal(27, CellReference.GetColumnIndex("aa"));
            Assert.Equal(26 * 26 + 26, CellReference.GetColumnIndex("zz"));
            Assert.Equal(0, CellReference.GetColumnIndex(""));
            Assert.Equal(0, CellReference.GetColumnIndex("AAA"));
        }
        [Fact]
        public void GetColumnNameTest()
        {
            Assert.Equal("", CellReference.GetColumnName(-1));
            Assert.Equal("B", CellReference.GetColumnName(2));
            Assert.Equal("AA", CellReference.GetColumnName(27));
            Assert.Equal("BA", CellReference.GetColumnName(27 + 26));
            Assert.Equal("ZZ", CellReference.GetColumnName(26 * 26 + 26));
            Assert.Equal("", CellReference.GetColumnName(1000));
        }
        [Theory]
        [InlineData("AA1","AA1", false, true, false)]
        [InlineData("A1", "AA1", true, false, false)]
        [InlineData("AA1", "A1", false, false, true)]
        [InlineData("A1", "A2", true, false, false)]
        [InlineData("AA1", "aa1", false, true, false)]
        [InlineData("AA4", "ZZ1", true, false, false)]
        public void ComparisonTest(string s0, string s1, bool lt, bool eq, bool gt)
        {
            var cr0 = new CellReference(s0);
            var cr1 = new CellReference(s1);
            Assert.Equal(eq, cr0 == cr1);
            Assert.NotEqual(eq, cr0 != cr1);
            Assert.Equal(lt, cr0 < cr1);
            Assert.Equal(gt, cr0 > cr1);
            Assert.Equal(lt || eq, cr0 <= cr1);
            Assert.Equal(gt || eq, cr0 >= cr1);
        }
        [Theory]
        [InlineData("A1", "A1")]
        [InlineData("a1", "A1")]
        [InlineData("aA1", "AA1")]
        [InlineData("A10", "A10")]
        [InlineData("z99", "Z99")]
        public void ToStringTest(string s0, string s1)
        {
            var cr0 = new CellReference(s0);
            Assert.Equal(s1, cr0.ToString());
        }
        [Fact]
        public void OtherTest()
        {
            Assert.False((new CellReference("A1").Equals("A1")));
            Assert.Equal("", CellReference.Null.ToString());
        }
    }
}
