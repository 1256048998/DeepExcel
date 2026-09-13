using DeepExcel.AddIn.Updates;
using Xunit;

namespace DeepExcel.Tests
{
    /// <summary>
    /// Version ordering.
    ///
    /// "Is this newer?" is the question that decides whether a machine installs
    /// something, and it is also the downgrade defence. Both of the interesting
    /// cases are boring-looking: the installed build reports four components
    /// ("0.5.0.0") while a manifest carries three ("0.5.0"), and System.Version
    /// would call those different.
    /// </summary>
    public class ReleaseVersionTests
    {
        private static ReleaseVersion V(string text)
        {
            Assert.True(ReleaseVersion.TryParse(text, out ReleaseVersion version), text);
            return version;
        }

        [Theory]
        [InlineData("0.5.0", "0.5.0.0")]
        [InlineData("1", "1.0.0.0")]
        [InlineData("0.5", "0.5.0")]
        public void Missing_components_are_zero_not_unspecified(string shorter, string longer)
        {
            Assert.Equal(0, V(shorter).CompareTo(V(longer)));
        }

        [Theory]
        [InlineData("0.5.1", "0.5.0")]
        [InlineData("0.6.0", "0.5.99")]
        [InlineData("1.0.0", "0.99.99")]
        [InlineData("0.5.0.1", "0.5.0")]
        [InlineData("0.10.0", "0.9.0")]   // numeric, not lexicographic
        public void Orders_numerically(string greater, string lesser)
        {
            Assert.True(V(greater) > V(lesser));
            Assert.True(V(lesser) < V(greater));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("v1.0")]
        [InlineData("1.0.0-beta")]
        [InlineData("1.0.0.0.0")]
        [InlineData("1..0")]
        [InlineData("-1.0")]
        [InlineData(" 1.0")]
        [InlineData("1.0 ")]
        [InlineData("1.0x")]
        [InlineData("65536.0")]
        [InlineData("999999")]
        public void Refuses_anything_that_is_not_dotted_numeric(string text)
        {
            Assert.False(ReleaseVersion.TryParse(text, out _));
        }

        [Fact]
        public void Equality_matches_comparison()
        {
            Assert.True(V("0.5.0") == V("0.5.0.0"));
            Assert.False(V("0.5.0") != V("0.5.0.0"));
            Assert.Equal(V("0.5.0").GetHashCode(), V("0.5.0.0").GetHashCode());
        }
    }
}
