using System;
using System.Globalization;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>
    /// A dotted numeric release version, ordered component by component.
    ///
    /// System.Version would almost do the job, but it stores an omitted
    /// component as -1 and therefore orders "0.5" BEFORE "0.5.0". The installed
    /// build reports an assembly version ("0.5.0.0") while a release manifest
    /// carries a product version ("0.5.0"), so that difference would make every
    /// same-version comparison come out as "newer" and re-install forever.
    /// Missing components are zero here, stated explicitly.
    /// </summary>
    public struct ReleaseVersion : IComparable<ReleaseVersion>, IEquatable<ReleaseVersion>
    {
        public const int MaxComponents = 4;
        private const int MaxComponentValue = 65535;

        private readonly int _a, _b, _c, _d;

        private ReleaseVersion(int a, int b, int c, int d)
        {
            _a = a; _b = b; _c = c; _d = d;
        }

        /// <summary>
        /// Parses "MAJOR[.MINOR[.PATCH[.BUILD]]]".
        ///
        /// Deliberately strict: no whitespace, no sign, no pre-release suffix, no
        /// leading "v". A version string is the input to a trust decision, and
        /// lenient parsing there turns a typo into a silent downgrade.
        /// </summary>
        public static bool TryParse(string text, out ReleaseVersion version)
        {
            version = default;
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            string[] parts = text.Split('.');
            if (parts.Length == 0 || parts.Length > MaxComponents)
            {
                return false;
            }

            var values = new int[MaxComponents];
            for (int i = 0; i < parts.Length; i++)
            {
                string part = parts[i];
                if (part.Length == 0 || part.Length > 5)
                {
                    return false;
                }
                foreach (char c in part)
                {
                    if (c < '0' || c > '9')
                    {
                        return false;
                    }
                }
                if (!int.TryParse(part, NumberStyles.None, CultureInfo.InvariantCulture, out int value) ||
                    value > MaxComponentValue)
                {
                    return false;
                }
                values[i] = value;
            }

            version = new ReleaseVersion(values[0], values[1], values[2], values[3]);
            return true;
        }

        public int CompareTo(ReleaseVersion other)
        {
            if (_a != other._a) return _a.CompareTo(other._a);
            if (_b != other._b) return _b.CompareTo(other._b);
            if (_c != other._c) return _c.CompareTo(other._c);
            return _d.CompareTo(other._d);
        }

        public bool Equals(ReleaseVersion other)
        {
            return CompareTo(other) == 0;
        }

        public override bool Equals(object obj)
        {
            return obj is ReleaseVersion other && Equals(other);
        }

        public override int GetHashCode()
        {
            return (((_a * 397) ^ _b) * 397 ^ _c) * 397 ^ _d;
        }

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "{0}.{1}.{2}.{3}", _a, _b, _c, _d);
        }

        public static bool operator >(ReleaseVersion left, ReleaseVersion right) => left.CompareTo(right) > 0;
        public static bool operator <(ReleaseVersion left, ReleaseVersion right) => left.CompareTo(right) < 0;
        public static bool operator >=(ReleaseVersion left, ReleaseVersion right) => left.CompareTo(right) >= 0;
        public static bool operator <=(ReleaseVersion left, ReleaseVersion right) => left.CompareTo(right) <= 0;
        public static bool operator ==(ReleaseVersion left, ReleaseVersion right) => left.CompareTo(right) == 0;
        public static bool operator !=(ReleaseVersion left, ReleaseVersion right) => left.CompareTo(right) != 0;
    }
}
