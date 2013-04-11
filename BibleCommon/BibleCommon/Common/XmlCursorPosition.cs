using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace BibleCommon.Common
{
    public struct XmlCursorPosition : IComparable
    {
        public int LineNumber;
        public int LinePosition;

        public int CompareTo(object obj)
        {
            var otherObj = (XmlCursorPosition)obj;

            var result = this.LineNumber.CompareTo(otherObj.LineNumber);
            if (result == 0)
                result = this.LinePosition.CompareTo(otherObj.LinePosition);

            return result;
        }

        public XmlCursorPosition(IXmlLineInfo lineInfo)
        {
            LineNumber = lineInfo.LineNumber;
            LinePosition = lineInfo.LinePosition;
        }

        public XmlCursorPosition(string lineInfo)
        {
            var parts = lineInfo.Split(new char[] { ';' });

            LineNumber = int.Parse(parts[0]);
            LinePosition = int.Parse(parts[1]);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is XmlCursorPosition))
                return false;

            var otherObj = (XmlCursorPosition)obj;

            return this.LineNumber == otherObj.LineNumber
                && this.LinePosition == otherObj.LinePosition;
        }

        public override string ToString()
        {
            return string.Format("{0};{1}", this.LineNumber, this.LinePosition);
        }

        public override int GetHashCode()
        {
            return this.LineNumber.GetHashCode() ^ this.LinePosition.GetHashCode();
        }

        public static bool operator >(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return cp1.CompareTo(cp2) > 0;
        }

        public static bool operator <(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return cp1.CompareTo(cp2) < 0;
        }

        public static bool operator ==(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return cp1.Equals(cp2);
        }

        public static bool operator !=(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return !cp1.Equals(cp2);
        }

        public static bool operator >=(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return cp1.CompareTo(cp2) >= 0;
        }

        public static bool operator <=(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return cp1.CompareTo(cp2) <= 0;
        }
    }
}
