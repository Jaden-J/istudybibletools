using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;

namespace BibleConfigurator.ModuleConverter
{
    public static class PredefinedBookDifferences
    {
        public static readonly BibleTranslationDifferences KJV = new BibleTranslationDifferences();

        private static BibleTranslationDifferences _rst;
        public static BibleTranslationDifferences RST
        {
            get
            {
                if (_rst == null)
                {
                    _rst = new BibleTranslationDifferences();
                    _rst.BookDifferences.AddRange(new List<BibleBookDifferences>()
                    {
                        new BibleBookDifferences(4,
                                    new BibleBookDifference("12:6", "13:1"),
                                    new BibleBookDifference("13:1-33", "13:X+1")),
                        new BibleBookDifferences(6,
                                    new BibleBookDifference("6:1", "5:16"),
                                    new BibleBookDifference("6:2-27", "6:X-1")),
                        new BibleBookDifferences(9,
                                    new BibleBookDifference("20:42", "20:42-43"),
                                    new BibleBookDifference("23:29", "24:1"),
                                    new BibleBookDifference("24:1-22", "24:X+1")),
                        new BibleBookDifferences(18,
                                    new BibleBookDifference("40:1-5", "39:31-35"),
                                    new BibleBookDifference("40:6-24", "40:X-5"),
                                    new BibleBookDifference("41:1-8", "40:20-27"),
                                    new BibleBookDifference("41:9-34", "41:X-8")),
                        new BibleBookDifferences(19,
                                    new BibleBookDifference("3:1-9:20", "X:X+1"),
                                    new BibleBookDifference("10:1", "9:22"),
                                    new BibleBookDifference("10:2-18", "9:X+21"),
                                    new BibleBookDifference("11:1-7", "10:X"),
                                    new BibleBookDifference("12:1-13:4", "X-1:X+1"),
                                    new BibleBookDifference("13:5-6", "12:6"),
                                    new BibleBookDifference("14:1-17:15", "X-1:X"),
                                    new BibleBookDifference("18:1-22:31", "X-1:X+1"),
                                    new BibleBookDifference("23:1-29:11", "X-1:X"),
                                    new BibleBookDifference("30:1-31:24", "X-1:X+1"),
                                    new BibleBookDifference("32:1-33:22", "X-1:X"),
                                    new BibleBookDifference("34:1-22", "33:X+1"),
                                    new BibleBookDifference("35:1-28", "34:X"),
                                    new BibleBookDifference("36:1-12", "35:X+1"),
                                    new BibleBookDifference("37:1-40", "36:X"),
                                    new BibleBookDifference("38:1-42:31", "X-1:X+1"),
                                    new BibleBookDifference("43:1-5", "42:X"),
                                    new BibleBookDifference("44:1-49:20", "X-1:X+1"),
                                    new BibleBookDifference("50:1-23", "49:X"),
                                    new BibleBookDifference("51:1-52:9", "X-1:X+2"),
                                    new BibleBookDifference("53:1-6", "52:X+1"),
                                    new BibleBookDifference("54:1-7", "53:X+2"),
                                    new BibleBookDifference("55:1-59:17", "X-1:X+1"),
                                    new BibleBookDifference("60:1-12", "59:X+2"),
                                    new BibleBookDifference("61:1-65:13", "X-1:X+1"),
                                    new BibleBookDifference("66:1-20", "65:X"),
                                    new BibleBookDifference("67:1-70:5", "X-1:X+1"),
                                    new BibleBookDifference("71:1-74:23", "X-1:X"),
                                    new BibleBookDifference("75:1-77:20", "X-1:X+1"),
                                    new BibleBookDifference("78:1-79:13", "X-1:X"),
                                    new BibleBookDifference("80:1-81:16", "X-1:X+1"),
                                    new BibleBookDifference("82:1-8", "81:X"),
                                    new BibleBookDifference("83:1-85:13", "X-1:X+1"),
                                    new BibleBookDifference("86:1-17", "85:X"),
                                    new BibleBookDifference("87:1-2", "86:2"),
                                    new BibleBookDifference("87:3-7", "86:X"),
                                    new BibleBookDifference("88:1-90:4", "X-1:X+1"),
                                    new BibleBookDifference("90:5-6", "89:6"),
                                    new BibleBookDifference("90:7-91:16", "X-1:X"),
                                    new BibleBookDifference("92:1-15", "91:X+1"),
                                    new BibleBookDifference("93:1-101:8", "X-1:X"),
                                    new BibleBookDifference("102:1-28", "101:X+1"),
                                    new BibleBookDifference("103:1-107:43", "X-1:X"),
                                    new BibleBookDifference("108:1-13", "107:X+1"),
                                    new BibleBookDifference("109:1-114:8", "X-1:X"),
                                    new BibleBookDifference("115:1-18", "113:X+8"),
                                    new BibleBookDifference("116:1-9", "114:X"),
                                    new BibleBookDifference("116:10-19", "115:X-9"),
                                    new BibleBookDifference("117:1-147:11", "X-1:X"),
                                    new BibleBookDifference("147:12-20", "147:X-11")),
                        new BibleBookDifferences(22,
                                    new BibleBookDifference("1:1", ""),
                                    new BibleBookDifference("1:2-17", "1:X-1"),
                                    new BibleBookDifference("6:13", "7:1"),
                                    new BibleBookDifference("7:1-13", "7:X+1")),
                        new BibleBookDifferences(23,
                                    new BibleBookDifference("3:19-20", "3:19"),
                                    new BibleBookDifference("3:21-26", "3:X-1")),
                        new BibleBookDifferences(27,
                                    new BibleBookDifference("4:1-3", "3:31-33"),
                                    new BibleBookDifference("4:4-37", "4:X-3")),
                        new BibleBookDifferences(28,
                                    new BibleBookDifference("13:16", "14:1"),
                                    new BibleBookDifference("14:1-9", "14:X+1")),
                        new BibleBookDifferences(32,
                                    new BibleBookDifference("1:17", "2:1"),
                                    new BibleBookDifference("2:1-10", "2:X+1")),
                        new BibleBookDifferences(45,
                                    new BibleBookDifference("16:25-27", "14:24-26")),
                        new BibleBookDifferences(47,
                                    new BibleBookDifference("11:32", "11:32-33"),
                                    new BibleBookDifference("13:12-13", "13:12"),
                                    new BibleBookDifference("13:14", "13:13"))
                    });
                }

                return _rst;
            }
        }
    }

    public static class PredefinedBookIndexes
    {
        private static IEnumerable<int> GetRange(int min, int max)
        {
            for (int i = min; i <= max; i++)
                yield return i;
        }

        private static List<int> _kjv;
        public static List<int> KJV
        {
            get
            {
                if (_kjv == null)
                {
                    _kjv = new List<int>(GetRange(1, 66));
                }

                return _kjv;
            }
        }

        private static List<int> _rst;
        public static List<int> RST
        {
            get
            {
                if (_rst == null)
                {
                    _rst = new List<int>(GetRange(1, 44));
                    _rst.AddRange(GetRange(59, 65));
                    _rst.AddRange(GetRange(45, 58));
                    _rst.Add(66);
                }

                return _rst;
            }
        }
    }
}
