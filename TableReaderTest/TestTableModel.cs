using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TRL = TableReaderLib;

namespace TableReaderTest
{
    /// <summary>
    /// Класс содержит образец данных. Создав такие же данные в источнике, можно сравнить результаты чтения источника с этим классом.
    /// </summary>
    class TestTableModel
    {
        ///Config your sourceTable like below.
        ///|---------------------------------------------------------|
        ///|   col0   |  col1  |   col2    |   col 3    |    col4    |
        ///| (String) |  (int) | (double)  | (DateTime) | (TimeSpan) |
        ///|---------------------------------------------------------|
        ///|"Рус"     |  25    |   34.4    | 03.04.2019 |  00:0:00   |
        ///|"English" |  0     |  -34.4    | 03.04.2019 |  01:24:10  |
        ///|   "1 2"  |  -10   |   0.0     | 03.04.2019 |  02:24:10  |
        ///|  Null    | 9999999| 0.000001  | 03.04.2019 |  03:24:10  |
        ///|  ""        |-9999999|   34.4    | 03.04.2019 |  04:24:10  |
        ///|  Null    |  Null  |   Null    |   Null     |    Null    |
        ///|---------------------------------------------------------|
        TRL.TableReader tr = new TRL.TableReader(null, null);

        public string[] SourceHeaders = new string[]
            {
                "col1",
                "col2",
                "col3",
                "col4",
                "col5",
            };
        public static IReadOnlyList<TRL.TableColumn> Columns =>
            new List<TRL.TableColumn>()
            {
                new TRL.TableColumn(){Name = "col0", Type = typeof(string), IndexInSource = 0},
                new TRL.TableColumn(){Name = "0", Type = typeof(int?), IndexInSource = 1},
                new TRL.TableColumn(){Name = "34,4", Type = typeof(double?), IndexInSource = 2},
                new TRL.TableColumn(){Name = null, Type = typeof(DateTime?), IndexInSource = 3},
                new TRL.TableColumn(){Name = "11:11:11", Type = typeof(TimeSpan?), IndexInSource = 4},
            };
        public static object[][] Data = new object[][]
        {
            new object[]{"Рус", 25, 34.4, new DateTime(2019,04,03), new TimeSpan(0,0,0) },
            new object[]{"English", 0, -34.4, new DateTime(2019,04,03), new TimeSpan(1,24,10) },
            new object[]{"1 2", -10, 0.0, new DateTime(2019,04,03), new TimeSpan(2,24,10) },
            new object[]{null, 9999999, 0.000001, new DateTime(2019,04,03), new TimeSpan(3, 24, 10) },
            new object[]{"", -9999999, 34.4, new DateTime(2019,04,03), new TimeSpan(4, 24, 10) },
            new object[]{ null, null, null, null, null },
        };
       
        
        internal class Row
        {
            public string Col0 { get; set; }
            public int? Col1 { get; set; }
            public double? Col2 { get; set; }
            public DateTime? Col3 { get; set; }
            public TimeSpan? Col4 { get; set; }
        }


    }
}
