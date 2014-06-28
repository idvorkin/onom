using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OnenoteCapabilities;

namespace OneNoteObjectModelTests
{
    [TestFixture]
    public class HumanizedDateParserTests
    {
        [Test]
        public void TestAll()
        {
            var possibleDates = new List<Tuple<string, bool,string>>()
            {
                Tuple.Create("#do something next month", true,"#do something"),
                Tuple.Create("#do something 7/16/2014", true, "#do something"),
                Tuple.Create("next month", true, ""),
                Tuple.Create("7/16/2014", true, ""),
                Tuple.Create("play on 7/16/2014", true, "play on"),
                Tuple.Create("play on the beach", false, ""),
                Tuple.Create("do it tomorrow", true, "do it"),
                Tuple.Create("do it next week", true, "do it"),
                Tuple.Create("tomorrow", true,""),
            };
            foreach (var possibleDate in possibleDates)
            {
                var parsed = HumanizedDateParser.ParseDateAtEndOfSentance(possibleDate.Item1);
                Assert.That(parsed.Parsed, Is.EqualTo(possibleDate.Item2));
                if (parsed.Parsed)
                {
                    Assert.That(parsed.SentanceWithoutDate, Is.EqualTo(possibleDate.Item3));
                }
            }
        }
    }

}
