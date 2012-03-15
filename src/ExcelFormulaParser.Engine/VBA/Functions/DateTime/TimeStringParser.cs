using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelFormulaParser.Engine.VBA.Functions.DateTime
{
    public class TimeStringParser
    {
        private double GetSerialNumber(int hour, int minute, int second)
        {
            var secondsInADay = 24d * 60d * 60d;
            return ((double)hour * 60 * 60 + (double)minute * 60 + (double)second) / secondsInADay;
        }

        private void ValidateValues(int hour, int minute, int second)
        {
            if (second < 0 || second > 59)
            {
                throw new FormatException("Illegal value for second: " + second);
            }
            if (minute < 0 || minute > 59)
            {
                throw new FormatException("Illegal value for minute: " + minute);
            }
        }

        public virtual double Parse(string input)
        {
            return InternalParse(input);
        }

        public virtual bool CanParse(string input)
        {
            return InternalParse(input) != -1;
        }

        private double InternalParse(string input)
        {
            if (Regex.IsMatch(input, @"^[0-9]{1,2}(:[0-9]){0,2}"))
            {
                int hour = 0, minute = 0, second = 0;
                var items = input.Split(':');
                hour = int.Parse(items[0]);
                if (items.Length > 1)
                {
                    minute = int.Parse(items[1]);
                }
                if (items.Length > 2)
                {
                    second = int.Parse(items[2]);
                }
                ValidateValues(hour, minute, second);
                return GetSerialNumber(hour, minute, second);
            }
            return -1;
        }
    }
}
