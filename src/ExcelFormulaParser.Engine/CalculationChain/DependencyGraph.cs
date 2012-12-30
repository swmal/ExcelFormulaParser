using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class DependencyGraph
    {
        private List<string> _items = new List<string>();
        private Dictionary<string, bool> _dict = new Dictionary<string, bool>();

        public virtual bool AddDependency(string from, string to)
        {
            if (_dict.Keys.Contains(from) && !_dict.Keys.Contains(to))
            {
                _items.Insert(_items.IndexOf(from), to);
                _dict[to] = true;
                return true;
            }
            else if (_dict.Keys.Contains(to) && !_dict.Keys.Contains(from))
            {
                _items.Insert(_items.IndexOf(to) + 1, from);
                _dict[from] = true;
                return true;
            }
            else if(_dict.Count == 0)
            {
                _items.Add(to);
                _items.Add(from);
                _dict[to] = true;
                _dict[from] = true;
                return true;
            }
            return false;
        }

        public virtual int Count
        {
            get { return _items.Count; }
        }

        public virtual IEnumerable<string> OrderedFormulas
        {
            get
            {
                return _items;
            }
        }

    }
}
