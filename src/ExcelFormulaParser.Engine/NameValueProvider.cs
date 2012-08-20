using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine
{
    public class NameValueProvider
    {
        private ExcelDataProvider _excelDataProvider;
        private IDictionary<string, object> _values;

        public NameValueProvider(ExcelDataProvider excelDataProvider)
        {
            _excelDataProvider = excelDataProvider;
            _values = _excelDataProvider.GetWorkbookNameValues();
        }

        public virtual bool IsNamedValue(string key)
        {
            return _values.ContainsKey(key);
        }

        public virtual object GetNamedValue(string key)
        {
            return _values[key];
        }

        public virtual void Reload()
        {
            _values = _excelDataProvider.GetWorkbookNameValues();
        }
    }
}
