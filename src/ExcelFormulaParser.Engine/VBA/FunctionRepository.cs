using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Functions;

namespace ExcelFormulaParser.Engine.VBA
{
    public static class FunctionRepository
    {
        private static Dictionary<string, VBAFunction> _functions = new Dictionary<string, VBAFunction>();

        public static void LoadModule(IVBAModule module)
        {
            foreach (var key in module.Functions.Keys)
            {
                var lowerKey = key.ToLower();
                _functions[lowerKey] = module.Functions[key];
            }
        }

        public static VBAFunction GetFunction(string name)
        {
            if(!_functions.ContainsKey(name.ToLower()))
            {
                throw new InvalidOperationException("Non supported function: " + name);
            }
            return _functions[name.ToLower()];
        }

        public static bool Exists(string name)
        {
            return _functions.ContainsKey(name.ToLower());
        }

        public static void Clear()
        {
            _functions.Clear();
        }
    }
}
