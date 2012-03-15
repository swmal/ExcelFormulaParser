using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.VBA.Functions;

namespace ExcelFormulaParser.Engine.VBA
{
    /// <summary>
    /// This class provides static methods for accessing/modifying VBA Functions.
    /// </summary>
    public static class FunctionRepository
    {
        private static Dictionary<string, VBAFunction> _functions = new Dictionary<string, VBAFunction>();

        /// <summary>
        /// Loads a module of <see cref="VBAFunction"/>s to the function repository.
        /// </summary>
        /// <param name="module">A <see cref="IVBAModule"/> that can be used for adding functions</param>
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

        /// <summary>
        /// Returns true if the the supplied <paramref name="name"/> exists in the repository.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool Exists(string name)
        {
            return _functions.ContainsKey(name.ToLower());
        }

        /// <summary>
        /// Removes all functions from the repository
        /// </summary>
        public static void Clear()
        {
            _functions.Clear();
        }
    }
}
