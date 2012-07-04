using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Exceptions;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class FormulaDependencies
    {
        public FormulaDependencies()
            : this(new FormulaDependencyFactory())
        {

        }

        public FormulaDependencies(FormulaDependencyFactory formulaDependencyFactory)
        {
            _formulaDependencyFactory = formulaDependencyFactory;
        }

        private readonly FormulaDependencyFactory _formulaDependencyFactory;
        private readonly Dictionary<string, FormulaDependency> _dependencies = new Dictionary<string, FormulaDependency>();

        public IEnumerable<KeyValuePair<string, FormulaDependency>> Dependencies { get { return _dependencies; } }

        public void AddFormulaScope(ParsingScope parsingScope)
        {
            var dependency = _formulaDependencyFactory.Create(parsingScope);
            try
            {
                _dependencies.Add(parsingScope.Address.ToString(), dependency);
            }
            catch (ArgumentException ae)
            {
                //throw new CircularReferenceException("A circular exception was detected, address: " + parsingScope.Address.ToString());
            }
            if (parsingScope.Parent != null)
            {
                if (_dependencies.ContainsKey(parsingScope.Parent.Address.ToString()))
                {
                    var parent = _dependencies[parsingScope.Parent.Address.ToString()];
                    parent.AddReferenceTo(parsingScope.Address);
                    dependency.AddReferenceFrom(parent.Address);
                }


            }
        }
    }
}
