using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Exceptions;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class FormulaDependency
    {
        public FormulaDependency(ParsingScope scope)
	    {
            ScopeId = scope.ScopeId;
            Address = scope.Address;
	    }
        public Guid ScopeId { get; private set; }

        public RangeAddress Address { get; private set; }

        private Dictionary<Guid, RangeAddress> _referencedBy = new Dictionary<Guid, RangeAddress>();

        private Dictionary<Guid, RangeAddress> _references = new Dictionary<Guid, RangeAddress>();

        public virtual void AddReferenceFrom(Guid scopeId, RangeAddress rangeAddress)
        {
            if (_references.ContainsKey(scopeId) || rangeAddress.CollidesWith(this.Address))
            {
                throw new CircularReferenceException("Circular reference detected");
            }
            _referencedBy.Add(scopeId, rangeAddress);
        }

        public virtual void AddReferenceTo(Guid scopeId, RangeAddress rangeAddress)
        {
            if (_referencedBy.ContainsKey(scopeId) || rangeAddress.CollidesWith(this.Address))
            {
                throw new CircularReferenceException("");
            }
            _references.Add(scopeId, rangeAddress);
        }
    }
}
