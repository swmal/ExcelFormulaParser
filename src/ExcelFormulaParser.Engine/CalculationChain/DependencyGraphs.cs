using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class DependencyGraphs
    {
        private readonly List<DependencyGraph> _graphs = new List<DependencyGraph>();
        private readonly Dictionary<string, DependencyGraph> _graphsDictionary = new Dictionary<string, DependencyGraph>();

        public virtual void Add(string from, string to)
        {
            var graph = GetGraph(from, to);
            if (graph != null)
            {
                graph.AddDependency(from, to);
                SetGraph(from, graph);
                SetGraph(to, graph);
                return;
            }
            var newGraph = new DependencyGraph();
            newGraph.AddDependency(from, to);
            _graphsDictionary.Add(from, newGraph);
            _graphsDictionary.Add(to, newGraph);
            _graphs.Add(newGraph);
        }

        private void SetGraph(string address, DependencyGraph graph)
        {
            if (!_graphsDictionary.ContainsKey(address))
            {
                _graphsDictionary.Add(address, graph);
            }
        }

        private DependencyGraph GetGraph(string from, string to)
        {
            if (_graphsDictionary.ContainsKey(from))
            {
                return _graphsDictionary[from];
            }
            if (_graphsDictionary.ContainsKey(to))
            {
                return _graphsDictionary[to];
            }
            return null;
        }

        public virtual IEnumerable<DependencyGraph> GetGraphs()
        {
            return _graphs.OrderByDescending(x => x.Count);
        }
    }
}
