using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentComparator
{
    public class Analysis
    {
        public Dictionary<string, string> deletedClauses = new Dictionary<string, string>();
        public Dictionary<string, string> modifiedClausesDeletedWords = new Dictionary<string, string>();
        public Dictionary<string, string> modifiedClausesNewWords = new Dictionary<string, string>();
        public Dictionary<string, string> newClauses = new Dictionary<string, string>();
        public Dictionary<string, string> modifiedClauses = new Dictionary<string, string>();
    }
}
