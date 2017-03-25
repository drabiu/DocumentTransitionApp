using System;
using System.Collections.Generic;

namespace OpenXMLTools.Helpers
{
    public class NameIndexer
    {
        private Dictionary<string, int>[] Indexes;

        public NameIndexer(IList<string> nameList)
        {
            Indexes = new Dictionary<string, int>[Enum.GetNames(typeof(ElementType)).Length + 1];
            for (int i = 0; i < Indexes.Length; i++)
            {
                Indexes[i] = new Dictionary<string, int>();
                foreach (var name in nameList)
                    Indexes[i].Add(name, 0);
            }

        }

        public int GetNextIndex(string name)
        {
            return Indexes[0][name]++;
        }

        public int GetNextIndex(string name, ElementType type)
        {
            return Indexes[(int)type + 1][name]++;
        }
    }
}
