using System;
using System.Collections.Generic;

namespace SplitDescriptionObjects
{
    public static class MarkerHelper<ElementType>
    {
        public static IList<int> GetCrossedElements(string id, string id2, IList<ElementType> elements, Func<ElementType, string> getElementId)
        {
            return GetCrossedElements(id, id2, elements, el => el is ElementType, getElementId);
        }

        public static IList<int> GetCrossedElements(string id, string id2, IList<ElementType> elements, Predicate<ElementType> elementsFilter, Func<ElementType, string> getElementId)
        {
            bool startSelection = false;
            IList<int> indexes = new List<int>();
            for (int index = 0; index < elements.Count; index++)
            {
                var element = elements[index];
                if (elementsFilter(element))
                {
                    if (getElementId(element) == id)
                        startSelection = true;

                    if (startSelection)
                        indexes.Add(index);

                    if (getElementId(element) == id2)
                        break;
                }
            }

            return indexes;
        }
    } 
}
