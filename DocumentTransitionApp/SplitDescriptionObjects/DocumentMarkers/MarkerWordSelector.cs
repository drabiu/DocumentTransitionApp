using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace SplitDescriptionObjects
{
    public class MarkerWordSelector
    {
        public string Email { get; set; }
        //public int BodyIndex { get; set; }
        public OpenXmlElement Element { get; set; }

        public MarkerWordSelector(OpenXmlElement element)
        {
            Element = element;
        }

        public MarkerWordSelector(OpenXmlElement element, string email)
        {
            Element = element;
            Email = email;
        }

        public static List<MarkerWordSelector> InitializeSelectorsList(Body body)
        {
            var result = new List<MarkerWordSelector>();

            foreach (var child in body.ChildElements)
            {
                result.Add(new MarkerWordSelector(child));
            }

            return result;
        }
    }
}
