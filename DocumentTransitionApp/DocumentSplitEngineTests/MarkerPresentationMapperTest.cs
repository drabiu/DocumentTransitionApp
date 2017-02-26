using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using UnmarshallingSplitXml;

namespace DocumentSplitEngineTests
{
    [TestClass]
    public class MarkerPresentationMapperTest
    {
        [TestInitialize]
        public void Init()
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Split));
            //Split splitXml = (Split)serializer.Deserialize(xmlFile);
            //using (PresentationDocument preDoc =
            //  PresentationDocument.Open(docFile, true))
            //{
            //    PresentationPart body = preDoc.PresentationPart;
            //    IMarkerMapper<SlideId> mapping = new MarkerPresentationMapper(DocumentName, splitXml, body);
            //    DocumentElements = mapping.Run();
            //}
        }

        [TestCleanup]
        public void Finish()
        {
        }
    }
}
