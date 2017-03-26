using System.Runtime.Serialization;

namespace SplitDescriptionObjects
{
    [DataContract]
    public class PersonFiles
    {
        [DataMember]
        public string Person { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public byte[] Data { get; set; }
    }
}
