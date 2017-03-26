using DocumentEditPartsEngine;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace DocumentTransitionAppWCF.Responses
{
    [DataContract]
    public class GetPartsFromXmlServiceResponse
    {
        [DataMember]
        public bool IsError { get; set; }
        [DataMember]
        public string Message { get; set; }
        [DataMember]
        public List<PartsSelectionTreeElement> Data { get; set; }

        public GetPartsFromXmlServiceResponse()
        {

        }

        public GetPartsFromXmlServiceResponse(string message)
        {
            IsError = true;
            Message = message;
        }

        public GetPartsFromXmlServiceResponse(List<PartsSelectionTreeElement> data)
        {
            IsError = false;
            Message = string.Empty;
            Data = data;
        }
    }
}