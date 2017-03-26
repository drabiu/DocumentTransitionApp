using System.Runtime.Serialization;

namespace DocumentTransitionAppService
{
    [DataContract]
    public class ServiceResponse
    {
        public ServiceResponse()
        {

        }

        public ServiceResponse(string message)
        {
            IsError = true;
            Message = message;
        }

        public ServiceResponse(object data)
        {
            IsError = false;
            Message = string.Empty;
            Data = data;
        }

        [DataMember]
        public bool IsError { get; set; }
        [DataMember]
        public string Message { get; set; }
        [DataMember]
        public object Data { get; set; }
    }
}