using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DocumentTransitionAppServices
{
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

        public bool IsError { get; set; }
        public string Message { get; set; }
        public object Data { get; set; }
    }
}