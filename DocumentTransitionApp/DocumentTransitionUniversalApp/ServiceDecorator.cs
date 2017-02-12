using System.ServiceModel;
using DocumentTransitionUniversalApp.TransitionAppServices;

namespace DocumentTransitionUniversalApp
{
    public class ServiceDecorator
    {
        public string DefaultEndpoint = "http://localhost:6943/TransitionAppServices.asmx";

        public Service1SoapClient GetInstance()
        {
            BasicHttpBinding myBinding = new BasicHttpBinding();
            myBinding.MaxBufferPoolSize = 2147483647;
            myBinding.MaxReceivedMessageSize = 2147483647;

            EndpointAddress myEndpoint = new EndpointAddress(DefaultEndpoint);
            Service1SoapClient serviceClient = new Service1SoapClient(myBinding, myEndpoint);

            return serviceClient;
        }
    }
}
