using System.ServiceModel;
using DocumentTransitionUniversalApp.TransitionAppServices;
using Windows.Storage;

namespace DocumentTransitionUniversalApp
{
    public class ServiceDecorator
    {
        public string DefaultEndpoint
        {
            get
            {
                if (ApplicationData.Current.LocalSettings.Values.ContainsKey("DefaultEndpoint"))
                    return (string)(ApplicationData.Current.LocalSettings.Values["DefaultEndpoint"]);
                else
                    return "http://localhost:6943/TransitionAppServices.asmx";

            }
            set
            {
                ApplicationData.Current.LocalSettings.Values["DefaultEndpoint"] = value;
            }
        }

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
