using DocumentTransitionUniversalApp.TransitionAppServices;
using DocumentTransitionUniversalApp.TransitionAppWCFSerivce;
using System.ServiceModel;
using Windows.Storage;

namespace DocumentTransitionUniversalApp.Repositories
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

        public TransitionAppServiceClient GetInstanceWCF()
        {
            BasicHttpBinding myBinding = new BasicHttpBinding();
            myBinding.MaxBufferPoolSize = 2147483647;
            myBinding.MaxReceivedMessageSize = 2147483647;
            myBinding.MaxBufferSize = 2147483647;
            myBinding.TransferMode = TransferMode.Buffered;
            myBinding.ReaderQuotas.MaxArrayLength = 2147483647;
            myBinding.ReaderQuotas.MaxBytesPerRead = 2147483647;
            myBinding.ReaderQuotas.MaxNameTableCharCount = 2147483647;
            myBinding.ReaderQuotas.MaxDepth = 128;
            myBinding.ReaderQuotas.MaxStringContentLength = 2147483647;

            EndpointAddress myEndpoint = new EndpointAddress(DefaultEndpoint);
            TransitionAppServiceClient serviceClient = new TransitionAppServiceClient(myBinding, myEndpoint);

            return serviceClient;
        }
    }
}
