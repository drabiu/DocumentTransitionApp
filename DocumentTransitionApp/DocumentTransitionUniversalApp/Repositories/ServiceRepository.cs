using DocumentTransitionUniversalApp.TransitionAppWCFSerivce;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace DocumentTransitionUniversalApp.Repositories
{
    public class ServiceRepository
    {
        TransitionAppServiceClient _serviceClient;

        public ServiceRepository()
        {
            ServiceDecorator service = new ServiceDecorator();
            _serviceClient = service.GetInstanceWCF();
        }

        public async Task<byte[]> MergeWordAsync(ObservableCollection<PersonFiles> files)
        {
            return await _serviceClient.MergeWordAsync(files);
        }

        public async Task<byte[]> MergeExcelAsync(ObservableCollection<PersonFiles> files)
        {
            return await _serviceClient.MergeExcelAsync(files);
        }

        public async Task<byte[]> MergePresentationAsync(ObservableCollection<PersonFiles> files)
        {
            return await _serviceClient.MergePresentationAsync(files);
        }

        public async Task<byte[]> GenerateSplitWordAsync(string docName, ObservableCollection<TransitionAppWCFSerivce.PartsSelectionTreeElement> parts)
        {
            return await _serviceClient.GenerateSplitWordAsync(docName, parts);
        }

        public async Task<byte[]> GenerateSplitExcelAsync(string docName, ObservableCollection<TransitionAppWCFSerivce.PartsSelectionTreeElement> parts)
        {
            return await _serviceClient.GenerateSplitExcelAsync(docName, parts);
        }

        public async Task<byte[]> GenerateSplitPresentationAsync(string docName, ObservableCollection<TransitionAppWCFSerivce.PartsSelectionTreeElement> parts)
        {
            return await _serviceClient.GenerateSplitPresentationAsync(docName, parts);
        }

        public async Task<GetPartsFromXmlServiceResponse> GetWordPartsFromXmlAsync(string docName, byte[] documentFile, byte[] splitFile)
        {
            return await _serviceClient.GetWordPartsFromXmlAsync(docName, documentFile, splitFile);
        }

        public async Task<GetPartsFromXmlServiceResponse> GetExcelPartsFromXmlAsync(string docName, byte[] documentFile, byte[] splitFile)
        {
            return await _serviceClient.GetExcelPartsFromXmlAsync(docName, documentFile, splitFile);
        }

        public async Task<GetPartsFromXmlServiceResponse> GetPresentationPartsFromXmlAsync(string docName, byte[] documentFile, byte[] splitFile)
        {
            return await _serviceClient.GetPresentationPartsFromXmlAsync(docName, documentFile, splitFile);
        }

        public async Task<ObservableCollection<PersonFiles>> SplitWordAsync(string docName, byte[] docFile, byte[] xmlFile)
        {
            return await _serviceClient.SplitWordAsync(docName, docFile, xmlFile);
        }

        public async Task<ObservableCollection<PersonFiles>> SplitExcelAsync(string docName, byte[] docFile, byte[] xmlFile)
        {
            return await _serviceClient.SplitExcelAsync(docName, docFile, xmlFile);
        }

        public async Task<ObservableCollection<PersonFiles>> SplitPresentationAsync(string docName, byte[] docFile, byte[] xmlFile)
        {
            return await _serviceClient.SplitPresentationAsync(docName, docFile, xmlFile);
        }

        public async Task<ObservableCollection<TransitionAppWCFSerivce.PartsSelectionTreeElement>> GetWordPartsAsync(string docName, byte[] documentFile)
        {
            return await _serviceClient.GetWordPartsAsync(docName, documentFile);
        }

        public async Task<ObservableCollection<TransitionAppWCFSerivce.PartsSelectionTreeElement>> GetPresentationPartsAsync(string docName, byte[] documentFile)
        {
            return await _serviceClient.GetPresentationPartsAsync(docName, documentFile);
        }

        public async Task<ObservableCollection<TransitionAppWCFSerivce.PartsSelectionTreeElement>> GetExcelPartsAsync(string docName, byte[] documentFile)
        {
            return await _serviceClient.GetExcelPartsAsync(docName, documentFile);
        }
    }
}
