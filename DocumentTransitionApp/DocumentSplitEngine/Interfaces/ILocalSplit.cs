using System;

namespace DocumentSplitEngine.Interfaces
{
    [Obsolete]
    public interface ILocalSplit
    {
        void OpenAndSearchDocument(string filePath, string xmlSplitDefinitionFilePath);
        void SaveSplitDocument(string filePath);
    }
}
