using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;

namespace DocumentTransitionUniversalApp.Helpers
{
    public static class FileHelper
    {
        public static async void SaveFile(string fileName, byte[] fileBinary, params string[] filters)
        {
            FolderPicker folderPicker = new FolderPicker();
            foreach (var filter in filters)
                folderPicker.FileTypeFilter.Add(filter);

            folderPicker.SuggestedStartLocation = PickerLocationId.Downloads;
            StorageFolder folder = await folderPicker.PickSingleFolderAsync();
            StorageFile newFile;
            if (folder != null)
            {
                try
                {
                    newFile = await folder.GetFileAsync(fileName);
                }
                catch (FileNotFoundException ex)
                {
                    newFile = await folder.CreateFileAsync(fileName);
                }

                using (var s = await newFile.OpenStreamForWriteAsync())
                {
                    s.Write(fileBinary, 0, fileBinary.Length);
                }
            }
        }

        public static async void SaveFile(byte[] fileBinary, string fileName, params string[] filters)
        {
            FileSavePicker fileSavePicker = new FileSavePicker();
            foreach (var filter in filters)
                fileSavePicker.FileTypeChoices.Add(string.Format("Document ({0})", filter), filters);

            fileSavePicker.SuggestedStartLocation = PickerLocationId.Downloads;
            fileSavePicker.SuggestedFileName = fileName;
            StorageFile file = await fileSavePicker.PickSaveFileAsync();
            if (file != null)
            {
                using (var s = await file.OpenStreamForWriteAsync())
                {
                    s.Write(fileBinary, 0, fileBinary.Length);
                }
            }
        }
    }
}
