using DataRecovery.Contracts;
using System.IO;

namespace DataRecovery.Helpers
{
    public class FileValidator : IFileValidator
    {
        public bool CheckFileType(string fileName)
        {
            string ext = Path.GetExtension(fileName);
            switch (ext.ToLower())
            {
                case ".xlsx":
                    return true;
                case ".xlx":
                    return false;
                default:
                    return false;
            }
        }
    }
}
