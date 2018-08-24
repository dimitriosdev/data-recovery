using System.Windows.Forms;

namespace DataRecoveryLib
{
    public class FileValidator
    {
        private string fileName;

        public string FileName { get { return fileName; } }

        public bool IsValidExcelFile()
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.InitialDirectory = @"C:\Users\dimitrios.metozis\Downloads";
            fd.ShowDialog();
            fileName = fd.FileName;
            return fd.FileName.EndsWith("xlsx");
        }
    }
}
