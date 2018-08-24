namespace DataRecovery.Contracts
{
    public interface IFileValidator
    {
        bool CheckFileType(string fileName);
    }
}
