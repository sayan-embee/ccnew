using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.CompanyCommunicator.PrivateBlob
{
    public interface IPrivateBlobAndContainerHelper
    {
        Task<string> GenerateSASToken(string connectionString, string containerName, string blobName);
    }
}