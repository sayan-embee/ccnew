using Azure.Storage.Blobs;
using Azure.Storage.Sas;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.CompanyCommunicator.PrivateBlob
{
    public class PrivateBlobAndContainerHelper : IPrivateBlobAndContainerHelper
    {
        public async Task<string> GenerateSASToken(string connectionString, string containerName, string blobName)
        {
            try
            {
                string privateUrl = string.Empty;

                // Create a BlobServiceClient object using the connection string
                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);

                // Get a reference to the container
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);

                // Get a reference to the blob
                BlobClient blobClient = containerClient.GetBlobClient(blobName);

                // Create a SAS token that's valid for specified hour.
                BlobSasBuilder sasBuilder = new BlobSasBuilder()
                {
                    BlobContainerName = containerName,
                    BlobName = blobName,
                    Resource = "b"
                };

                // 100 years from 2023
                sasBuilder.ExpiresOn = DateTimeOffset.UtcNow.AddHours(876000);
                sasBuilder.SetPermissions(BlobSasPermissions.Read);

                privateUrl = blobClient.GenerateSasUri(sasBuilder).AbsoluteUri;

                await Task.Delay(0);

                return privateUrl;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
