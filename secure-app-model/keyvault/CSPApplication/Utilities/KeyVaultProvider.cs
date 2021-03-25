// -----------------------------------------------------------------------
// <copyright file="KeyVaultProvider.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// -----------------------------------------------------------------------

namespace CSPApplication.Utilities
{
    using Azure.Identity;
    using Azure.Security.KeyVault.Secrets;
    using System;
    using System.Configuration;
    using System.Threading.Tasks;

    /// <summary>
    /// Provider for accessing secrets from the Azure KeyVault
    /// </summary>
    public class KeyVaultProvider
    {
        private readonly string KeyVaultTenantId = ConfigurationManager.AppSettings["ida:KeyVaultTenantId"];
        private readonly string KeyVaultClientId = ConfigurationManager.AppSettings["ida:KeyVaultClientId"];
        private readonly string KeyVaultClientSecret = ConfigurationManager.AppSettings["ida:KeyVaultClientSecret"];
        private readonly string BaseUrl = ConfigurationManager.AppSettings["KeyVaultEndpoint"];

        public async Task<string> GetSecretAsync(string key)
        {
            SecretClient secretClient = new SecretClient(new Uri(BaseUrl),
               new ClientSecretCredential(
                   KeyVaultTenantId,
                   KeyVaultClientId,
                   KeyVaultClientSecret));

            KeyVaultSecret secret = await secretClient.GetSecretAsync(key.Replace("@", string.Empty).Replace(".", string.Empty)).ConfigureAwait(false);
            return secret.Value;
        }
    }
}
