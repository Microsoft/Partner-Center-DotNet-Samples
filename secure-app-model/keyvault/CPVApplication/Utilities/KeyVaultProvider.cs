// -----------------------------------------------------------------------
// <copyright file="KeyVaultProvider.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// -----------------------------------------------------------------------

namespace CPVApplication.Utilities
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
        /// <summary>
        /// The base address for the instance of Azure Key Vault.
        /// </summary>
        private readonly string BaseUrl = ConfigurationManager.AppSettings["KeyVaultEndpoint"];

        /// <summary>
        /// The identifier for the Azure AD application configured to access the instance of Azure Key Vault.
        /// </summary>
        private readonly string KeyVaultTenantId = ConfigurationManager.AppSettings["ida:KeyVaultTenantId"];

        /// <summary>
        /// The identifier for the Azure AD application configured to access the instance of Azure Key Vault.
        /// </summary>
        private readonly string KeyVaultClientId = ConfigurationManager.AppSettings["ida:KeyVaultClientId"];

        /// <summary>
        /// The application secret for the Azure AD application configured to access the instance of Azure Key Vault.
        /// </summary>
        private readonly string KeyVaultClientSecret = ConfigurationManager.AppSettings["ida:KeyVaultClientSecret"];

        /// <summary>
        /// The client used to interact with the instance of Azure Key Vault.
        /// </summary>
        private readonly SecretClient secretClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="KeyVaultProvider" /> class.
        /// </summary>
        public KeyVaultProvider()
        {
            secretClient = new SecretClient(new Uri(BaseUrl),
                new ClientSecretCredential(
                    KeyVaultTenantId,
                    KeyVaultClientId,
                    KeyVaultClientSecret));
        }

        /// <summary>
        /// Gets the specified entity from the vault. 
        /// </summary>
        /// <param name="key">Identifier of the entity to be retrieved.</param>
        /// <returns>The value retrieved from the vault.</returns>
        public async Task<string> GetSecretAsync(string key)
        {
            KeyVaultSecret secret = await secretClient.GetSecretAsync(key.Replace("@", string.Empty).Replace(".", string.Empty)).ConfigureAwait(false);
            return secret.Value;
        }
    }
}
