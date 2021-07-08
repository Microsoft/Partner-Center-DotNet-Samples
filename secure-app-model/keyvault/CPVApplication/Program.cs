// -----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// -----------------------------------------------------------------------

namespace CPVApplication
{
    using System;
    using System.Configuration;
    using System.Threading.Tasks;
    using CPVApplication.Utilities;
    using Microsoft.Store.PartnerCenter;
    using Microsoft.Store.PartnerCenter.Extensions;
    using Microsoft.Store.PartnerCenter.Models;
    using Microsoft.Store.PartnerCenter.Models.Users;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    internal class Program
    {
        /**
         * The following code assumes that the context of the partner is pre-determined by some external process.
         */

        private static readonly string AADInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private static readonly string CPVApplicationId = ConfigurationManager.AppSettings["ida:CPVApplicationId"];
        private static readonly string CPVApplicationSecret = ConfigurationManager.AppSettings["ida:CPVApplicationSecret"];

        private static void Main(string[] args)
        {
            Task.Run(() => Run()).Wait();
        }

        private static async Task Run()
        {
            // The following properties indicate which partner and customer context the calls are going to be made.
            string PartnerId = "<Partner tenant id>";
            string CustomerId = "<Customer tenant id>";

            Console.WriteLine(" ===================== Partner center  API calls ============================", DateTime.Now);
            IAggregatePartner ops = await GetUserPartnerOperationsAsync(PartnerId);
            SeekBasedResourceCollection<CustomerUser> customerUsers = ops.Customers.ById(CustomerId).Users.Get();
            Console.WriteLine(JsonConvert.SerializeObject(customerUsers));

            Console.WriteLine(" ===================== Partner graph  API calls ============================", DateTime.Now);
            Tuple<string, DateTimeOffset> tokenResult = await LoginToGraph(PartnerId);
            JObject mydetails = await ApiCalls.GetAsync(tokenResult.Item1, "https://graph.microsoft.com/v1.0/me");
            Console.WriteLine(JsonConvert.SerializeObject(mydetails));

            // The customer graph calls require customer admin to consent the application
            Console.WriteLine(" ===================== Customer consent for graph  API calls ============================", DateTime.Now);

            // Enable consent
            Tuple<string, DateTimeOffset> tokenPartnerResult = await LoginToPartnerCenter(PartnerId);
            JObject contents = new JObject
            {
                // Provide your application display name
                ["displayName"] = "CPV Marketplace",

                // Provide your application id
                ["applicationId"] = CPVApplicationId,

                // Provide your application grants
                ["applicationGrants"] = new JArray(
                    JObject.Parse("{\"enterpriseApplicationId\": \"00000002-0000-0000-c000-000000000000\", \"scope\":\"Domain.ReadWrite.All,User.ReadWrite.All,Directory.Read.All\"}"), // for graph api access,  Directory.Read.All
                    JObject.Parse("{\"enterpriseApplicationId\": \"797f4846-ba00-4fd7-ba43-dac1f8f63013\", \"scope\":\"user_impersonation\"}")) // for ARM api access
            };

            /** The following steps have to performed once in per customer tenant if your application is Control panel vendor application and requires customer tenant graph access  **/

            // delete the previous grant into customer tenant
            JObject consentDeletion = await ApiCalls.DeleteAsync(
                tokenPartnerResult.Item1,
               string.Format("https://api.partnercenter.microsoft.com/v1/customers/{0}/applicationconsents/{1}", CustomerId, CPVApplicationId));
            Console.WriteLine(JsonConvert.SerializeObject(consentDeletion));

            // create new grants for the application given the setting in application grants payload.
            JObject consentCreation = await ApiCalls.PostAsync(
                tokenPartnerResult.Item1,
                string.Format("https://api.partnercenter.microsoft.com/v1/customers/{0}/applicationconsents", CustomerId),
                contents.ToString());
            Console.WriteLine(JsonConvert.SerializeObject(consentCreation));


            Console.WriteLine(" ===================== Customer graph  API calls ============================", DateTime.Now);

            Tuple<string, DateTimeOffset> tokenCustomerResult = await LoginToCustomerGraph(PartnerId, CustomerId);
            JObject customerDomainsUsingGraph = await ApiCalls.GetAsync(tokenCustomerResult.Item1, "https://graph.windows.net/" + CustomerId + "/domains?api-version=1.6");
            Console.WriteLine(JsonConvert.SerializeObject(customerDomainsUsingGraph));

            Console.ReadLine();
        }

        public static async Task<Tuple<string, DateTimeOffset>> LoginToPartnerCenter(string tenantId)
        {
            KeyVaultProvider provider = new KeyVaultProvider();
            string refreshToken = await provider.GetSecretAsync(tenantId);

            JObject token = await AuthorizationUtilities.GetAADTokenFromRefreshToken(
                $"{AADInstance}/{tenantId}",
                "https://api.partnercenter.microsoft.com",
                CPVApplicationId,
                CPVApplicationSecret,
                refreshToken);

            return new Tuple<string, DateTimeOffset>(token["access_token"].ToString(), DateTimeOffset.UtcNow + TimeSpan.FromTicks(long.Parse(token["expires_on"].ToString())));
        }

        /// <summary>
        /// Generates a token for Partner tenant using partner user refresh token
        /// The expiry time calculation explined below is not strong. 
        /// please use stardard ADAL library to gain access token by refresh token, which provides strongly typed classes with proper expirty time calculation.
        /// </summary>
        /// <param name="partnerTenantId">partner tenant id</param>
        /// <returns>
        /// Access token and expiry time. 
        /// </returns>
        public static async Task<Tuple<string, DateTimeOffset>> LoginToGraph(string partnerTenantId)
        {
            KeyVaultProvider provider = new KeyVaultProvider();
            string refreshToken = await provider.GetSecretAsync(partnerTenantId);

            JObject token = await AuthorizationUtilities.GetAADTokenFromRefreshToken(
                $"{AADInstance}/{partnerTenantId}",
                "https://graph.microsoft.com",
                CPVApplicationId,
                CPVApplicationSecret,
                refreshToken);

            return new Tuple<string, DateTimeOffset>(token["access_token"].ToString(), DateTimeOffset.UtcNow + TimeSpan.FromTicks(long.Parse(token["expires_on"].ToString())));
        }

        /// <summary>
        /// Generates a token for Customer tenant using partner user refresh token
        /// The expiry time calculation explined below is not strong. 
        /// please use stardard ADAL library to gain access token by refresh token, which provides strongly typed classes with proper expirty time calculation.
        /// </summary>
        /// <param name="partnerTenantId">partner tenant id</param>
        /// <param name="customerTenantId">customer tenant id</param>
        /// <returns>
        /// Access token and expiry time. 
        /// </returns>
        public static async Task<Tuple<string, DateTimeOffset>> LoginToCustomerGraph(string partnerTenantId, string customerTenantId)
        {
            KeyVaultProvider provider = new KeyVaultProvider();
            string refreshToken = await provider.GetSecretAsync(partnerTenantId);

            JObject token = await AuthorizationUtilities.GetAADTokenFromRefreshToken(
                $"{AADInstance}/{customerTenantId}",
                "https://graph.windows.net",
                CPVApplicationId,
                CPVApplicationSecret,
                refreshToken);

            return new Tuple<string, DateTimeOffset>(token["access_token"].ToString(), DateTimeOffset.UtcNow + TimeSpan.FromTicks(long.Parse(token["expires_on"].ToString())));
        }

        /// <summary>
        /// Using Partner center .NET SDK to make partner center calls. If you are using java application, you can either us REST calls or use Java SDK to make partner center calls
        /// In all the mentioned cases, the token generated by LoginToPartnerCenter method should work.
        /// </summary>
        /// <param name="PartnerId">partner tenant id</param>
        /// <returns>SDK reference for contextual calls</returns>
        public static async Task<IAggregatePartner> GetUserPartnerOperationsAsync(string PartnerId)
        {
            Tuple<string, DateTimeOffset> aadAuthenticationResult = await LoginToPartnerCenter(PartnerId);

            // Authenticate by user context with the partner service
            IPartnerCredentials userCredentials = PartnerCredentials.Instance.GenerateByUserCredentials(
                CPVApplicationId,
                new AuthenticationToken(
                    aadAuthenticationResult.Item1,
                    aadAuthenticationResult.Item2),
                async delegate
                {
                    // token has expired, re-Login to Azure Active Directory
                    Tuple<string, DateTimeOffset> aadToken = await LoginToPartnerCenter(PartnerId);
                    return new AuthenticationToken(aadToken.Item1, aadToken.Item2);
                });

            return PartnerService.Instance.CreatePartnerOperations(userCredentials);
        }
    }
}