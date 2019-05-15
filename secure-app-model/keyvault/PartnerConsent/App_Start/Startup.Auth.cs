﻿// -----------------------------------------------------------------------
// <copyright file="Startup.Auth.cs" company="Microsoft">
//      Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// -----------------------------------------------------------------------

namespace PartnerConsent
{
    using System.Configuration;
    using System.IdentityModel.Claims;
    using System.Threading.Tasks;
    using Microsoft.Owin.Security;
    using Microsoft.Owin.Security.Cookies;
    using Microsoft.Owin.Security.Notifications;
    using Microsoft.Owin.Security.OpenIdConnect;
    using Owin;
    using Utilities;

    public partial class Startup
    {
        private static string CSPApplicationId = ConfigurationManager.AppSettings["ida:CSPApplicationId"];
        private static string CSPApplicationSecret = ConfigurationManager.AppSettings["ida:CSPApplicationSecret"];
        private static string AuthorityCommon = ConfigurationManager.AppSettings["ida:AADInstance"] + "common";

        /// <summary>
        /// Configure authentication pipeline
        /// </summary>
        /// <param name="app">Owin app builder</param>
        public void ConfigureAuth(IAppBuilder app)
        {
            app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);

            app.UseCookieAuthentication(new CookieAuthenticationOptions { });

            app.UseOpenIdConnectAuthentication(
                new OpenIdConnectAuthenticationOptions
                {
                    ClientId = CSPApplicationId,
                    Authority = AuthorityCommon,
                    TokenValidationParameters = new Microsoft.IdentityModel.Tokens.TokenValidationParameters
                    {
                        // instead of using the default validation (validating against a single issuer value, as we do in line of business apps), 
                        // we inject our own multitenant validation logic
                        ValidateIssuer = false,
                        // If the app needs access to the entire organization, then add the logic
                        // of validating the Issuer here.
                        // IssuerValidator
                    },
                    Notifications = new OpenIdConnectAuthenticationNotifications()
                    {
                        SecurityTokenValidated = (context) =>
                        {
                            // If your authentication logic is based on users then add your logic here
                            return Task.FromResult(0);
                        },
                        AuthenticationFailed = (context) =>
                        {
                            // Pass in the context back to the app
                            context.OwinContext.Response.Redirect("/Home/Error");
                            context.HandleResponse(); // Suppress the exception
                            return Task.FromResult(0);
                        },
                        AuthorizationCodeReceived = HandleAuthorizationCodeReceivedNotification
                    }
                });
        }

        /// <summary>
        /// Recieve the AuthCode and exchange it for an access token and refresh token
        /// </summary>
        /// <param name="notificationMessage"></param>
        private static async Task HandleAuthorizationCodeReceivedNotification(AuthorizationCodeReceivedNotification notificationMessage)
        {
            string signInUserId = notificationMessage.AuthenticationTicket.Identity.FindFirst(ClaimTypes.Upn).Value;
            string partnerTenantId = notificationMessage.AuthenticationTicket.Identity.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            // Acquire a token using AuthCode
            Newtonsoft.Json.Linq.JObject tokenResult = await AuthorizationUtilities.GetAADTokenFromAuthCode(
                "https://login.microsoftonline.com/" + partnerTenantId,
                "https://api.partnercenter.microsoft.com",
                CSPApplicationId,
                CSPApplicationSecret,
                notificationMessage.Code,
                notificationMessage.OwinContext.Request.Uri.ToString());

            string refreshToken = tokenResult["refresh_token"].ToString();

            // Store the refresh token using partner tenant id as the key. 
            // Marketplace application will use the partner tenant id as a key to retrive the refresh token to get authenticated against the user.
            KeyVaultProvider provider = new KeyVaultProvider();
            await provider.AddSecretAsync(partnerTenantId, refreshToken);
        }
    }
}
