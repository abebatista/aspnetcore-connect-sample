/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.AppConfig;
using MicrosoftGraphAspNetCoreConnectSample.Extensions;
using System;
using System.Threading.Tasks;

namespace MicrosoftGraphAspNetCoreConnectSample.Helpers
{
    public class GraphAuthProvider : IGraphAuthProvider
    {
        private IConfidentialClientApplication _app;
        private readonly string[] _scopes;

        public GraphAuthProvider(IConfiguration configuration)
        {
            var path = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
            if (String.IsNullOrEmpty(path))
            {
                path = "Production";
            }
            var azureOptions = new AzureAdOptions();
            configuration.Bind("AzureAd"+ path, azureOptions);

            // More info about MSAL Client Applications: https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/Client-Applications
            _app = ConfidentialClientApplicationBuilder.Create(azureOptions.ClientId)
                    .WithClientSecret(azureOptions.ClientSecret)
                    .WithAuthority(AzureCloudInstance.AzurePublic, AadAuthorityAudience.AzureAdAndPersonalMicrosoftAccount)
                    .WithRedirectUri(azureOptions.BaseUrl + azureOptions.CallbackPath)
                    .Build();

            _scopes = azureOptions.GraphScopes.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }

        // Gets an access token. First tries to get the access token from the token cache.
        // Using password (secret) to authenticate. Production apps should use a certificate.
        public async Task<string> GetUserAccessTokenAsync(string userId)
        {
            var account = await _app.GetAccountAsync(userId);
            if (account == null)
            {
                throw new ServiceException(new Error
                {
                    Code = "TokenNotFound",
                    Message = "User not found in token cache. Maybe the server was restarted."
                });
            }

            try
            {
                var result = await _app.AcquireTokenSilentAsync(_scopes, account);
                return result.AccessToken;
            }

            // Unable to retrieve the access token silently.
            catch (Exception)
            {
                throw new ServiceException(new Error
                {
                    Code = GraphErrorCode.AuthenticationFailure.ToString(),
                    Message = "Caller needs to authenticate. Unable to retrieve the access token silently."
                });
            }
        }

        public async Task<AuthenticationResult> GetUserAccessTokenByAuthorizationCode(string authorizationCode)
        {
            return await _app.AcquireTokenByAuthorizationCode(_scopes, authorizationCode).ExecuteAsync();
        }
    }

    public interface IGraphAuthProvider
    {
        Task<string> GetUserAccessTokenAsync(string userId);

        Task<AuthenticationResult> GetUserAccessTokenByAuthorizationCode(string authorizationCode);
    }
}