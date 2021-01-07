// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// Generated with Bot Builder V4 SDK Template for Visual Studio EchoBot v4.6.2

using System.Collections.Generic;
using System.Threading;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema.Teams;
using System;
using System.IdentityModel.Tokens.Jwt;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Tokens;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using AdaptiveCards;
//using System.Security.Cryptography.HashAlgorithm;

namespace HelloWorldBot.Bots
{
    public class EchoBot : TeamsActivityHandler
    {
        protected override async Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
        {
            if(turnContext != null && turnContext.Activity != null)
            {
                JObject valueObject = JObject.FromObject(turnContext.Activity.Value);
                
                if (valueObject["authentication"] != null)
                {
                    JObject test = JObject.FromObject(valueObject["authentication"]);
                    foreach (JProperty property in test.Properties())
                    {
                        Console.WriteLine(property.Name + " - " + property.Value);
                    }
                    // name1 - value1
                    // name2 - value2

                    foreach (KeyValuePair<string, JToken> property in test)
                    {
                        Console.WriteLine(property.Key + " - " + property.Value);
                    }
                    string accessToken = GetPostTransformedPFTToken((test["token"]).ToString());
                    string actorToken = await getActorToken();
                    string spMetadata = await getSharePointMetadata(accessToken, actorToken, "");

                    string card = "{\r\n      \"type\": \"AdaptiveCard\",\r\n      \"msTeams\": {\r\n        \"width\": \"full\"\r\n      },\r\n      \"body\": [\r\n                {\r\n                    \"type\": \"Container\",\r\n                    \"items\": [\r\n                        {\r\n                            \"type\": \"ColumnSet\",\r\n                            \"columns\": [\r\n                                \r\n                                {\r\n                                    \"type\": \"Column\",\r\n                                    \"width\": \"auto\",\r\n                                    \"items\": [\r\n                                        {\r\n                                            \"type\": \"Container\",\r\n                                            \"items\": [\r\n                                                {\r\n                                                    \"type\": \"TextBlock\",\r\n                                                    \"text\": \"Leadership Connection\",\r\n                                                    \"weight\": \"Bolder\",\r\n                                                    \"fontType\": \"Default\",\r\n                                                    \"size\": \"small\",\r\n                                                    \"spacing\": \"None\"\r\n                                                },\r\n                                                {\r\n                                                    \"type\": \"TextBlock\",\r\n                                                    \"text\": \"Singapore building update\",\r\n                                                    \"fontType\": \"Default\",\r\n                                                    \"size\": \"Large\",\r\n                                                    \"spacing\": \"default\"\r\n                                                },\r\n                                                {\r\n                                                    \"type\": \"ColumnSet\",\r\n                                                    \"columns\": [\r\n                                                        {\r\n                                                            \"type\": \"Column\",\r\n                                                            \"width\": \"stretch\",\r\n                                                            \"items\": [\r\n                                                                {\r\n                                                                    \"type\": \"TextBlock\",\r\n                                                                    \"text\": \" \",\r\n                                                                    \"fontType\": \"Default\",\r\n                                                                    \"isSubtle\": true,\r\n                                                                    \"size\": \"Small\",\r\n                                                                    \"spacing\": \"Large\"\r\n                                                                },\r\n                                                                {\r\n                                                                    \"type\": \"TextBlock\",\r\n                                                                    \"text\": \"Patti Fernandez\",\r\n                                                                    \"fontType\": \"Default\",\r\n                                                                    \"isSubtle\": true,\r\n                                                                    \"size\": \"Small\",\r\n                                                                    \"spacing\": \"medium\"\r\n                                                                },\r\n                                                                {\r\n                                                                    \"type\": \"TextBlock\",\r\n                                                                    \"text\": \"Aug 25, 2020\",\r\n                                                                    \"fontType\": \"Default\",\r\n                                                                    \"isSubtle\": true,\r\n                                                                    \"size\": \"Small\",\r\n                                                                    \"spacing\": \"None\"\r\n                                                                }\r\n                                                            ]\r\n                                                        }\r\n                                                    ]\r\n                                                }\r\n                                            ]\r\n                                        }\r\n                                    ]\r\n                                }\r\n                            ]\r\n                        }\r\n                    ]\r\n                }\r\n            ],\r\n      \"selectAction\": {\r\n        \"type\": \"Action.OpenUrl\",\r\n        \"url\": \"https://www.youtube.com/watch?v=YPlIhgAX9AQ\"\r\n      },\r\n      \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\",\r\n      \"version\": \"1.2\"\r\n    }";

                    var parsedResult = AdaptiveCard.FromJson(card);
                    var attachment = new MessagingExtensionAttachment
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = parsedResult.Card,
                    };

                    var result = new MessagingExtensionResult("list", "result", new[] { attachment });

                    return new MessagingExtensionResponse(result);
                }
                else
                {
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = new MessagingExtensionResult
                        {
                            Type = "silentAuth"
                        }
                    };
                }
            }
            return null;
        }

        private string GetPostTransformedPFTToken(String preTransformedPFTToken)
        {
            JwtSecurityToken jwtSecurityToken = new JwtSecurityToken(preTransformedPFTToken);
            JObject headerObject = JObject.Parse(Base64UrlEncoder.Decode(jwtSecurityToken.RawHeader));

            string nonceIn = headerObject.GetValue("nonce", StringComparison.OrdinalIgnoreCase).ToString();
            string hashName = headerObject.GetValue("alg", StringComparison.OrdinalIgnoreCase).ToString();

            HashAlgorithm hashAlgorithm = EchoBot.GetHashFunction(hashName);
            string nonceOut = Base64UrlEncoder.Encode(hashAlgorithm.ComputeHash(Encoding.UTF8.GetBytes(nonceIn)));
            headerObject["nonce"] = nonceOut;
            string newHeaderString = Base64UrlEncoder.Encode(headerObject.ToString(Formatting.None));
            string postTransformedToken = $"{newHeaderString}.{jwtSecurityToken.RawPayload}.{jwtSecurityToken.RawSignature}";
            return postTransformedToken;
        }

        private static HashAlgorithm GetHashFunction(string hashAlgorithmName)
        {
            switch (hashAlgorithmName)
            {
                case "RS256":
                    return SHA256.Create();
                default:
                    string errorMessage = $"Algorithm [{hashAlgorithmName}] not supported for PFT at this time";
                    throw new NotSupportedException(errorMessage);
            }
        }

        private async Task<string> getActorToken()
        {
            string url = "https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a/oauth2/v2.0/token";
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(url);
            string clientAssertion = GetSignedClientAssertion();

            var values = new Dictionary<string, string>
            {
                { "grant_type", "client_credentials" },
                { "thing2", "e5e15768-1702-474d-ba7b-904c7cad2bcf" },
                {"client_assertion", clientAssertion },
                {"scope", "https://microsoft.sharepoint.com/.default" },
                {"client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer" }
            };

            client.DefaultRequestHeaders.Add("Accept", "application/json");

            var content = new FormUrlEncodedContent(values);
            var response = await client.PostAsync("https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a/oauth2/v2.0/token", content);
            var responseString = await response.Content.ReadAsStringAsync();
            JObject test = JObject.Parse(responseString);
            return ((test["access_token"]).ToString());
        }

        // Code to fetch the Client Assertion from KeyVault
        private string GetSignedClientAssertion()
        {
            X509Certificate2 selfSignedCertificate = new X509Certificate2(@"C:\Users\riagarwa\Downloads\oct5keyvault-ProdCert2-20201211.pfx", "", X509KeyStorageFlags.EphemeralKeySet);

            // AAD Prod Tenant ID: f8cdef31-a31e-4b4a-93e4-5f571e91255a
            string aud = $"https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a/v2.0";

            // client_id
            string clientID = "e5e15768-1702-474d-ba7b-904c7cad2bcf";

            // no need to add exp, nbf as JsonWebTokenHandler will add them by default.
            var claims = new Dictionary<string, object>()
            {
                { "aud", aud },
                { "iss", clientID },
                { "jti", Guid.NewGuid().ToString() },
                { "sub", clientID }
            };

            var securityTokenDescriptor = new SecurityTokenDescriptor
            {
                Claims = claims,
                SigningCredentials = new X509SigningCredentials(selfSignedCertificate)
            };

            JsonWebTokenHandler handler = new JsonWebTokenHandler();
            string signedClientAssertion = handler.CreateToken(securityTokenDescriptor);
            return signedClientAssertion;
        }

        private async Task<string> getSharePointMetadata(string accessToken, string actorToken, string url)
        {
            /*HttpClient client = new HttpClient();
            string authorizationRequestHeaderValue = "MSAuth1.0 actortoken=Bearer " + actorToken + ", accesstoken=Bearer " + accessToken + ", type=PFAT";

            client.DefaultRequestHeaders.Add("Authorization", authorizationRequestHeaderValue);
            HttpResponseMessage response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();
            var resp = await response.Content.ReadAsStringAsync();

            return resp;*/
            return null;
        }
    }
}
