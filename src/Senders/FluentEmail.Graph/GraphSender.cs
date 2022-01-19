using FluentEmail.Core;
using FluentEmail.Core.Interfaces;
using FluentEmail.Core.Models;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PublicClientApplication = Microsoft.Graph.PublicClientApplication;

namespace FluentEmail.Graph
{
    public class GraphSender : ISender
    {
        private readonly string _appId;
        private readonly string _tenantId;
        private readonly string _graphSecret;
        private bool _saveSent;

        private ClientCredentialProvider _clientAuthProvider;
        private DelegateAuthenticationProvider _publicAuthProvider;
        private GraphServiceClient _graphClient;
        private IConfidentialClientApplication _clientApp;
        
        public GraphSender(
            string GraphEmailAppId,
            string GraphEmailTenantId,
            string GraphEmailSecret,
            bool SaveSentItems)
        {
            _appId = GraphEmailAppId;
            _tenantId = GraphEmailTenantId;
            _graphSecret = GraphEmailSecret;
            _saveSent = SaveSentItems;

            _clientApp = ConfidentialClientApplicationBuilder
                .Create(_appId)
                .WithTenantId(_tenantId)
                .WithClientSecret(_graphSecret)
                .Build();

            _clientAuthProvider = new ClientCredentialProvider(_clientApp);

            _graphClient = new GraphServiceClient(_clientAuthProvider);
        }
        
        public GraphSender(string appId, string tenantId, bool saveSentItems)
        {
            _appId = appId;
            _tenantId = tenantId;
            _saveSent = saveSentItems;

            var pca = PublicClientApplicationBuilder
                .Create(_appId)
                .WithTenantId(_tenantId)
                .Build();

            // DelegateAuthenticationProvider is a simple auth provider implementation
            // that allows you to define an async function to retrieve a token
            // Alternatively, you can create a class that implements IAuthenticationProvider
            // for more complex scenarios
            Task.Run(() =>
            {
                _publicAuthProvider = new DelegateAuthenticationProvider(async (request) => {
                    // Use Microsoft.Identity.Client to retrieve token
                    var result = await new MSALNET.Token().Get();

                    request.Headers.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", result.AccessToken);
                });
            }).Wait();
            
            _graphClient = new GraphServiceClient(_publicAuthProvider);
        }

        public SendResponse Send(IFluentEmail email, CancellationToken? token = null)
        {
            return SendAsync(email, token).GetAwaiter().GetResult();
        }

        public async Task<SendResponse> SendAsync(IFluentEmail email, CancellationToken? token = null)
        {
            var message = new Message
            {
                Subject = email.Data.Subject,
                Body = new ItemBody
                {
                    Content = email.Data.Body,
                    ContentType = email.Data.IsHtml ? BodyType.Html : BodyType.Text
                },
                From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = email.Data.FromAddress.EmailAddress,
                        Name = email.Data.FromAddress.Name
                    }
                }
            };

            if(email.Data.ToAddresses != null && email.Data.ToAddresses.Count > 0)
            {
                var toRecipients = new List<Recipient>();

                email.Data.ToAddresses.ForEach(r => toRecipients.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    { 
                        Address = r.EmailAddress.ToString(),
                        Name = r.Name
                    }
                }));

                message.ToRecipients = toRecipients;
            }

            if(email.Data.BccAddresses != null && email.Data.BccAddresses.Count > 0)
            {
                var bccRecipients = new List<Recipient>();

                email.Data.BccAddresses.ForEach(r => bccRecipients.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = r.EmailAddress.ToString(),
                        Name = r.Name
                    }
                }));

                message.BccRecipients = bccRecipients;
            }

            if (email.Data.CcAddresses != null && email.Data.CcAddresses.Count > 0)
            {
                var ccRecipients = new List<Recipient>();

                email.Data.CcAddresses.ForEach(r => ccRecipients.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = r.EmailAddress.ToString(),
                        Name = r.Name
                    }
                }));

                message.CcRecipients = ccRecipients;
            }

            if(email.Data.Attachments != null && email.Data.Attachments.Count > 0)
            {
                message.Attachments = new MessageAttachmentsCollectionPage();

                email.Data.Attachments.ForEach(a =>
                {
                    var attachment = new FileAttachment
                    {
                        Name = a.Filename,
                        ContentType = a.ContentType,
                        ContentBytes = GetAttachmentBytes(a.Data)
                    };

                    message.Attachments.Add(attachment);
                });
            }

            switch(email.Data.Priority)
            {
                case Priority.High:
                    message.Importance = Importance.High;
                    break;
                case Priority.Normal:
                    message.Importance = Importance.Normal;
                    break;
                case Priority.Low:
                    message.Importance = Importance.Low;
                    break;
                default:
                    message.Importance = Importance.Normal;
                    break;
            }

            try
            {
                await _graphClient.Users[email.Data.FromAddress.EmailAddress]
                    .SendMail(message, _saveSent)
                    .Request()
                    .PostAsync();

                return new SendResponse
                {
                    MessageId = message.Id
                };
            }
            catch (Exception ex)
            {
                return new SendResponse
                {
                    ErrorMessages = new List<string> { ex.Message }
                };
            }
        }

        private static byte[] GetAttachmentBytes(Stream stream)
        {
            using(MemoryStream m = new MemoryStream())
            {
                stream.CopyTo(m);
                return m.ToArray();
            }
        }
    }
}
