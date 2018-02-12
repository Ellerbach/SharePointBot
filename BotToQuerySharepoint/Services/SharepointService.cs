using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using BotToQuerySharepoint.Models;
using Microsoft.Graph;

namespace BotToQuerySharepoint.Services
{
    public class SharepointService
    {
        public async Task<SharepointSiteValidationResponse> ValidateSite(string accessToken, Uri url)
        {
            var graphService = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                        return Task.FromResult(0);
                    }));

            var siteValidationResponse = new SharepointSiteValidationResponse();

            if (url.IsAbsoluteUri)
            {
                try
                {
                    var siteInfo = await graphService.Sites.GetByPath(string.Join(string.Empty, url.Segments.Skip(1).Take(3)), "root").Request().GetAsync();
                    siteValidationResponse.IsValid = true;
                    siteValidationResponse.MatchingUris = new List<Uri>() {new Uri(siteInfo.WebUrl)};
                    siteValidationResponse.SiteId = siteInfo.Id;
                }
                catch (Exception e)
                {
                    siteValidationResponse.IsValid = false;
                }
            }
            else
            {
                siteValidationResponse.IsValid = false;
            }

            return siteValidationResponse;

        }

        public async Task<SharepointSiteValidationResponse> ValidateSite(string token, string name, bool exactMatch= false)
        {
            var graphserviceClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                        return Task.FromResult(0);
                    }));

            var sites = await graphserviceClient.Sites.Request(new []{new QueryOption("search", name)}).GetAsync();

            var siteValidationResponse = new SharepointSiteValidationResponse();
            if (sites.Count >= 1)
            {
                if (exactMatch)
                {
                    siteValidationResponse.IsValid = true;
                    siteValidationResponse.SiteId = sites.First().Id;
                    siteValidationResponse.MatchingUris = new List<Uri>(){ new Uri(sites.First().WebUrl)};
                    siteValidationResponse.MatchingSites = new List<string>() {sites.First().DisplayName};
                }
                else
                {
                    siteValidationResponse.IsValid = true;
                    siteValidationResponse.SiteId = sites.Count > 1 ? null : sites.First().Id;
                    siteValidationResponse.MatchingUris = sites.CurrentPage.Select(i => new Uri(i.WebUrl)).ToList();
                    siteValidationResponse.MatchingSites = sites.CurrentPage.Select(i => i.DisplayName).ToList();
                }
            }
            else
            {
                siteValidationResponse.IsValid = false;
            }

            return siteValidationResponse;
        }

        public async void TestCalls(string accessToken)
        {
            var graphserviceClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                        return Task.FromResult(0);
                    }));
            //  var users = await graphserviceClient.Me.SendMail(new Message()).Request().GetAsync();



            //var siteInfo2 = await graphserviceClient.Sites[siteInfo.Id].Request().GetAsync();

            //var groupOwners = await graphserviceClient.Groups.Request().Filter("groupTypes/any(a:a%20eq%20'unified')").GetAsync();
            var groupOwners = await graphserviceClient.Groups.Request().Filter("startswith(displayName,'cool')").GetAsync();

            var x = groupOwners.Where(g => g.DisplayName.Contains("cool")).ToList();

            var groups = await graphserviceClient.Me.GetMemberGroups(true).Request().PostAsync();
            // var sites = await graphserviceClient.Sites.Request().GetAsync();
        }

        public async Task SendEmail(string accessToken, string emailAddress, string body, string subject)
        {
            var message = new Message();
            message.Body = new ItemBody() { Content = body, ContentType = BodyType.Text };
            message.ToRecipients = new List<Recipient>()
            {
                new Recipient() {EmailAddress = new EmailAddress() {Address = emailAddress}}
            };
            message.Subject = subject;

            var graphserviceClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                        return Task.FromResult(0);
                    }));

            await graphserviceClient.Me.SendMail(message, true).Request().PostAsync();
        }

   


        public async Task<UserInfo> GetOwnerNameForSite(string token, string siteId)
        {
            var graphService = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {

                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                        return Task.FromResult(0);
                    }));

            var s = await graphService.Sites[siteId]
                .Drives.Request()
                .GetAsync();

            var user = await graphService.Users[s.First().CreatedBy.User.Id].Request().GetAsync();
            return new UserInfo()
            {
                EmailAddress = user.Mail,
                Name = user.DisplayName
            };
        }

        private async Task<Site> GetSiteInformation(GraphServiceClient graphService, Uri uri)
        {
            Site siteInfo = null;
            if (uri.IsAbsoluteUri)
            {
                try
                {
                    siteInfo = await graphService.Sites
                        .GetByPath(string.Join(string.Empty, uri.Segments.Skip(1).Take(3)), "root").Request()
                        .GetAsync();
                }
                catch (Exception)
                {
                    siteInfo = null;
                }
            }

            return siteInfo;
        }
    }


}