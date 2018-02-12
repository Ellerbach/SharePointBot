using System;
using System.Collections.Generic;
using System.Linq;
using BotToQuerySharepoint.Models;
using BotToQuerySharepoint.Services;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Builder.FormFlow.Advanced;
using Microsoft.Bot.Builder.Luis.Models;

namespace BotToQuerySharepoint.Forms
{
    [Serializable]
    public class SpFormHandler
    {
        public IFormDialog<SharepointModel> GetFormDialog(IEnumerable<EntityRecommendation> entities, string token)
        {
            var spForm = new SharepointModel();
            spForm.Token = token;
            foreach (EntityRecommendation entity in entities)
            {
                if (entity.Type == "sp-sitename" || entity.Type == "builtin.url")
                {
                    spForm.SitenameOrUrl = entity.Entity;
                }

                if (entity.Type == "sp-accessright")
                {
                    Enum.TryParse(entity.Entity, true, out AccessRights rights);
                    spForm.AccessRights = rights;
                }
            }

            return new FormDialog<SharepointModel>(spForm, BuildForm, FormOptions.PromptInStart);
        }

        private IForm<SharepointModel> BuildForm()
        {
            return new FormBuilder<SharepointModel>()
                .Field(new FieldReflector<SharepointModel>(nameof(SharepointModel.SitenameOrUrl))
                    .SetValidate(async (state, value) =>
                    {
                        if (string.Equals((string) value, "None", StringComparison.Ordinal))
                        {
                            return new ValidateResult
                            {
                                IsValid = false,
                                Feedback = "Sorry I couldn't find the site based on your input."
                            };
                        }

                        string siteNameOrUrl = value.ToString();
                        bool isUri = Uri.TryCreate(siteNameOrUrl, UriKind.Absolute, out Uri uri)
                                      && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);

                        var service = new SharepointService();

                        if (isUri)
                        {
                            var response = await service.ValidateSite(state.Token, uri);
                            if (response.IsValid && response.MatchingUris.Count == 1)
                            {
                                state.Url = uri;
                                state.SiteId = response.SiteId;
                                return new ValidateResult
                                {
                                    IsValid = response.IsValid,
                                    Value = value
                                };
                            }
                            else
                            {
                                return new ValidateResult
                                {
                                    IsValid = false,
                                    Feedback = "The site url is not valid."
                                };
                            }
                        }
                        else
                        {
                            var response = await service.ValidateSite(state.Token, value.ToString(), state.ChoicesGiven);
                            if (response.IsValid && response.MatchingUris.Count == 1)
                            {
                                state.Url = response.MatchingUris.First();
                                state.Sitename = siteNameOrUrl;
                                state.SiteId = response.SiteId;
                                return new ValidateResult()
                                {
                                    IsValid = response.IsValid,
                                    Value = value
                                };
                            }
                            else if (response.IsValid)
                            {
                                state.ChoicesGiven = true;
                                var x = new ValidateResult()
                                {
                                    IsValid = false,
                                    Choices = response.MatchingSites.Select(u => new Choice()
                                    {
                                        Value = u,
                                        Description = new DescribeAttribute(u),
                                        Terms = new TermsAttribute(u)
                                    }).Union(new[] { new Choice() { Value = "None", Description = new DescribeAttribute("None"), Terms = new TermsAttribute("None") } }),
                                    Value = value,
                                    Feedback = "Here are some sites I found."
                                };
                                return x;
                            }
                            else
                            {
                                return new ValidateResult()
                                {
                                    IsValid = response.IsValid,
                                    Feedback = "The name of site is not valid."
                                };
                            }
                        }
                    }))
                .Field(nameof(SharepointModel.AccessRights))
                .Confirm(async spModel => new PromptAttribute($"Do you confirm that this is the site you would like {spModel.AccessRights} access for: {spModel.Url.ToString()}?"))
                .Build();
        }
    }
}