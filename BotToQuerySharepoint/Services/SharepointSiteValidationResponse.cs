using System;
using System.Collections.Generic;

namespace BotToQuerySharepoint.Services
{
    public class SharepointSiteValidationResponse
    {
        public bool IsValid { get; set; }
        public string SiteId { get; set; }
        public List<Uri> MatchingUris { get; set; }
        public List<string> MatchingSites { get; set; }
    }
}