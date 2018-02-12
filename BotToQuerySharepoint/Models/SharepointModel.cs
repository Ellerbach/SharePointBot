using System;
using BotToQuerySharepoint.Services;

namespace BotToQuerySharepoint.Models
{
    [Serializable]
    public class SharepointModel
    {
        public string SitenameOrUrl { get; set; }
        public Uri Url { get; set; }
        public string Sitename { get; set; }
        public AccessRights AccessRights { get; set; }
        public string Token { get; set; }
        public UserInfo Owner { get; set; }
        public string SiteId { get; set; }
        public bool ChoicesGiven { get; set; }
    }

    public enum AccessRights
    {
        Read = 1,
        Write,
        FullPermissions
    }
}