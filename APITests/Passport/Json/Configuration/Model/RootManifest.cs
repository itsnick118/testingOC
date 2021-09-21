using System;
using System.Collections.Generic;

namespace APITests.Passport.Json.Configuration.Model
{
    public class RootManifest
    {
        public Feedback Feedback { get; set; }
        public string SendLogEmail { get; set; }
        public bool UsePassportCmisObject { get; set; }
        public AutoUpdate AutoUpdate { get; set; }
        public Pane[] Panes { get; set; }
        public Dictionary<OfficeApp, AvailableAppDefinition[]> AvailableApps;
        public UserPreferences UserPreferences { get; set; }
        public string Name { get; set; }
        public string HelpUrl { get; set; }
        public Version OcVersion { get; set; }
        public string DefaultUrl { get; set; }
        public bool EnableJackrabbitFoldering { get; set; }
    }
}
