using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Suite.Infrastructure
{
    [DataContract]
    public sealed class SuiteSettings
    {
        public SuiteSettings()
        {
            SchemaVersion = 1;
            Theme = "System";
            OutputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            Connections = new List<ServiceConnectionSettings>();
            DocumentColors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        [DataMember(Order = 1)] public int SchemaVersion { get; set; }
        [DataMember(Order = 2)] public string Theme { get; set; }
        [DataMember(Order = 3)] public string OutputDirectory { get; set; }
        [DataMember(Order = 4)] public List<ServiceConnectionSettings> Connections { get; set; }
        [DataMember(Order = 5)] public Dictionary<string, string> DocumentColors { get; set; }
    }

    [DataContract]
    public sealed class ServiceConnectionSettings
    {
        [DataMember(Order = 1)] public string Name { get; set; }
        [DataMember(Order = 2)] public string Provider { get; set; }
        [DataMember(Order = 3)] public string Endpoint { get; set; }
        [DataMember(Order = 4)] public string Model { get; set; }
        [DataMember(Order = 5)] public string SecretName { get; set; }
        [DataMember(Order = 6)] public bool Enabled { get; set; }
    }
}
