using System;
using System.IO;
using NUnit.Framework;

namespace OnTraceTests
{
    /// <summary>
    /// Under the .NET Framework test host, the process config file is the test host's, not ours.
    /// So we redirect the application configuration file to this test assembly's config (which
    /// contains the log4net section) before the first OnTrace access, mirroring how OnTrace reads
    /// an application's app.config in production.
    /// </summary>
    [SetUpFixture]
    public class AssemblyInit
    {
        [OneTimeSetUp]
        public void RedirectAppConfig()
        {
            var configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OnTrace.Tests.dll.config");
            if (File.Exists(configFile))
            {
                AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", configFile);
            }
        }
    }
}
