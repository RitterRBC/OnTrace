using System;
using System.Linq;
using log4net;
using log4net.Appender;
using log4net.Repository.Hierarchy;
using NUnit.Framework;

namespace OnTraceTests
{
    /// <summary>
    /// These tests exercise the routing/level logic of OnTrace through real log4net
    /// MemoryAppenders that are wired up in App.config.
    /// </summary>
    [TestFixture]
    public class OnTraceRoutingTests
    {
        private static MemoryAppender AppenderFor(string loggerName, string appenderName)
        {
            var logger = (Logger)LogManager.GetLogger(loggerName).Logger;
            return (MemoryAppender)logger.GetAppender(appenderName);
        }

        [SetUp]
        public void BeforeEach()
        {
            // Force enablement so tests do not depend on config reading, and make sure log4net is
            // configured (idempotent) so the MemoryAppenders are attached before we clear them.
            OnTrace.AppLoggerEnabled = true;
            OnTrace.EventLoggerEnabled = true;
            OnTrace.StartupLoggerEnabled = true;
            OnTrace.PatchAppenders();
        }

        private static void ClearAll()
        {
            AppenderFor("AppLogger", "MemoryApp").Clear();
            AppenderFor("EventLogger", "MemoryEvent").Clear();
            AppenderFor("StartupLogger", "MemoryStartup").Clear();
        }

        [Test]
        public void Error_IsAlwaysWritten_RegardlessOfTraceLevel()
        {
            OnTrace.TraceLevel = (int)OnTrace.enmTraceLevel.Level_3; // mask = 4
            ClearAll();

            // Level_1 (1) is NOT part of the mask (4), but errors must still be written.
            OnTrace.TraceError("fatal-condition", OnTrace.enmTraceLevel.Level_1, OnTrace.enmTarget.AppLog);

            var events = AppenderFor("AppLogger", "MemoryApp").GetEvents();
            Assert.That(events.Any(e => e.MessageObject.ToString().Contains("fatal-condition")),
                "An error message must be written even when its trace level is not enabled.");
        }

        [Test]
        public void FatalError_IsAlwaysWritten_RegardlessOfTraceLevel()
        {
            OnTrace.TraceLevel = (int)OnTrace.enmTraceLevel.Level_3; // mask = 4
            ClearAll();

            OnTrace.TraceFatal("fatal-condition", OnTrace.enmTraceLevel.Level_2, OnTrace.enmTarget.AppLog);

            Assert.That(AppenderFor("AppLogger", "MemoryApp").GetEvents().Any(e => e.MessageObject.ToString().Contains("fatal-condition")),
                "A fatal message must be written even when its trace level is not enabled.");
        }

        [Test]
        public void NonError_IsFilteredByTraceLevel()
        {
            OnTrace.TraceLevel = (int)OnTrace.enmTraceLevel.Level_3; // mask = 4
            ClearAll();

            // Level_1 (1) not in mask (4) -> must NOT be written.
            OnTrace.TraceInformation("should-be-filtered", OnTrace.enmTraceLevel.Level_1, OnTrace.enmTarget.AppLog);
            // Level_3 (4) in mask -> must be written.
            OnTrace.TraceInformation("should-be-written", OnTrace.enmTraceLevel.Level_3, OnTrace.enmTarget.AppLog);

            var events = AppenderFor("AppLogger", "MemoryApp").GetEvents().Select(e => e.MessageObject.ToString()).ToList();
            Assert.That(events, Does.Not.Contain("should-be-filtered"));
            Assert.That(events, Does.Contain("should-be-written"));
        }

        [Test]
        public void TraceErrorFormat_RespectsEventLogTarget()
        {
            // Regression test for the bug where TraceErrorFormat/TraceFatalFormat ignored the
            // enmTarget argument and always wrote to the AppLog.
            OnTrace.TraceLevel = (int)OnTrace.enmTraceLevel.Level_All;
            ClearAll();

            OnTrace.TraceErrorFormat("code {0}", 42, OnTrace.enmTraceLevel.Level_All, OnTrace.enmTarget.EventLog);

            var appEvents = AppenderFor("AppLogger", "MemoryApp").GetEvents().Select(e => e.MessageObject.ToString()).ToList();
            var eventEvents = AppenderFor("EventLogger", "MemoryEvent").GetEvents().Select(e => e.MessageObject.ToString()).ToList();

            Assert.That(eventEvents, Does.Contain("code 42"),
                "Error routed to the EventLog target must land in the EventLog logger.");
            Assert.That(appEvents, Does.Not.Contain("code 42"),
                "Error routed to the EventLog target must NOT land in the AppLog logger.");
        }

        [Test]
        public void TraceFatalFormat_RespectsStartupLogTarget()
        {
            OnTrace.TraceLevel = (int)OnTrace.enmTraceLevel.Level_All;
            ClearAll();

            OnTrace.TraceFatalFormat("fatal {0}", 7, OnTrace.enmTraceLevel.Level_All, OnTrace.enmTarget.StartupLog);

            var startupEvents = AppenderFor("StartupLogger", "MemoryStartup").GetEvents().Select(e => e.MessageObject.ToString()).ToList();
            var appEvents = AppenderFor("AppLogger", "MemoryApp").GetEvents().Select(e => e.MessageObject.ToString()).ToList();

            Assert.That(startupEvents, Does.Contain("fatal 7"));
            Assert.That(appEvents, Does.Not.Contain("fatal 7"));
        }

        [Test]
        public void GetErrorStr_IncludesMessageAndInnerException()
        {
            var ex = new InvalidOperationException("outer", new ArgumentException("inner"));
            var result = OnTrace.GetErrorStr(ex, "ctx");

            Assert.That(result, Does.Contain("ctx"));
            Assert.That(result, Does.Contain("outer"));
            Assert.That(result, Does.Contain("inner"));
        }
    }
}
