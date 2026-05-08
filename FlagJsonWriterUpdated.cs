using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace Acc.ClientFlags.Plugin
{
    // ─────────────────────────────────────────────────────────────────────────
    // FlagJsonWriter
    //
    // Fires on Associate / Disassociate of acc_client_acc_clientflag.
    // Resolves each flag's /WebResources/... URL to the actual SVG markup
    // by querying the webresource table, then writes the full JSON
    // (with inline SVG) into acc_clientflagsjson on acc_client.
    // ─────────────────────────────────────────────────────────────────────────
    public class FlagJsonWriter : IPlugin
    {
        // ── Column names ──────────────────────────────────────────────────────
        private const string ClientEntity      = "acc_client";
        private const string ClientPrimaryKey  = "acc_clientid";
        private const string ClientJsonField   = "acc_clientflagsjson";

        private const string FlagEntity        = "acc_clientflag";
        private const string FlagPrimaryKey    = "acc_clientflagid";
        private const string FlagNameField     = "acc_name";
        private const string FlagIconUrlField  = "acc_iconurl";
        private const string FlagIsActiveField = "acc_isactive";

        // ── Static SVG cache ──────────────────────────────────────────────────
        // Web resource content never changes between deployments, so it is safe
        // to cache for the lifetime of the plugin instance. This means the
        // webresource query runs once per AppDomain, not once per plugin fire.
        private static Dictionary<string, string> _svgCache;
        private static readonly object _cacheLock = new object();

        // ─────────────────────────────────────────────────────────────────────
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)
                serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);
            var tracer  = (ITracingService)
                serviceProvider.GetService(typeof(ITracingService));

            try
            {
                var target = context.InputParameters.Contains("Target")
                    ? context.InputParameters["Target"] as EntityReference
                    : null;

                if (target == null || target.LogicalName != ClientEntity)
                {
                    tracer.Trace("Target is not acc_client — skipping.");
                    return;
                }

                Guid clientId = target.Id;
                tracer.Trace($"FlagJsonWriter firing for client: {clientId}");

                // 1. Load all active flag definitions (name + iconUrl)
                var allFlags = LoadAllFlagDefinitions(service, tracer);
                tracer.Trace($"Loaded {allFlags.Count} flag definitions.");

                // 2. Resolve every unique iconUrl to inline SVG markup.
                //    Uses the static cache — only queries webresource table
                //    on first run or when a new URL appears.
                var svgMap = ResolveSvgContent(
                    allFlags.Select(f => f.IconUrl).Distinct().ToList(),
                    service,
                    tracer);

                // 3. Load IDs of flags currently linked to this client
                var activeIds = LoadActiveFlagIds(service, clientId, tracer);
                tracer.Trace($"Client has {activeIds.Count} active flags.");

                // 4. Build JSON — iconUrl replaced with resolved SVG markup
                var flagArray = allFlags.Select(f => new FlagJsonEntry
                {
                    Id       = f.Id.ToString("D"),
                    Name     = f.Name,
                    // Store the SVG markup directly; fall back to the original
                    // URL string if resolution failed so the PCF can still
                    // attempt a client-side fetch as a last resort.
                    IconUrl  = svgMap.TryGetValue(f.IconUrl, out var svg) && !string.IsNullOrEmpty(svg)
                                 ? svg
                                 : f.IconUrl,
                    IsActive = activeIds.Contains(f.Id),
                }).ToList();

                string json = JsonConvert.SerializeObject(
                    flagArray,
                    Formatting.None,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                tracer.Trace($"Writing JSON ({json.Length} chars) to {ClientJsonField}.");

                var update = new Entity(ClientEntity, clientId);
                update[ClientJsonField] = json;
                service.Update(update);

                tracer.Trace("FlagJsonWriter completed successfully.");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(
                    $"FlagJsonWriter failed: {ex.Message}", ex);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // ResolveSvgContent
        //
        // Takes a list of iconUrl values from acc_clientflag records, e.g.:
        //   "/WebResources/acc_icon_violence.svg"
        //   "WebResources/acc_icon_absconding.svg"
        //
        // Strips the /WebResources/ prefix to get the web resource logical name,
        // queries the webresource table for any that aren't already cached,
        // decodes the base64 content field, and returns a map of
        //   iconUrl → SVG markup string
        // ─────────────────────────────────────────────────────────────────────
        private Dictionary<string, string> ResolveSvgContent(
            List<string> iconUrls,
            IOrganizationService service,
            ITracingService tracer)
        {
            // Result maps original iconUrl → SVG markup
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (iconUrls == null || iconUrls.Count == 0)
                return result;

            // Build a map of webResourceName → original iconUrl so we can
            // reverse-look up after querying
            var nameToUrl = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var url in iconUrls)
            {
                if (string.IsNullOrWhiteSpace(url)) continue;

                var wrName = ExtractWebResourceName(url);
                if (string.IsNullOrEmpty(wrName)) continue;

                if (!nameToUrl.ContainsKey(wrName))
                    nameToUrl[wrName] = url;
            }

            if (nameToUrl.Count == 0) return result;

            lock (_cacheLock)
            {
                if (_svgCache == null)
                    _svgCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // Find names not yet in the cache
                var missing = nameToUrl.Keys
                    .Where(n => !_svgCache.ContainsKey(n))
                    .ToList();

                if (missing.Count > 0)
                {
                    tracer.Trace($"Fetching {missing.Count} web resource(s) from Dataverse.");

                    // Single query for all missing names using an OR filter
                    var query = new QueryExpression("webresource")
                    {
                        ColumnSet = new ColumnSet("name", "content"),
                        Criteria  = new FilterExpression(LogicalOperator.Or),
                    };

                    foreach (var name in missing)
                        query.Criteria.AddCondition("name", ConditionOperator.Equal, name);

                    var wrResults = service.RetrieveMultiple(query);
                    tracer.Trace($"Retrieved {wrResults.Entities.Count} web resource record(s).");

                    foreach (var wr in wrResults.Entities)
                    {
                        var wrName  = wr.GetAttributeValue<string>("name");
                        var base64  = wr.GetAttributeValue<string>("content");

                        if (string.IsNullOrEmpty(wrName) || string.IsNullOrEmpty(base64))
                            continue;

                        try
                        {
                            var svgMarkup = Encoding.UTF8.GetString(Convert.FromBase64String(base64));
                            _svgCache[wrName] = svgMarkup.Trim();
                            tracer.Trace($"Cached SVG for: {wrName} ({svgMarkup.Length} chars)");
                        }
                        catch (Exception ex)
                        {
                            tracer.Trace($"Failed to decode content for {wrName}: {ex.Message}");
                            _svgCache[wrName] = string.Empty;
                        }
                    }

                    // Mark names that weren't found so we don't query again
                    foreach (var name in missing.Where(n => !_svgCache.ContainsKey(n)))
                    {
                        tracer.Trace($"Web resource not found: {name}");
                        _svgCache[name] = string.Empty;
                    }
                }

                // Build result map: original iconUrl → SVG markup
                foreach (var kvp in nameToUrl)
                {
                    if (_svgCache.TryGetValue(kvp.Key, out var svg))
                        result[kvp.Value] = svg;
                }
            }

            return result;
        }

        // ─────────────────────────────────────────────────────────────────────
        // ExtractWebResourceName
        //
        // Converts a URL path to the web resource logical name.
        //
        // "/WebResources/acc_icon_violence.svg"  →  "acc_icon_violence.svg"
        // "WebResources/acc_icon_violence.svg"   →  "acc_icon_violence.svg"
        // "acc_icon_violence.svg"                →  "acc_icon_violence.svg"
        // ─────────────────────────────────────────────────────────────────────
        private static string ExtractWebResourceName(string iconUrl)
        {
            if (string.IsNullOrWhiteSpace(iconUrl)) return string.Empty;

            var trimmed = iconUrl.Trim().TrimStart('/');

            // Strip "WebResources/" prefix (case-insensitive)
            const string prefix = "WebResources/";
            if (trimmed.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                trimmed = trimmed.Substring(prefix.Length);

            return trimmed.Trim();
        }

        // ── Load all active flag definitions ──────────────────────────────────
        private List<FlagDefinition> LoadAllFlagDefinitions(
            IOrganizationService service,
            ITracingService tracer)
        {
            var query = new QueryExpression(FlagEntity)
            {
                ColumnSet = new ColumnSet(FlagPrimaryKey, FlagNameField, FlagIconUrlField),
                Criteria  = new FilterExpression(LogicalOperator.And),
            };
            query.Criteria.AddCondition(FlagIsActiveField, ConditionOperator.Equal, true);
            query.AddOrder(FlagNameField, OrderType.Ascending);

            var results = service.RetrieveMultiple(query);
            return results.Entities.Select(e => new FlagDefinition
            {
                Id      = e.Id,
                Name    = e.GetAttributeValue<string>(FlagNameField)  ?? string.Empty,
                IconUrl = e.GetAttributeValue<string>(FlagIconUrlField)?.Trim() ?? string.Empty,
            }).ToList();
        }

        // ── Load IDs of flags currently linked to this client ─────────────────
        private HashSet<Guid> LoadActiveFlagIds(
            IOrganizationService service,
            Guid clientId,
            ITracingService tracer)
        {
            var query = new QueryExpression(FlagEntity)
            {
                ColumnSet = new ColumnSet(FlagPrimaryKey),
            };
            var link = query.AddLink(
                "acc_client_acc_clientflag",
                FlagPrimaryKey,
                FlagPrimaryKey);
            link.LinkCriteria.AddCondition(ClientPrimaryKey, ConditionOperator.Equal, clientId);

            var results = service.RetrieveMultiple(query);
            return new HashSet<Guid>(results.Entities.Select(e => e.Id));
        }

        // ── Internal DTOs ──────────────────────────────────────────────────────
        private class FlagDefinition
        {
            public Guid   Id      { get; set; }
            public string Name    { get; set; }
            public string IconUrl { get; set; }
        }

        private class FlagJsonEntry
        {
            [JsonProperty("id")]
            public string Id { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("iconUrl")]
            public string IconUrl { get; set; }

            [JsonProperty("isActive")]
            public bool IsActive { get; set; }
        }
    }
}
