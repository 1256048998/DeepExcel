using System;
using System.Collections.Generic;
using System.Text;

namespace DeepExcel.Repair
{
    /// <summary>
    /// Stable identities for every condition that can stop the add-in loading.
    /// The code is what a user reports and what support searches for, so the
    /// string values must never be reused for a different meaning.
    /// </summary>
    public static class Codes
    {
        public const string ComClsid64 = "E-REG-001";
        public const string ComClsid32 = "E-REG-002";
        public const string ComTaskPane = "E-REG-003";
        public const string LoadBehavior = "E-REG-004";
        public const string HklmResidual = "E-REG-005";
        public const string CodeBaseStale = "E-REG-006";

        public const string Activation64 = "E-ACT-001";
        public const string Activation32 = "E-ACT-002";

        public const string DotNet48 = "E-ENV-001";
        public const string WebView2 = "E-ENV-002";
        public const string EmbeddedPython = "E-ENV-003";
        public const string SidecarScript = "E-ENV-004";
        public const string AddInDll = "E-ENV-005";

        /// <summary>Excel loaded the add-in but its constructor threw.</summary>
        public const string LastLoadFailed = "E-LOAD-001";
        // E-LOAD-002/003, E-SIDE-*, E-ENG-* are taken by the add-in's startup
        // telemetry (DeepExcel.AddIn.Account.StartupErrorCodes). Do not reuse.

        public const string DisabledItems = "E-RES-001";
        public const string CrashingAddinList = "E-RES-002";
        public const string MarkOfTheWeb = "E-RES-003";
        public const string OrphanSidecar = "E-RES-004";
    }

    public enum Severity
    {
        /// <summary>Add-in cannot load until this is fixed.</summary>
        Blocking,

        /// <summary>Degrades the product but the ribbon still appears.</summary>
        Warning
    }

    public sealed class Finding
    {
        public Finding(string code, Severity severity, string summary, bool repairable)
        {
            Code = code;
            Severity = severity;
            Summary = summary;
            Repairable = repairable;
        }

        public string Code { get; }
        public Severity Severity { get; }
        public string Summary { get; }
        public bool Repairable { get; }

        /// <summary>Set by the repair pass so the report can say what happened.</summary>
        public string RepairOutcome { get; set; }

        public override string ToString()
        {
            var line = string.Format(
                "[{0}] {1} {2}",
                Severity == Severity.Blocking ? "BLOCK" : "WARN ",
                Code,
                Summary);
            if (!string.IsNullOrEmpty(RepairOutcome))
            {
                line += "  -> " + RepairOutcome;
            }
            return line;
        }
    }

    public sealed class Report
    {
        private readonly List<Finding> _findings = new List<Finding>();

        public IReadOnlyList<Finding> Findings { get { return _findings; } }

        public void Add(Finding finding)
        {
            _findings.Add(finding);
        }

        public bool HasBlocking
        {
            get
            {
                foreach (var finding in _findings)
                {
                    if (finding.Severity == Severity.Blocking)
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        /// <summary>Blocking findings that no automated repair can clear.</summary>
        public bool HasUnrepairableBlocking
        {
            get
            {
                foreach (var finding in _findings)
                {
                    if (finding.Severity == Severity.Blocking && !finding.Repairable)
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        public string Render()
        {
            if (_findings.Count == 0)
            {
                return "DeepExcel 自检通过，未发现问题。";
            }

            var builder = new StringBuilder();
            builder.AppendLine(string.Format("发现 {0} 个问题：", _findings.Count));
            foreach (var finding in _findings)
            {
                builder.AppendLine("  " + finding);
            }
            return builder.ToString();
        }
    }
}
