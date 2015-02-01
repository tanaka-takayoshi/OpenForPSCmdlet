// Guids.cs
// MUST match guids.h
using System;

namespace tanaka_733.OpenForPSCmdlet
{
    static class GuidList
    {
        public const string guidOpenForPSCmdletPkgString = "192124ca-9705-4a99-97c4-1aacada178f6";
        public const string guidOpenForPSCmdletCmdSetString = "3f10aafe-0ce2-4379-b203-d809fc9169aa";

        public static readonly Guid guidOpenForPSCmdletCmdSet = new Guid(guidOpenForPSCmdletCmdSetString);
    };
}