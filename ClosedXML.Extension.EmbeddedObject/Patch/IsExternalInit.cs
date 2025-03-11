
// reference: https://blog.ndepend.com/using-c9-record-and-init-property-in-your-net-framework-4-x-net-standard-and-net-core-projects/
// .NET 5.0 added the new IsExternalInit static class in System.Runtime.dll, try to support the feature in netstandard2.0
// https://learn.microsoft.com/en-us/dotnet/api/system.runtime.compilerservices.isexternalinit?view=net-5.0

namespace ClassLibrary
{
    public record Class(string Str)
    {
        internal int Int { get; init; }
    }
}

namespace System.Runtime.CompilerServices
{
    using System.ComponentModel;
    /// <summary>
    /// Reserved to be used by the compiler for tracking metadata.
    /// This class should not be used by developers in source code.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal static class IsExternalInit
    {
    }
}
