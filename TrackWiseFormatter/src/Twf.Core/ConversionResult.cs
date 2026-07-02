namespace Twf.Core;

/// <summary>The outcome of a single conversion run.</summary>
public sealed class ConversionResult
{
    public required string OutputPath { get; init; }
    public required SopDocument Document { get; init; }

    public IReadOnlyList<ExceptionRecord> Exceptions => Document.Exceptions;
    public bool HasErrors => Document.Exceptions.Any(e => e.Severity == Severity.Error);
}
