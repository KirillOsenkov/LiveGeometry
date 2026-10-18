using System;

namespace LiveGeometry;

public class ValidationResult
{
    public bool IsValid { get; set; }
    public object Value { get; set; }
    public string Error { get; set; }

    /// <summary>The caption of a button under the error that puts it right ("Swap names with A"); null for none</summary>
    public string FixCaption { get; set; }

    /// <summary>What that button does</summary>
    public Action Fix { get; set; }
}