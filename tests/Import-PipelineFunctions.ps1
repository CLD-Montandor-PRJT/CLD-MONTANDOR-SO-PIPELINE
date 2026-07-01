<#
.SYNOPSIS
    Test-harness loader: extracts named function definitions from a pipeline script via the
    PowerShell AST and defines them in the CALLER's scope, without executing the script's
    top-level code (credential reads, OAuth token acquisition, Graph/BC calls).
.DESCRIPTION
    Invoke-SalesOrderPipeline.ps1 and Watch-SalesOrderEmail.ps1 both perform real network/
    credential I/O unconditionally at script top-level (even in -FunctionsOnly mode for the
    former). Dot-sourcing either file directly would violate the "no BC calls, no Graph
    calls, no credentials required" constraint for offline unit tests. This loader parses
    the file's AST, isolates only the requested FunctionDefinitionAst nodes, and dot-sources
    just their text as scriptblocks — pulling in zero top-level side effects.
.PARAMETER ScriptPath
    Path to the .ps1 file containing the function(s).
.PARAMETER FunctionNames
    Names of the functions to extract and define in the caller's scope.
.NOTES
    MUST be invoked with the dot-source operator (. .\Import-PipelineFunctions.ps1 ...) so the
    functions it defines land in the caller's scope rather than a private child scope.
#>
param(
    [Parameter(Mandatory)][string]$ScriptPath,
    [Parameter(Mandatory)][string[]]$FunctionNames
)

$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) {
    throw "Parse errors in ${ScriptPath}: $($parseErrors -join '; ')"
}

foreach ($name in $FunctionNames) {
    $funcAst = $ast.FindAll(
        { param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name },
        $true
    ) | Select-Object -First 1
    if (-not $funcAst) { throw "Function '$name' not found in $ScriptPath" }
    . ([scriptblock]::Create($funcAst.Extent.Text))
}
