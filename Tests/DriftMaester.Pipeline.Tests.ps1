#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Tests for the NUnit XML emitter in Invoke-DriftMaester.ps1. The runbook has script-level execution at the
# bottom, so we cannot dot-source it wholesale. Instead we extract just the emitter functions via the AST,
# which keeps the runbook as the single source of truth (no duplicated logic to drift out of sync).

BeforeAll {
    $runbookPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Runbooks\Invoke-DriftMaester.ps1'
    $runbookPath = (Resolve-Path -Path $runbookPath).Path

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($runbookPath, [ref]$tokens, [ref]$parseErrors)

    $functionsToLoad = @(
        'ConvertTo-XmlSafeText',
        'Get-MaesterTestService',
        'Get-MaesterTestHelpUrl',
        'Get-MaesterTestErrorText',
        'Export-MaesterNUnitXml'
    )

    foreach ($functionName in $functionsToLoad) {
        $definition = $ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
            }, $true) | Select-Object -First 1

        if (-not $definition) {
            throw "Could not find function '$functionName' in $runbookPath"
        }

        . ([scriptblock]::Create($definition.Extent.Text))
    }

    function New-FakeTest {
        param(
            [string] $Result,
            [string] $Title,
            [string] $Service,
            [string] $Severity = 'Medium',
            [string] $Evidence = '',
            [string] $SkippedReason = '',
            [string] $Id = ([guid]::NewGuid().ToString())
        )

        [PSCustomObject]@{
            Id           = $Id
            Name         = $Title
            Title        = $Title
            Result       = $Result
            Severity     = $Severity
            Block        = $Service
            HelpUrl      = 'https://maester.dev/docs'
            ErrorRecord  = $null
            ResultDetail = [PSCustomObject]@{
                Service         = $Service
                TestDescription = "Description for $Title"
                TestResult      = $Evidence
                SkippedReason   = $SkippedReason
            }
        }
    }

    $script:tests = @(
        New-FakeTest -Result 'Passed' -Title 'Passed test A' -Service 'Entra'
        New-FakeTest -Result 'Failed' -Title 'Failed test B' -Service 'Exchange' -Severity 'High' -Evidence 'Found 3 bad things'
        New-FakeTest -Result 'Error' -Title 'Error test C' -Service 'Teams'
        New-FakeTest -Result 'Investigate' -Title 'Investigate test D' -Service 'SharePoint'
        New-FakeTest -Result 'Skipped' -Title 'Skipped test E' -Service 'Intune' -SkippedReason 'Not licensed'
        New-FakeTest -Result 'NotRun' -Title 'NotRun test F' -Service 'Defender'
        New-FakeTest -Result 'Failed' -Title 'Failed test B' -Service 'Exchange'
        New-FakeTest -Result 'Failed' -Title '<b>Bad &amp; ugly ]]> test</b>' -Service 'Graph' -Evidence 'weird <xml> & ]]> content'
    )

    $script:result = [PSCustomObject]@{
        TenantId   = '11111111-1111-1111-1111-111111111111'
        TenantName = 'Contoso'
        Tests      = $script:tests
    }

    $script:xmlPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("driftmaester-nunit-{0}.xml" -f [guid]::NewGuid().ToString('N'))
    Export-MaesterNUnitXml -Result $script:result -Path $script:xmlPath

    [xml] $script:doc = Get-Content -Path $script:xmlPath -Raw
}

AfterAll {
    if ($script:xmlPath -and (Test-Path -Path $script:xmlPath)) {
        Remove-Item -Path $script:xmlPath -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Export-MaesterNUnitXml' {
    It 'produces a well-formed XML document with an NUnit test-results root' {
        $script:doc | Should -Not -BeNullOrEmpty
        $script:doc.DocumentElement.LocalName | Should -Be 'test-results'
        $script:doc.DocumentElement.GetAttribute('name') | Should -Be 'DriftMaester'
    }

    It 'reports the correct summary counts on the root element' {
        $root = $script:doc.SelectSingleNode('/test-results')
        $root.GetAttribute('total') | Should -Be '8'
        $root.GetAttribute('failures') | Should -Be '3'
        $root.GetAttribute('errors') | Should -Be '1'
        $root.GetAttribute('inconclusive') | Should -Be '1'
        $root.GetAttribute('skipped') | Should -Be '1'
        $root.GetAttribute('ignored') | Should -Be '1'
        $root.GetAttribute('not-run') | Should -Be '2'
    }

    It 'names the test-suite after the tenant and marks it failed when there are findings' {
        $suite = $script:doc.SelectSingleNode('/test-results/test-suite')
        $suite.GetAttribute('name') | Should -Be 'Contoso'
        $suite.GetAttribute('result') | Should -Be 'Failure'
        $suite.GetAttribute('success') | Should -Be 'False'
    }

    It 'emits one test-case per test' {
        $script:doc.SelectNodes('//test-case').Count | Should -Be 8
    }

    It 'maps a passed test to Success with no failure child' {
        $case = $script:doc.SelectSingleNode("//test-case[@name='Entra.Passed test A']")
        $case.GetAttribute('result') | Should -Be 'Success'
        $case.GetAttribute('success') | Should -Be 'True'
        $case.GetAttribute('executed') | Should -Be 'True'
        $case.SelectSingleNode('failure') | Should -BeNullOrEmpty
    }

    It 'maps a failed test to Failure with a message containing severity, service and evidence' {
        $case = $script:doc.SelectSingleNode("//test-case[@name='Exchange.Failed test B']")
        $case.GetAttribute('result') | Should -Be 'Failure'
        $case.GetAttribute('success') | Should -Be 'False'
        $message = $case.SelectSingleNode('failure/message').InnerText
        $message | Should -Match 'Severity: High'
        $message | Should -Match 'Service: Exchange'
        $message | Should -Match 'Found 3 bad things'
    }

    It 'maps an error test to Error with a failure child' {
        $case = $script:doc.SelectSingleNode("//test-case[@name='Teams.Error test C']")
        $case.GetAttribute('result') | Should -Be 'Error'
        $case.SelectSingleNode('failure') | Should -Not -BeNullOrEmpty
    }

    It 'maps an investigate test to Inconclusive with a reason' {
        $case = $script:doc.SelectSingleNode("//test-case[@name='SharePoint.Investigate test D']")
        $case.GetAttribute('result') | Should -Be 'Inconclusive'
        $case.SelectSingleNode('reason/message') | Should -Not -BeNullOrEmpty
    }

    It 'maps a skipped test to a not-executed Ignored case carrying the skipped reason' {
        $case = $script:doc.SelectSingleNode("//test-case[@name='Intune.Skipped test E']")
        $case.GetAttribute('result') | Should -Be 'Ignored'
        $case.GetAttribute('executed') | Should -Be 'False'
        $case.SelectSingleNode('reason/message').InnerText | Should -Match 'Not licensed'
    }

    It 'de-duplicates identical test names' {
        $script:doc.SelectNodes("//test-case[@name='Exchange.Failed test B']").Count | Should -Be 1
        $script:doc.SelectNodes("//test-case[@name='Exchange.Failed test B (2)']").Count | Should -Be 1
    }

    It 'escapes special characters so evidence with markup stays valid XML' {
        $cases = @($script:doc.SelectNodes('//test-case') | Where-Object { $_.GetAttribute('name').StartsWith('Graph.') })
        $cases.Count | Should -Be 1
        $cases[0].SelectSingleNode('failure/message').InnerText | Should -Match 'weird <xml> & \]\]> content'
    }
}

Describe 'ConvertTo-XmlSafeText' {
    It 'strips characters that are illegal in XML 1.0' {
        $value = "before$([char]0x00)$([char]0x1F)after"
        ConvertTo-XmlSafeText -Value $value | Should -Be 'beforeafter'
    }

    It 'keeps tabs, newlines and normal text' {
        $value = "line1`tcol`r`nline2"
        ConvertTo-XmlSafeText -Value $value | Should -Be $value
    }

    It 'returns an empty string for null' {
        ConvertTo-XmlSafeText -Value $null | Should -Be ''
    }
}
