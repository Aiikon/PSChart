Import-Module (Get-Module PSChart).Path -Force


$dataSet1 = foreach ($v in 'A', 'B', 'C', 'D', 'E', 'F', 'G')
{
    [pscustomobject]@{ Label = $v; Category = 'Old'; Count = 1..10 | Get-Random }
}
$dataSet2 = foreach ($v in 'A', 'B', 'C', 'D', 'E', 'F', 'G')
{
    [pscustomobject]@{ Label = $v; Category = 'New'; Count = 1..10 | Get-Random }
}


Show-UIWindow -AddLoaded { $this.Activate() } {
    New-UIUniformGrid -Rows 1 -Columns 2 {
        $dataSet1 |
            New-PSChart -As WpfControl -Type Pie -XProperty Label -Title "Sample Pie Chart"

        & { $dataSet1; $dataSet2 } |
            New-PSChart -As WpfControl -Type StackedColumn -XProperty Label -ZProperty Category -Title "Sample Stacked Column Chart"
        
    }
}