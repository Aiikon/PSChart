Import-Module $PSScriptRoot -Force

$width = 500
$height = 400

$valueSet1 = 1,0,2,3,4,6,5,7,8,9
$dataSet1 = for ($i = 0; $i -lt $valueSet1.Count; $i++)
{
    [pscustomobject]@{
        Label = "Series $([string][char]([int][char]"A"+$i))"
        Count = $valueSet1[$i]
    }
}

$valueSet2 = 'January', 'February', 'March', 'April', 'May'
$dataSet2 = for ($i = 0; $i -lt $valueSet1.Count; $i++)
{
    for ($j = 0; $j -lt $valueSet2.Count; $j++)
    {
        [pscustomobject]@{
            Label = "Series $([string][char]([int][char]"A"+$i))"
            Month = $valueSet2[$j]
            Count = $valueSet1[($i+$j*2)%$valueSet1.Count]
        }
    }
}

$defaults1 = @{}
$defaults1.Width = $width
$defaults1.Height = $height
$defaults1.XProperty = 'Label'
$defaults1.YProperty = 'Count'
$defaults1.As = 'ImgTag'

$defaultsPie = @{} + $defaults1
$defaultsPie.Type = 'Pie'

$defaultsColumn = @{} + $defaults1
$defaultsColumn.Type = 'Column'

$defaults2 = @{} + $defaults1
$defaults2.ZProperty = 'Month'

@(
    "<h1>Baseline</h1>"

    $dataSet1 | New-PSChart @defaultsPie
    $dataSet1 | New-PSChart @defaultsPie -Title "Chart Title"

    "<h1>Single Series Types</h1>"

    $dataset1 | New-PSChart @defaults1 -Type Area -Title "Type Area"
    $dataset1 | New-PSChart @defaults1 -Type Bar -Title "Type Bar"
    $dataset1 | New-PSChart @defaults1 -Type Column -Title "Type Column"
    $dataset1 | New-PSChart @defaults1 -Type Doughnut -Title "Type Doughnut"
    $dataset1 | New-PSChart @defaults1 -Type FastLine -Title "Type FastLine"
    $dataset1 | New-PSChart @defaults1 -Type Funnel -Title "Type Funnel"
    $dataset1 | New-PSChart @defaults1 -Type Kagi -Title "Type Kagi"
    $dataset1 | New-PSChart @defaults1 -Type Pyramid -Title "Type Pyramid"
    $dataset1 | New-PSChart @defaults1 -Type Radar -Title "Type Polar"
    $dataset1 | New-PSChart @defaults1 -Type Renko -Title "Type Renko"
    $dataset1 | New-PSChart @defaults1 -Type Spline -Title "Type Spline"
    $dataset1 | New-PSChart @defaults1 -Type SplineArea -Title "Type SplineArea"
    $dataset1 | New-PSChart @defaults1 -Type StepLine -Title "Type StepLine"
    $dataset1 | New-PSChart @defaults1 -Type ThreeLineBreak -Title "Type ThreeLineBreak"

    "<h1>Two Series Types</h1>"

    $dataset2 | New-PSChart @defaults2 -Type Bar -Title "Type Bar"
    $dataset2 | New-PSChart @defaults2 -Type Column -Title "Type Column"
    $dataset2 | New-PSChart @defaults2 -Type StackedArea -Title "Type StackedArea"
    $dataset2 | New-PSChart @defaults2 -Type StackedArea100 -Title "Type StackedArea100"
    $dataset2 | New-PSChart @defaults2 -Type StackedBar -Title "Type StackedBar"
    $dataset2 | New-PSChart @defaults2 -Type StackedBar100 -Title "Type StackedBar100"
    $dataset2 | New-PSChart @defaults2 -Type StackedColumn -Title "Type StackedColumn"
    $dataset2 | New-PSChart @defaults2 -Type StackedColumn100 -Title "Type StackedColumn100"

    "<h1>Legend Position</h1>"

    $dataSet1 | New-PSChart @defaultsPie -Title "Pie Legend Right" -LegendPosition Right
    $dataSet1 | New-PSChart @defaultsPie -Title "Pie Legend Bottom" -LegendPosition Bottom
    $dataSet1 | New-PSChart @defaultsPie -Title "Pie Legend Left" -LegendPosition Left
    $dataSet1 | New-PSChart @defaultsPie -Title "Pie Legend Top" -LegendPosition Top
    $dataSet1 | New-PSChart @defaultsPie -Title "Pie Legend None" -LegendPosition None
    
    $dataSet2 | New-PSChart @defaultsColumn -ZProperty Month -Title "Column Legend Right" -LegendPosition Right
    $dataSet2 | New-PSChart @defaultsColumn -ZProperty Month -Title "Column Legend Bottom" -LegendPosition Bottom
    $dataSet2 | New-PSChart @defaultsColumn -ZProperty Month -Title "Column Legend Left" -LegendPosition Left
    $dataSet2 | New-PSChart @defaultsColumn -ZProperty Month -Title "Column Legend Top" -LegendPosition Top
    $dataSet2 | New-PSChart @defaultsColumn -ZProperty Month -Title "Column Legend None" -LegendPosition None

    "<h1>Assorted Styles</h1>"

    $dataSet1 | New-PSChart @defaultsPie -NoChartBorder -Title "No Chart Border"
    $dataSet1 | New-PSChart -Width ($width*2) -Height ($height*2) -Type Pie -As ImgTag -XProperty Label -YProperty Count -Title "Double Width/Height"

) | Out-File $env:TEMP\PSChartTests.html

& "$env:TEMP\PSChartTests.html"


return
# UI Tests

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