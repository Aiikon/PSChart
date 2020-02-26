[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms.DataVisualization')

Function New-PSChart
{
    [CmdletBinding(PositionalBinding=$false)]
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(Mandatory=$true, Position=0)]
            [ValidateSet('Point', 'FastPoint', 'Bubble', 'Line', 'Spline', 'StepLine', 'FastLine', 'Bar',
                'StackedBar', 'StackedBar100', 'Column', 'StackedColumn', 'StackedColumn100', 'Area',
                'SplineArea', 'StackedArea', 'StackedArea100', 'Pie', 'Doughnut', 'Stock', 'Candlestick', 'Range',
                'SplineRange', 'RangeBar', 'RangeColumn', 'Radar', 'Polar', 'ErrorBar', 'BoxPlot', 'Renko',
                'ThreeLineBreak', 'Kagi', 'PointAndFigure', 'Funnel', 'Pyramid')]
            [string] $Type,
        [Parameter(Position=1)] [string] $XProperty,
        [Parameter(Position=2)] [string] $YProperty = 'Count',
        [Parameter(Position=3)] [string] $ZProperty,
        [Parameter(Mandatory=$true)] [ValidateSet('ImgTag', 'WinFormControl', 'WpfControl')] [string] $As,
        [Parameter()] [switch] $UngroupedInput,
        [Parameter()] [string] $GroupProperty = 'Group',
        [Parameter()] [string] $LabelProperty,
        [Parameter()] [switch] $NoChartBorder,
        [Parameter()] [ValidateSet('Left', 'Top', 'Right', 'Bottom', 'None')] [string] $LegendPosition,
        [Parameter()] [int] $Width,
        [Parameter()] [int] $Height,
        [Parameter()] [string] $Title,
        [Parameter()] [int] $XAxisInterval = 1,
        [Parameter()] [int] $YAxisInterval,
        [Parameter()] [System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles] $XAxisAutoFitStyle
        
    )
    Begin
    {
        $inputObjectList = New-Object System.Collections.Generic.List[object]
    }
    Process
    {
        if (!$InputObject) { return }
        $inputObjectList.Add($InputObject)
    }
    End
    {
        trap { $PSCmdlet.ThrowTerminatingError($_) }

        if ($UngroupedInput)
        {
            if ($PSBoundParameters['YProperty']) { throw '-YProperty cannot be provided with -UngroupedInput' }
            if (!$XProperty) { $XProperty = $inputObjectList[0].PSObject.Properties.Name | Select-Object -First 1 }
            $keys = @($XProperty)
            if ($ZProperty) { $keys += $ZProperty }
            $dictionary = $inputObjectList | ConvertTo-Dictionary $keys
            $inputObjectList = foreach ($group in $dictionary.Values)
            {
                $result = [ordered]@{}
                $result.$XProperty = $group[0].$XProperty
                if ($ZProperty) { $result.$ZProperty = $group[0].$ZProperty }
                $result.$YProperty = $group.Count
                $result.$GroupProperty = $group
                [pscustomobject]$result
            }
        }

        $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
        $chart.ChartAreas.Add((New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea))
        $chart.Palette = [System.Windows.Forms.DataVisualization.Charting.ChartColorPalette]::None
        $chart.PaletteCustomColors = Get-PSChartColors
        $chart.BackColor = 'White'

        if($YAxisInterval) { $chart.ChartAreas[0].AxisY.MajorGrid.Interval = $YAxisInterval }
        $chart.ChartAreas[0].AxisX.LabelAutoFitStyle = [System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles]::StaggeredLabels
        $chart.ChartAreas[0].AxisX.Interval = $XAxisInterval
        $chart.ChartAreas[0].AxisX.MajorGrid.Enabled = $false
        if ($XAxisAutoFitStyle) { $chart.ChartAreas[0].AxisX.LabelAutoFitStyle = $XAxisAutoFitStyle }

        $parameterDict = @{} + $PSBoundParameters
        $splatDict = @{}
        $splatDict.ShowInLegend = $LegendPosition -or $ZProperty -or $ShowValueInLegend
        if (!$Script:NewPSChartDataSeriesParameters)
        {
            $Script:NewPSChartDataSeriesParameters = (Get-Command New-PSChartDataSeries).Parameters.Keys
        }
        foreach ($parameter in $Script:NewPSChartDataSeriesParameters)
        {
            if($parameterDict.Contains($parameter)) { $splatDict[$parameter] = $parameterDict[$parameter] }
        }
        $splatDict.Remove('InputObject')

        if ($ZProperty)
        {
            $xPropertyList = $inputObjectList |
                ConvertTo-Dictionary -Keys $XProperty -Ordered |
                ForEach-Object Keys
            $zPropertyDict = $inputObjectList |
                ConvertTo-Dictionary -Keys $ZProperty -Ordered
            foreach ($zPropertyValue in $zPropertyDict.Keys)
            {
                $group = $zPropertyDict[$zPropertyValue]
                $xPropertyDict = $group | ConvertTo-Dictionary -Keys $XProperty -Ordered
                $objectList = foreach ($xPropertyValue in $xPropertyList)
                {
                    $group = $xPropertyDict[$xPropertyValue]
                    if (!$group)
                    {
                        $group = 1 | Select-Object $XProperty, $YProperty, $ZProperty
                        $group.$XProperty = $xPropertyValue
                        $group.$YProperty = 0
                        $group.$ZProperty = $zPropertyValue
                    }
                    $group
                }
                $series = $objectList | New-PSChartDataSeries @splatDict -LegendText $zPropertyValue
                $chart.Series.Add($series)
            }
                
        }
        else
        {
            $series = $inputObjectList | New-PSChartDataSeries @splatDict
            $chart.Series.Add($series)
        }

        $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
        $legend.Name = 'Legend0'

        if ($LegendPosition -in 'Left', 'Right', 'Top', 'Bottom')
        {
            $legend.Docking = $LegendPosition
        }

        if ($LegendPosition -ne 'None') { $chart.Legends.Add($legend) }

        if ($Width) { $chart.Width = $Width }
        if ($Height) { $chart.Height = $Height }

        if (!$NoChartBorder)
        {
            $chart.BorderWidth = 1
            $chart.BorderColor = 'Black'
            $chart.BorderDashStyle = 'Solid'
        }

        if ($Title)
        {
            $chartTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
            $chartTitle.Text = $Title
            $chartTitle.Font = 'Segoe UI,18pt'
            $chart.Titles.Add($chartTitle)
        }

        if ($As -eq 'ImgTag')
        {
            if (!$Width) { $chart.Width = [Math]::Ceiling([System.Windows.SystemParameters]::PrimaryScreenWidth * 0.40) }
            if (!$Height) { $chart.Height = [Math]::Ceiling([System.Windows.SystemParameters]::PrimaryScreenHeight * 0.40) }
            $memStream = New-Object System.IO.MemoryStream
            $chart.SaveImage($memStream, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)
            $memStream.Close()

            $bytes = $memStream.ToArray()
            $base64 = [Convert]::ToBase64String($bytes)
            "<img src='data:image/png;base64,$base64' />"
        }
        elseif ($As -eq 'WinFormControl')
        {
            $chart
        }
        else
        {
            [void][Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration')
            $formsHost = New-Object System.Windows.Forms.Integration.WindowsFormsHost
            $formsHost.Child = $chart
            $formsHost
        }
    }
}

Function Get-PSChartColors
{
    if ($Script:ChartColors) { return $Script:ChartColors }
    $chartColorValues = [Array]::CreateInstance([int[]], 15)
    $chartColorValues[0] = 67, 134, 216
    $chartColorValues[1] = 255, 154, 46
    $chartColorValues[2] = 219, 68, 63
    $chartColorValues[3] = 168, 212, 79
    $chartColorValues[4] = 133, 96, 179
    $chartColorValues[5] = 60, 191, 227
    $chartColorValues[6] = 175, 216, 248
    $chartColorValues[7] = 0, 142, 142
    $chartColorValues[8] = 139, 186, 0
    $chartColorValues[9] = 250, 189, 15
    $chartColorValues[10] = 250, 110, 70
    $chartColorValues[11] = 157, 8, 13
    $chartColorValues[12] = 161, 134, 190
    $chartColorValues[13] = 204, 102, 0
    $chartColorValues[14] = 253, 198, 137
    $Script:ChartColors = foreach ($color in $chartColorValues)
    {
        [System.Drawing.Color]::FromArgb($color[0], $color[1], $color[2])
    }
    $Script:ChartColors
}

Function New-PSChartDataSeries
{
    [CmdletBinding(PositionalBinding=$false)]
    Param
    (
        [Parameter()] [string] $LegendText,
        [Parameter()] $ShowInLegend = $true,
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(Mandatory=$true)]
            [ValidateSet('Point', 'FastPoint', 'Bubble', 'Line', 'Spline', 'StepLine', 'FastLine', 'Bar',
                'StackedBar', 'StackedBar100', 'Column', 'StackedColumn', 'StackedColumn100', 'Area',
                'SplineArea', 'StackedArea', 'StackedArea100', 'Pie', 'Doughnut', 'Stock', 'Candlestick', 'Range',
                'SplineRange', 'RangeBar', 'RangeColumn', 'Radar', 'Polar', 'ErrorBar', 'BoxPlot', 'Renko',
                'ThreeLineBreak', 'Kagi', 'PointAndFigure', 'Funnel', 'Pyramid')]
            [string] $Type,
        [Parameter()] [string] $XProperty,
        [Parameter()] [string] $YProperty = 'Count',
        [Parameter()] [string] $GroupProperty = 'Group',
        [Parameter()] [string] $LabelProperty
    )
    Begin
    {
        $inputObjectList = New-Object System.Collections.Generic.List[object]
    }
    Process
    {
        if (!$InputObject) { return }
        $inputObjectList.Add($InputObject)
    }
    End
    {
        $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
        $series.ChartType = $Type
        $series.LegendText = $LegendText
        if ($Type -eq 'Pie')
        {
            $series.SetCustomProperty('PieLabelStyle', 'Outside')
            $series.SetCustomProperty('PieLineColor', 'Black')
            if ($LabelProperty)
            {
                $series.Label = $InputObject.$LabelProperty
            }
            else
            {
                $series.Label = "$LegendText (#VALY)"
            }
        }

        if (!$XProperty)
        {
            $XProperty = $inputObjectList[0].PSObject.Properties |
                Where-Object Name -ne $YProperty |
                Select-Object -First 1 -ExpandProperty Name
        }

        foreach ($object in $inputObjectList)
        {
            $dataPoint = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
            $dataPoint.AxisLabel = $object.$XProperty
            $dataPoint.YValues = $object.$YProperty
            if (!$LegendText)
            {
                $dataPoint.LegendText = $dataPoint.AxisLabel
                if ($Type -eq 'Pie') { $dataPoint.Label = "$($dataPoint.AxisLabel) (#VALY)" }
            }
            if ($LabelProperty) { $dataPoint.Label = $object.$LabelProperty }
            $series.Points.Add($dataPoint)
        }

        $series
    }
}

Function ConvertTo-Dictionary
{
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(Mandatory=$true,Position=0)] [string[]] $Keys,
        [Parameter()] [string] $Value,
        [Parameter()] [string] $KeyJoin = '|',
        [Parameter()] [switch] $Ordered
    )
    Begin
    {
        if ($Ordered) { $dict = [ordered]@{} }
        else { $dict = @{} }
    }
    Process
    {
        $keyValue = $(foreach ($key in $Keys) { $InputObject.$key }) -join $KeyJoin
        if ($Value)
        {
            if (!$dict.Contains($keyValue))
            {
                Write-Warning "Dictionary already contains key '$keyValue'."
                return
            }
            $dict[$keyValue] = $InputObject.$Value
        }
        else
        {
            if (!$dict.Contains($keyValue))
            {
                $dict[$keyValue] = New-Object System.Collections.Generic.List[object]
            }
            $dict[$keyValue].Add($InputObject)
        }
    }
    End
    {
        $dict
    }
}