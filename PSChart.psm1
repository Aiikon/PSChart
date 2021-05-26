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
        [Parameter()] [string] $LabelProperty,
        [Parameter()] [switch] $NoChartBorder,
        [Parameter()] [ValidateSet('Left', 'Top', 'Right', 'Bottom', 'None')] [string] $LegendPosition,
        [Parameter()] [int] $Width,
        [Parameter()] [int] $Height,
        [Parameter()] [string] $Title,
        [Parameter()] [int] $XAxisInterval = 1,
        [Parameter()] [int] $YAxisInterval,
        [Parameter()] [System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles] $XAxisAutoFitStyle,
        [Parameter()] [object] $BackColor = 'White',
        [Parameter()] [object] $PieStartAngle = 0,
        [Parameter()] [object] $PieLineColor = 'Black',
        [Parameter()] [ValidateSet('Disabled', 'Inside', 'Outside', 'Ellipse')] [string] $PieLabelStyle = 'Outside',
        [Parameter()] [object[]] $ChartColors
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

        $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
        $chart.ChartAreas.Add((New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea))
        $chart.Palette = [System.Windows.Forms.DataVisualization.Charting.ChartColorPalette]::None
        $chart.PaletteCustomColors = Get-PSChartColors
        if ($ChartColors) { $chart.PaletteCustomColors = $ChartColors }
        $chart.BackColor = $BackColor

        if ($YAxisInterval) { $chart.ChartAreas[0].AxisY.MajorGrid.Interval = $YAxisInterval }
        $chart.ChartAreas[0].AxisX.LabelAutoFitStyle = [System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles]::StaggeredLabels
        $chart.ChartAreas[0].AxisX.Interval = $XAxisInterval
        $chart.ChartAreas[0].AxisX.MajorGrid.Enabled = $false
        $chart.ChartAreas[0].BackColor = $BackColor
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
            if ($parameterDict.Contains($parameter)) { $splatDict[$parameter] = $parameterDict[$parameter] }
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
        $legend.BackColor = $BackColor

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

Function New-PSSparklineChart
{
    [CmdletBinding(PositionalBinding=$false)]
    Param
    (
        [Parameter(ValueFromPipeline=$true)] [object] $InputObject,
        [Parameter(Mandatory=$true)] [string] $XProperty,
        [Parameter(Mandatory=$true)] [string] $YProperty,
        [Parameter()] [string] $ZProperty,
        [Parameter(Mandatory=$true)] [ValidateSet('ImgTag')] [string] $As,
        [Parameter()] [int] $Width,
        [Parameter()] [int] $Height,
        [Parameter()] [double] $YAxisMaximum,
        [Parameter()] [hashtable] $SeriesColors
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

        $chart = [System.Windows.Forms.DataVisualization.Charting.Chart]::new()
        $chart.ChartAreas.Add(([System.Windows.Forms.DataVisualization.Charting.ChartArea]::new()))
        $chart.Palette = [System.Windows.Forms.DataVisualization.Charting.ChartColorPalette]::None
        $chart.PaletteCustomColors = Get-PSChartColors
        $chart.BackColor = 'White'
        $chart.Width = $Width
        $chart.Height = $Height

        $chart.ChartAreas[0].AxisX.MajorGrid.Enabled = $false
        $chart.ChartAreas[0].AxisX.MajorTickMark.Enabled = $false
        $chart.ChartAreas[0].AxisX.LabelStyle.Enabled = $false
        $chart.ChartAreas[0].AxisX.LineWidth = 0
        $chart.ChartAreas[0].AxisY.MajorGrid.Enabled = $false
        $chart.ChartAreas[0].AxisY.MajorTickMark.Enabled = $false
        $chart.ChartAreas[0].AxisY.LabelStyle.Enabled = $false
        $chart.ChartAreas[0].AxisY.LineWidth = 0

        $chart.ChartAreas[0].AxisX.IntervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Seconds
        $chart.ChartAreas[0].AxisX.Interval = 1

        $chart.ChartAreas[0].Position.Auto = $false
        $chart.ChartAreas[0].Position.X = 0
        $chart.ChartAreas[0].Position.Width = 100
        $chart.ChartAreas[0].Position.Y = 0
        $chart.ChartAreas[0].Position.Height = 100

        if ($YAxisMaximum) { $chart.ChartAreas[0].AxisY.Maximum = $YAxisMaximum }

        if (!$ZProperty) { $ZProperty = [Guid]::NewGuid().ToString() }
        $seriesGroupList = $inputObjectList |
            ConvertTo-Dictionary -Keys $ZProperty -Ordered

        foreach ($seriesGroup in $seriesGroupList.GetEnumerator())
        {
            $series = [System.Windows.Forms.DataVisualization.Charting.Series]::new()
            $series.ChartType = 'FastLine'
            $series.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime
            if ($SeriesColors) { $series.Color = $SeriesColors[$seriesGroup.Key] }
            foreach ($dataPointObject in $seriesGroup.Value)
            {
                $dataPoint = [System.Windows.Forms.DataVisualization.Charting.DataPoint]::new()
                $dataPoint.XValue = ([datetime]$dataPointObject.$XProperty).ToOADate()
                $dataPoint.YValues = $dataPointObject.$YProperty
                $series.Points.Add($dataPoint)
            }
            $chart.Series.Add($series)
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
        else
        {
            throw "Unknown value for As."
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
        [Parameter()] [string] $LabelProperty,
        [Parameter()] [int] $PieStartAngle = 0,
        [Parameter()] [object] $PieLineColor = 'Black',
        [Parameter()] [ValidateSet('Disabled', 'Inside', 'Outside', 'Ellipse')] [string] $PieLabelStyle = 'Outside'
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
        if ($Type -in 'Pie', 'Doughnut')
        {
            $series.SetCustomProperty('PieLabelStyle', $PieLabelStyle)
            $series.SetCustomProperty('PieLineColor', $PieLineColor)
            $series.SetCustomProperty('PieStartAngle', $PieStartAngle)
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
                if ($Type -in 'Pie', 'Doughnut') { $dataPoint.Label = "$($dataPoint.AxisLabel) (#VALY)" }
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