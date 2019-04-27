# PSChart
Module wrapping System.Windows.Forms.DataVisualization.Charting charts.

# Samples

The New-PSChart cmdlet creates all charts. It always takes pipeline input. You can feed it grouped data from Group-Object and it'll automatically use Count as the YProperty and the first property (in this case Name) as the XProperty:

```powershell
Get-ChildItem C:\Windows |
    Group-Object Attributes |
    New-PSChart -Type Pie -As ImgTag |
    Out-File ~\Desktop\Chart1.html

& "~\Desktop\Chart1.html"
```

You can also feed it ungrouped input objects and provide XProperty and optionally a second ZProperty (how well a ZProperty works depends on the type of graph).
```powershell
Get-Service |
    New-PSChart -Type StackedColumn -XProperty StartType -ZProperty Status -UngroupedInput -As ImgTag |
    Out-File ~\Desktop\Chart2.html

& "~\Desktop\Chart2.html"
```
