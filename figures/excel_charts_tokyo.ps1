# Builds the report's three charts natively in Microsoft Excel from chart_data.json,
# exports each as PNG, and saves charts_tokyo.xlsx so the team can edit them in Excel.
$dir  = $PSScriptRoot
$data = Get-Content "$dir\chart_data.json" -Raw | ConvertFrom-Json
$NAVY = 3024675; $BLUE = 2949308; $LIGHTBLUE = 9211020; $GREY = 12236465; $LIGHTGREY = 14540253  # palette: charcoal, Tokyo red, greys

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()

  function New-Chart($ws, $type, $range, $png, $w, $h) {
    $shape = $ws.Shapes.AddChart2(-1, $type, 260, 10, $w, $h)
    $ch = $shape.Chart
    $ch.SetSourceData($ws.Range($range))
    $ch.HasTitle = $false                                  # the report caption says what the chart shows
    $ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = 28
    $ch.ChartArea.Format.Line.Visible = 0
    return $ch
  }

  # 1. Data sources by domain and timeliness (stacked bar)
  $ws = $wb.Worksheets.Item(1); $ws.Name = "Sources"
  $ws.Cells.Item(1, 1) = "Domain"
  for ($j = 0; $j -lt $data.sources.series.Count; $j++) { $ws.Cells.Item(1, $j + 2) = $data.sources.series[$j].name }
  for ($i = 0; $i -lt $data.sources.cats.Count; $i++) {
    $ws.Cells.Item($i + 2, 1) = $data.sources.cats[$i]
    for ($j = 0; $j -lt $data.sources.series.Count; $j++) { $v = $data.sources.series[$j].values[$i]; if ($v -gt 0) { $ws.Cells.Item($i + 2, $j + 2) = $v } }
  }
  # clustered columns (not stacked bars) so it reads differently from the initiatives chart; red = real-time
  $ch = New-Chart $ws 51 "A1:D$($data.sources.cats.Count + 1)" "" 1520 720
  $ch.Legend.Position = -4107
  $colors = @($BLUE, $NAVY, $LIGHTGREY)
  for ($j = 1; $j -le 3; $j++) { $ch.SeriesCollection($j).Format.Fill.ForeColor.RGB = $colors[$j - 1]; $ch.SeriesCollection($j).HasDataLabels = $true }
  $ch.ChartGroups(1).GapWidth = 70; $ch.ChartGroups(1).Overlap = 0
  $ch.Axes(2).HasTitle = $true; $ch.Axes(2).AxisTitle.Text = "Number of open data sources"
  $null = $ch.Export("$dir\fig_sources_excel_tokyo.png")

  # 2. How existing initiatives act, by domain (stacked bar)
  $ws = $wb.Worksheets.Add([Type]::Missing, $ws); $ws.Name = "Initiatives"
  $ws.Cells.Item(1, 1) = "Domain"
  for ($j = 0; $j -lt $data.actions.series.Count; $j++) { $ws.Cells.Item(1, $j + 2) = $data.actions.series[$j].name }
  for ($i = 0; $i -lt $data.actions.cats.Count; $i++) {
    $ws.Cells.Item($i + 2, 1) = $data.actions.cats[$i]
    for ($j = 0; $j -lt $data.actions.series.Count; $j++) {
      $v = $data.actions.series[$j].values[$i]
      if ($v -gt 0) { $ws.Cells.Item($i + 2, $j + 2) = $v }   # blanks, so empty segments get no "0" label
    }
  }
  $ch = New-Chart $ws 58 "A1:F$($data.actions.cats.Count + 1)" "" 1520 760
  $ch.Axes(1).ReversePlotOrder = $true; $ch.Axes(1).Crosses = 2
  $ch.Legend.Position = -4107
  $colors = @($BLUE, $NAVY, $LIGHTBLUE, $GREY, $LIGHTGREY)
  for ($j = 1; $j -le 5; $j++) { $ch.SeriesCollection($j).Format.Fill.ForeColor.RGB = $colors[$j - 1]; $ch.SeriesCollection($j).HasDataLabels = $true }
  for ($j = 1; $j -le 2; $j++) { $ch.SeriesCollection($j).DataLabels().Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215 }
  $ch.ChartGroups(1).GapWidth = 60
  $ch.Axes(2).HasTitle = $true; $ch.Axes(2).AxisTitle.Text = "Number of existing initiatives"
  $null = $ch.Export("$dir\fig_actions_excel_tokyo.png")

  # 3. Bus delays (column chart, 3+ minutes highlighted)
  $ws = $wb.Worksheets.Add([Type]::Missing, $ws); $ws.Name = "Bus delays"
  $ws.Cells.Item(1, 1) = "Minutes late"; $ws.Cells.Item(1, 2) = "Buses"
  for ($i = 0; $i -lt $data.delays.cats.Count; $i++) { $ws.Cells.Item($i + 2, 1) = "'" + $data.delays.cats[$i]; $ws.Cells.Item($i + 2, 2) = $data.delays.values[$i] }
  $ws.Cells.Item(8, 1) = "Source: ODPT Toei bus feed vs timetable, 30 Sep 2026 07:22 JST, $($data.delays.n) buses (lower bounds)"
  $ch = New-Chart $ws 51 "A1:B6" "" 1400 680
  $ch.HasLegend = $false
  $s = $ch.SeriesCollection(1); $s.HasDataLabels = $true
  for ($i = 1; $i -le 5; $i++) { $s.Points($i).Format.Fill.ForeColor.RGB = $(if ($i -ge 3) { $BLUE } else { $GREY }) }
  $ch.ChartGroups(1).GapWidth = 50
  $ch.Axes(1).HasTitle = $true; $ch.Axes(1).AxisTitle.Text = "Minutes late (lower bound)"
  $ch.Axes(2).HasTitle = $true; $ch.Axes(2).AxisTitle.Text = "Number of buses"
  $null = $ch.Export("$dir\fig_delays_excel_tokyo.png")

  $wb.SaveAs("$dir\charts_tokyo.xlsx", 51)
  "charts exported"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
