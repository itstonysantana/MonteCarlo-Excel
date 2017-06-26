Attribute VB_Name = "Module3"
' This runs a Monte Carlo simulation, with no variance reduction methods implemented _

Public Sheet As String
Public s0 As Double
Public a As Double
Public div As Double
Public vol As Double
Public steps As Integer
Public sims As Integer
Public T As Double
Sub MonteCarloStrongTaylorInputs()

Sheet = InputBox("Which sheet will you use for the simulations?", "Monte Carlo Inputs")
s0 = InputBox("What is the initial stock price?", "Monte Carlo Inputs")
a = InputBox("What is the stock's expected rate of return?", "Monte Carlo Inputs")
div = InputBox("What is the stock's dividend yield rate?", "Monte Carlo Inputs")
vol = InputBox("What is the stock's volatility? (stepwise)", "Monte Carlo Inputs")
steps = InputBox("How many steps will the stock price take?", "Monte Carlo Inputs")
sims = InputBox("How many simulations do you want to run?", "Monte Carlo Inputs")
T = InputBox("What is the time period? (proportional to years)", "Monte Carlo Inputs")

Call MonteCarloStrongTaylor(Sheet, s0, a, div, vol, steps, sims, T)

End Sub
Sub MonteCarloStrongTaylor(Sheet As String, s0 As Double, a As Double, div As Double, vol As Double, steps As Integer, sims As Integer, T As Double)

Worksheets(Sheet).Activate
Worksheets(Sheet).Cells.ClearContents

drift = (a - div - (1 / 2) * (vol) ^ 2) * T
drift_i = a - div - (1 / 2) * (vol) ^ 2
vol_i = vol * Sqr(T)
dt = T / steps
vol_ii = 1 / 2 * vol * T ^ (-1 / 2)
vol_iii = -1 / 4 * vol * T ^ (-3 / 2)

ReDim SP(sims, steps)
For i = 0 To sims - 1
    SP(i, 0) = s0
Next i
ReDim WP(sims, steps)
For i = 0 To sims - 1
    For j = 0 To steps - 1
        WP(i, j) = (Sqr(-2 * WorksheetFunction.Ln(Rnd)) * Cos(2 * WorksheetFunction.Pi() * Rnd))
    Next j
Next i
For i = 0 To sims - 1
    For j = 1 To steps
        delta_WP = WP(i, j) - WP(i, j - 1)
        delta_Z = 1 / 2 * dt * (delta_WP + (Sqr(dt) * WP(i, j - 1)) / Sqr(3))
        SP1 = SP(i, j) + drift * dt + vol * Sqr(dt)
       SP(i, j) = SP(i, j - 1) + SP(i, j - 1) * ((drift * dt + vol_i * Sqr(dt) * delta_WP + 1 / 2 * vol_i * vol_ii * ((delta_WP) ^ 2 - dt)) + (1 / 2) * (drift * drift_i + 1 / 2 * (vol) ^ 2 * 0) * dt ^ 2 + (drift * vol_ii + 1 / 2 * vol_i ^ 2 * vol_iii) * (delta_WP * dt - delta_Z) + 1 / 2 * vol_i * (vol_i * vol_iii + vol_ii ^ 2) * (1 / 3 * delta_WP ^ 2 - dt) * delta_WP)
        If SP(i, j) < 0 Then ' Prevents stock prices from going negative
            SP(i, j) = 0.01
        End If
        Cells(i + 1, j) = SP(i, j)
    Next j
Next i

ActiveSheet.Shapes.AddChart2(227, xlLineMarkers).Select
ActiveChart.SetSourceData Source:=Range(Cells(1, 1), Cells(sims, steps))
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = "Price Paths"
Selection.Format.TextFrame2.TextRange.Characters.Text = "Price Paths"
With Selection.Format.TextFrame2.TextRange.Characters(1, 11).ParagraphFormat
  .TextDirection = msoTextDirectionLeftToRight
  .Alignment = msoAlignCenter
End With
With Selection.Format.TextFrame2.TextRange.Characters(1, 11).Font
  .BaselineOffset = 0
  .Bold = msoFalse
  .NameComplexScript = "+mn-cs"
  .NameFarEast = "+mn-ea"
  .Fill.Visible = msoTrue
  .Fill.ForeColor.RGB = RGB(89, 89, 89)
  .Fill.Transparency = 0
  .Fill.Solid
  .Size = 14
  .Italic = msoFalse
  .Kerning = 12
    .Name = "+mn-lt"
    .UnderlineStyle = msoNoUnderline
    .Spacing = 0
    .Strike = msoNoStrike
    End With
   End Sub






