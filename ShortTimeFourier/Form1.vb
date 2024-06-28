Imports System.IO
Imports OxyPlot
Imports OxyPlot.Axes
Imports OxyPlot.Series
Imports OxyPlot.WindowsForms

Public Class Form1
    Private Const sampleRate As Integer = 16000
    Private Const n_fft As Integer = 256
    Private Const hop_length As Integer = 64
    Private flowLayoutPanel As FlowLayoutPanel
    Private plotView As PlotView
    Private splitContainer As SplitContainer

    Private showNyquist As Boolean = True
    Private selectedFilePath As String = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Maximize the form to start in full screen mode
        Me.WindowState = FormWindowState.Maximized

        ' Initialize SplitContainer
        splitContainer = New SplitContainer With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Vertical,
            .SplitterDistance = 1,
            .Padding = New Padding(1)
        }

        ' Initialize FlowLayoutPanel and PlotView
        flowLayoutPanel = New FlowLayoutPanel With {
            .Dock = DockStyle.Left,
            .AutoScroll = True,
            .Width = 200
        }
        plotView = New PlotView With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Black
        }



        ' Add FlowLayoutPanel, Toggle Button, and Button1 to SplitContainer Panel1
        splitContainer.Panel1.Controls.Add(flowLayoutPanel)

        ' Add PlotView to SplitContainer Panel2
        splitContainer.Panel2.Controls.Add(plotView)

        ' Add SplitContainer to the Form
        Me.Controls.Add(splitContainer)

        ' Get the current directory
        Dim currentDirectory As String = My.Computer.FileSystem.CurrentDirectory
        ' Navigate to the "Vowel_recordings" folder
        Dim folderPath As String = Path.GetFullPath(Path.Combine(currentDirectory, "Vowel_recordings"))

        ' Check if the folder exists
        If Not Directory.Exists(folderPath) Then
            MessageBox.Show("Folder does not exist: " & folderPath)
            Exit Sub
        End If

        ' Get the text files in the folder
        Dim textFiles() As String = Directory.GetFiles(folderPath, "*.txt")

        ' Create a button for each text file
        For Each filePath As String In textFiles
            Dim fileName As String = Path.GetFileName(filePath)
            Dim button As New Button With {
                .Text = fileName,
                .Tag = filePath,
                .Width = 180,
                .Height = 30
            }
            AddHandler button.Click, AddressOf FileButton_Click
            flowLayoutPanel.Controls.Add(button)
        Next
    End Sub



    Private Sub FileButton_Click(sender As Object, e As EventArgs)
        ' Update the selected file and refresh the plot
        selectedFilePath = CType(sender, Button).Tag.ToString()
        RefreshPlot()
    End Sub

    Private Sub RefreshPlot()
        If String.IsNullOrEmpty(selectedFilePath) Then
            Return
        End If

        Dim filePath As String = selectedFilePath

        ' Read all lines from the file and convert to Double
        Dim lines() As String = File.ReadAllLines(filePath)
        Dim samples(lines.Length - 1) As Double
        For i As Integer = 0 To lines.Length - 1
            samples(i) = Convert.ToDouble(lines(i))
        Next

        ' Perform STFT
        Dim stftResults As List(Of Complex()) = STFT(samples, n_fft, hop_length)

        ' Compute magnitudes and scale from 0 to 255
        Dim maxMagnitude As Double = 0
        For Each frame In stftResults
            For Each c In frame
                maxMagnitude = Math.Max(maxMagnitude, c.Magnitude())
            Next
        Next

        Dim magnitudes(stftResults.Count - 1)() As Double
        For i As Integer = 0 To stftResults.Count - 1
            magnitudes(i) = New Double(stftResults(i).Length - 1) {}
            For j As Integer = 0 To stftResults(i).Length - 1
                magnitudes(i)(j) = 255 * (stftResults(i)(j).Magnitude() / maxMagnitude)
            Next
        Next

        ' Determine the range to display
        Dim displayLength As Integer = n_fft
        If showNyquist Then
            displayLength = n_fft \ 2
        End If

        Dim frequencies(displayLength - 1) As Double
        Dim nyquist As Double = sampleRate / 2.0

        ' Calculate the frequency resolution
        Dim freqResolution As Double = sampleRate / n_fft

        ' Calculate the frequency bins
        For i As Integer = 0 To displayLength - 1
            frequencies(i) = i * freqResolution
        Next

        ' Calculate total time spanned by all frames including initial offset
        Dim totalTime As Double = samples.Length / sampleRate
        '    Dim timePerHop As Double = hop_length / sampleRate

        ' Plot the spectrogram
        Dim plotModel As New PlotModel With {
        .Title = "STFT Magnitude - " & Path.GetFileName(filePath),
        .Background = OxyColors.Black,
        .TextColor = OxyColors.White,
        .PlotAreaBorderColor = OxyColors.White
    }

        ' Create and configure ColorAxis for greyscale
        Dim colorAxis As New LinearColorAxis With {
        .Position = AxisPosition.Right,
        .Palette = OxyPalettes.Gray(256),
        .LowColor = OxyColors.Black,
        .HighColor = OxyColors.White,
        .Minimum = 0,
        .Maximum = 255
    }
        plotModel.Axes.Add(colorAxis)

        ' Create and configure HeatMapSeries
        Dim heatMapSeries As New HeatMapSeries With {
        .X0 = 0,
        .X1 = totalTime / 2, ' Adjust X-axis span to total time
        .Y0 = 0,
        .Y1 = nyquist * 2, ' Adjust Y-axis span to Nyquist frequency
        .Data = New Double(magnitudes.Length - 1, displayLength - 1) {}
    }

        ' Fill the heatmap data
        For i As Integer = 0 To magnitudes.Length - 1
            For j As Integer = 0 To displayLength - 1

                heatMapSeries.Data(i, j) = magnitudes(i)(j)
            Next
        Next

        plotModel.Series.Add(heatMapSeries)

        ' Configure X and Y axes
        Dim xAxis As New LinearAxis With {
        .Position = AxisPosition.Bottom,
        .Title = "Time (s)",
        .TextColor = OxyColors.White,
        .TitleColor = OxyColors.White,
        .TicklineColor = OxyColors.White
    }
        plotModel.Axes.Add(xAxis)

        Dim yAxis As New LinearAxis With {
        .Position = AxisPosition.Left,
        .Title = "Frequency (Hz)",
        .TextColor = OxyColors.White,
        .TitleColor = OxyColors.White,
        .TicklineColor = OxyColors.White
    }
        plotModel.Axes.Add(yAxis)

        plotView.Model = plotModel
    End Sub






    Private Function STFT(samples() As Double, n_fft As Integer, hop_length As Integer) As List(Of Complex())
        Dim stftResults As New List(Of Complex())()

        ' Create the Hann window
        Dim hannWindow() As Double = CreateHammWindow(n_fft)

        For i As Integer = 0 To samples.Length - n_fft Step hop_length
            Dim window(n_fft - 1) As Double
            Array.Copy(samples, i, window, 0, n_fft)

            ' Apply the Hann window
            For j As Integer = 0 To n_fft - 1
                window(j) *= hannWindow(j)
            Next

            stftResults.Add(FFT(window))
        Next

        Return stftResults
    End Function

    Private Function CreateHammWindow(size As Integer) As Double()
        Dim window(size - 1) As Double
        For i As Integer = 0 To size - 1
            window(i) = 0.54 - 0.46 * Math.Cos(2 * Math.PI * i / (size - 1))
        Next
        Return window
    End Function

    Private Function FFT(samples() As Double) As Complex()
        Dim n As Integer = samples.Length
        If n = 1 Then
            Return New Complex() {New Complex(samples(0), 0)}
        End If

        If (n And (n - 1)) <> 0 Then
            Throw New ArgumentException("The number of samples is not a power of 2")
        End If

        Dim evenSamples(n \ 2 - 1) As Double
        Dim oddSamples(n \ 2 - 1) As Double

        For i As Integer = 0 To n \ 2 - 1
            evenSamples(i) = samples(2 * i)
            oddSamples(i) = samples(2 * i + 1)
        Next

        Dim evenFFT() As Complex = FFT(evenSamples)
        Dim oddFFT() As Complex = FFT(oddSamples)

        Dim result(n - 1) As Complex
        For k As Integer = 0 To n \ 2 - 1
            Dim t As Complex = Complex.FromPolarCoordinates(1, -2 * Math.PI * k / n) * oddFFT(k)
            result(k) = evenFFT(k) + t
            result(k + n \ 2) = evenFFT(k) - t
        Next

        Return result
    End Function
End Class

Public Class Complex
    Public Property Real As Double
    Public Property Imaginary As Double

    Public Sub New(real As Double, imaginary As Double)
        Me.Real = real
        Me.Imaginary = imaginary
    End Sub

    Public Function Magnitude() As Double
        Return Math.Sqrt(Real * Real + Imaginary * Imaginary)
    End Function
    Public Function MagnitudeSq() As Double
        Return (Real * Real + Imaginary * Imaginary)
    End Function

    Public Shared Operator +(c1 As Complex, c2 As Complex) As Complex
        Return New Complex(c1.Real + c2.Real, c1.Imaginary + c2.Imaginary)
    End Operator

    Public Shared Operator -(c1 As Complex, c2 As Complex) As Complex
        Return New Complex(c1.Real - c2.Real, c1.Imaginary - c2.Imaginary)
    End Operator

    Public Shared Operator *(c1 As Complex, c2 As Complex) As Complex
        Return New Complex(c1.Real * c2.Real - c1.Imaginary * c2.Imaginary, c1.Real * c2.Imaginary + c1.Imaginary * c2.Real)
    End Operator

    Public Shared Function FromPolarCoordinates(magnitude As Double, phase As Double) As Complex
        Return New Complex(magnitude * Math.Cos(phase), magnitude * Math.Sin(phase))
    End Function

    Public Shared Function Exp(c As Complex) As Complex
        Dim expReal As Double = Math.Exp(c.Real)
        Return New Complex(expReal * Math.Cos(c.Imaginary), expReal * Math.Sin(c.Imaginary))
    End Function

    Public Shared ReadOnly Property Zero As Complex
        Get
            Return New Complex(0, 0)
        End Get
    End Property
End Class
