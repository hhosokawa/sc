Dim i As Double
Dim j As Double
Dim k As Integer
Dim length As Double
Dim low_bound As Double
Dim upper_bound As Double
Dim num_of_digits As Integer

Dim region As String
Dim client As Currency
Dim corp As Currency
Dim essn As Currency
Dim microsoft As Currency
Dim services As Currency
Dim rebate As Currency
Dim mdf As Currency
Dim comp As Currency
Dim sga As Currency
Dim coop As Currency

Dim client_range As Currency
Dim corp_range As Currency
Dim essn_range As Currency
Dim microsoft_range As Currency
Dim services_range As Currency
Dim rebate_range As Currency
Dim mdf_range As Currency
Dim comp_range As Currency
Dim sga_range As Currency
Dim coop_range As Currency
Dim y_origin As Integer
Dim x_origin As Integer

Dim ebita As Currency
Dim freq(0 To 39) As Integer
Dim bins_array(0 To 39) As Double
Dim ebitda_array(0 To 49999) As Double
Dim y_x_origins(0 To 6) As Integer

' ############### main - cmd_refresh_Click() ###############
Private Sub cmd_refresh_Click()
    Application.ScreenUpdating = False
    
    ' Iterate through all Regions (y, x)
    For k = 0 To UBound(y_x_origins)
        x_origin = 6
        y_origin = 5 + (22 * k)
        calc_x_origin = 1 + (2 * k)
   
        region = Cells(y_origin - 1, 1).Value
        client = Cells(y_origin, x_origin).Value
        corp = Cells(y_origin + 1, x_origin).Value
        essn = Cells(y_origin + 2, x_origin).Value
        microsoft = Cells(y_origin + 3, x_origin).Value
        services = Cells(y_origin + 4, x_origin).Value
        rebate = Cells(y_origin + 6, x_origin).Value
        mdf = Cells(y_origin + 7, x_origin).Value
        comp = Cells(y_origin + 9, x_origin).Value
        sga = Cells(y_origin + 10, x_origin).Value
        coop = Cells(y_origin + 11, x_origin).Value
          
        client_range = Cells(y_origin, x_origin + 2).Value - Cells(y_origin, x_origin).Value
        corp_range = Cells(y_origin + 1, x_origin + 2).Value - Cells(y_origin + 1, x_origin).Value
        essn_range = Cells(y_origin + 2, x_origin + 2).Value - Cells(y_origin + 2, x_origin).Value
        microsoft_range = Cells(y_origin + 3, x_origin + 2).Value - Cells(y_origin + 3, x_origin).Value
        services_range = Cells(y_origin + 4, x_origin + 2).Value - Cells(y_origin + 4, x_origin).Value
        rebate_range = Cells(y_origin + 6, x_origin + 2).Value - Cells(y_origin + 6, x_origin).Value
        mdf_range = Cells(y_origin + 7, x_origin + 2).Value - Cells(y_origin + 7, x_origin).Value
        comp_range = Cells(y_origin + 9, x_origin + 2).Value - Cells(y_origin + 9, x_origin).Value
        sga_range = Cells(y_origin + 10, x_origin + 2).Value - Cells(y_origin + 10, x_origin).Value
        coop_range = Cells(y_origin + 11, x_origin + 2).Value - Cells(y_origin + 11, x_origin).Value
    
        ' ebitda = (fm + rebates + mdf) - opex
        For i = 0 To UBound(ebitda_array)
            ebitda = ((Rnd * client_range + client + _
                       Rnd * corp_range + corp + _
                       Rnd * essn_range + essn + _
                       Rnd * microsoft_range + microsoft + _
                       Rnd * services_range + services + _
                       Rnd * rebate_range + rebate + _
                       Rnd * mdf_range + mdf) - _
                      (Rnd * comp_range + comp + _
                       Rnd * sga_range + sga + _
                       Rnd * coop_range + coop))
            ebitda_array(i) = ebitda
        Next i
    
        ' statistics - ebitda_array
        ebitda_count = WorksheetFunction.Count(ebitda_array)
        ebitda_max = WorksheetFunction.Max(ebitda_array)
        ebitda_min = WorksheetFunction.Min(ebitda_array)
        mean = WorksheetFunction.Average(ebitda_array)
        std_dev = WorksheetFunction.StDev(ebitda_array)
        std_error = std_dev / Sqr(ebitda_count)
        
        Cells(y_origin, x_origin + 6).Value = ebitda_count
        Cells(y_origin + 1, x_origin + 6).Value = mean
        Cells(y_origin + 2, x_origin + 6).Value = WorksheetFunction.Median(ebitda_array)
        Cells(y_origin + 3, x_origin + 6).Value = ebitda_max
        Cells(y_origin + 4, x_origin + 6).Value = ebitda_min
        Cells(y_origin + 5, x_origin + 6).Value = std_dev
        Cells(y_origin + 6, x_origin + 6).Value = std_error
        Cells(y_origin + 7, x_origin + 6).Value = mean + (2.575 * std_error)
        Cells(y_origin + 8, x_origin + 6).Value = mean - (2.575 * std_error)
        Cells(y_origin + 9, x_origin + 6).Value = WorksheetFunction.Quartile(ebitda_array, 3)
        Cells(y_origin + 10, x_origin + 6).Value = WorksheetFunction.Quartile(ebitda_array, 1)
        
        ' scatter plot
        '   - Linear Interpolation
        num_of_digits = -(Len(Round(ebitda_min)) - 2)
        lower_bound = WorksheetFunction.RoundUp(ebitda_min, num_of_digits)
        upper_bound = WorksheetFunction.RoundUp(ebitda_max, num_of_digits)
        length = (upper_bound - lower_bound) / 40
        For i = LBound(bins_array) To UBound(bins_array)
            freq(i) = 0
            bins_array(i) = lower_bound + (length * i)
        Next i
        
        '   - Counting # of occurences for bins
        For i = LBound(ebitda_array) To UBound(ebitda_array)
            If ebitda_array(i) <= bins_array(LBound(bins_array)) Then
                freq(0) = freq(0) + 1
            End If
            If ebitda_array(i) >= bins_array(UBound(bins_array)) Then
                freq(39) = freq(39) + 1
            End If
            For j = 1 To (UBound(freq) - 1)
                If ebitda_array(i) > bins_array(j - 1) And ebitda_array(i) <= bins_array(j) Then
                    freq(j) = freq(j) + 1
                End If
            Next j
        Next i
      
        '   - bins / frequency output -> calc tab
        Sheets("calc").Cells(1, calc_x_origin).Value = region & " bins"
        Sheets("calc").Cells(1, calc_x_origin + 1).Value = region & " frequency"
        For i = 0 To 39
            Sheets("calc").Cells(i + 2, calc_x_origin).Value = bins_array(i)
            Sheets("calc").Cells(i + 2, calc_x_origin + 1).Value = freq(i) / (UBound(ebitda_array) + 1)
        Next i
        
    Next k
    Application.ScreenUpdating = True
End Sub
