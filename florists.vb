Sub Main()
        Dim florist_code As String
        Dim acres_country_flowers As Single
        Dim acres_covered_flowers As Single

        Dim sum_florist_financial_support As Integer
        Dim sum_financial_support As Integer

        Dim sum_above_10k As Single

        sum_financial_support = 0
        sum_above_10k = 0

        Dim max_country_flowers As Single
        Dim min_country_flowers As Single
        Dim diff_max_min As Single

        max_country_flowers = 0
        min_country_flowers = 9999999

        Dim sum_acres_country_flowers As Single
        Dim sum_acres As Single

        sum_acres_country_flowers = 0
        sum_acres = 0

        Do
            florist_code = InputBox("Enter a flower grower code or 0 for exit")

            If florist_code <> 0 Then
                acres_country_flowers = InputBox("Give acres of country flowers")
                acres_covered_flowers = InputBox("Give acres of flowers under cover")

                sum_florist_financial_support = acres_country_flowers * 1250 + acres_covered_flowers * 2500

                sum_financial_support = sum_financial_support + sum_florist_financial_support

                If sum_florist_financial_support > 10000 Then
                    sum_above_10k = sum_above_10k + 1
                End If

                If acres_country_flowers > max_country_flowers Then
                    max_country_flowers = acres_country_flowers
                End If

                If acres_country_flowers < min_country_flowers Then
                    min_country_flowers = acres_country_flowers
                End If

                sum_acres_country_flowers = sum_acres_country_flowers + acres_country_flowers
                sum_acres = sum_acres + acres_country_flowers + acres_covered_flowers

            End If

        Loop Until florist_code = 0

        MsgBox("The total amount of financial support is " & sum_financial_support)

        MsgBox("The number of flower growers receiving aid over 10000 is " & sum_above_10k)

        diff_max_min = max_country_flowers - min_country_flowers
        If diff_max_min > 1000 Then
            MsgBox("OK")
        End If

        Dim percentage_acres_country As Single

        percentage_acres_country = (sum_acres_country_flowers * 100) / sum_acres
        MsgBox("The percentage of acres of country flowers in the total acres is " & percentage_acres_country)

    End Sub
