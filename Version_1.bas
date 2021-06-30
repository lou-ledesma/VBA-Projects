Attribute VB_Name = "Module1"
Sub tossing_dice_lledesma_version()

Range("A1") = "Dice Roll Outcomes"
Range("B1") = "Frequency"
Range("C1") = "Distribution"

num_tosses = InputBox("How many tosses?")
Range("B13") = num_tosses
Range("M1") = "Starting Time"
Range("M2") = "Ending Time"
Range("M3") = "Elapsed Time"
Range("N1") = Now()

For counter = 2 To 12
Range("A" & counter) = counter
Next

For toss = 1 To num_tosses
roll_one = Application.RandBetween(1, 6)
roll_two = Application.RandBetween(1, 6)

total_roll = roll_one + roll_two

Select Case total_roll

Case 2: step_2 = step_2 + 1
Case 3: step_3 = step_3 + 1
Case 4: step_4 = step_4 + 1
Case 5: step_5 = step_5 + 1
Case 6: step_6 = step_6 + 1
Case 7: step_7 = step_7 + 1
Case 8: step_8 = step_8 + 1
Case 9: step_9 = step_9 + 1
Case 10: step_10 = step_10 + 1
Case 11: step_11 = step_11 + 1
Case 12: step_12 = step_12 + 1

End Select

Next

'If I take this chunk of code and position before the "Next" keyword, then it would incrementally increase'
Range("B2") = step_2
Range("B3") = step_3
Range("B4") = step_4
Range("B5") = step_5
Range("B6") = step_6
Range("B7") = step_7
Range("B8") = step_8
Range("B9") = step_9
Range("B10") = step_10
Range("B11") = step_11
Range("B12") = step_12
'Taking this chunk of code and positioning it after the "Next" keyword outputs the final numbers ONLY'


'Improve upon the way that it was initially done in class'
Range("A15") = "Scale"
granularity = InputBox("Enter a scale number:", "Scale of Bell")
Range("B15") = granularity

Range("C2") = Application.Rept("!", step_2 / granularity)
Range("C3") = Application.Rept("!", step_3 / granularity)
Range("C4") = Application.Rept("!", step_4 / granularity)
Range("C5") = Application.Rept("!", step_5 / granularity)
Range("C6") = Application.Rept("!", step_6 / granularity)
Range("C7") = Application.Rept("!", step_7 / granularity)
Range("C8") = Application.Rept("!", step_8 / granularity)
Range("C9") = Application.Rept("!", step_9 / granularity)
Range("C10") = Application.Rept("!", step_10 / granularity)
Range("C11") = Application.Rept("!", step_11 / granularity)
Range("C12") = Application.Rept("!", step_12 / granularity)


Range("N2") = Now()
Range("N3") = Range("N2") - Range("N1")

Call Time_Convert

End Sub
