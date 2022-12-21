Python 3.11.0 (v3.11.0:deaf509e8f, Oct 24 2022, 14:43:23) [Clang 13.0.0 (clang-1300.0.29.30)] on darwin
Type "help", "copyright", "credits" or "license()" for more information.
>>> Codingan :
... Private Sub Start_Click()
... Range("B12").Value = 0 '0
... delta_t = Range("B8").Value '0.1
... While Range("B12").Value < 9
... Range("B12").Value = Range("B12").Value + delta_t
... DoEvents
... Wend
