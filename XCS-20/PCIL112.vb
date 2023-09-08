Module PCIL112
    Declare Function B_pll12_initial Lib "PCI_L112.dll" Alias "_l112_open" (existCards As Integer) As Integer
    Declare Function B_pl112_close Lib "PCI_L112.dll" Alias "_l112_close" () As Integer
End Module
