Imports System.Math
Public Class LIDM


    







    Function DW_headloss(ByVal rough As Single, ByVal diam As Single, ByVal discharge As Single, ByVal length As Single) As Single

        ' Water viscosity = 0.000891 kg/m.sec at 25°C
        ' water density = 1000 kg/m3
        ' vitesse = m/sec
        ' diam = mm
        'discharge : l/s
        'Re = density*vitesse*diam/viscosity

        Dim Re As Integer
        Dim lamda As Single
        Dim vitesse As Single
        discharge = discharge * 3.6 'l/s to m3 /h
        vitesse = (discharge / 3600) / (3.14 * ((diam / 1000) ^ 2) / 4)
        Re = (vitesse * diam * 1000000) / 891

        If Re < 2000 Then
            If Re = 0 Then
                lamda = 0
            Else
                lamda = 64 / Re
            End If

        Else


            Dim N As Integer
            Dim X0, XN, FX, df As Single

            'friction mm
            'diameter mm
            Dim NI As Integer = 50 ' maximum number of iteration
            Dim EP As Single = 0.0000001 'error
            Dim XI As Single = 0.001 'estimated value
            Dim A, B As Single

            A = 2.51 / Re
            B = rough / (3.7 * diam)


            N = 0
            X0 = XI

200:        N = N + 1

            FX = (X0 ^ (-0.5)) + 2 * Log10((A * (X0 ^ (-0.5))) + B)

            df = -0.5 * (X0 ^ (-1.5)) - ((A * (X0 ^ (-1.5))) / (2.3 * (A * (X0 ^ (-0.5)) + B)))


            XN = X0 - (FX / df)
            FX = (XN ^ (-0.5)) + 2 * Log10((A * (XN ^ (-0.5))) + B)

            If Abs(XN - X0) <= EP Then
                lamda = XN
                GoTo 120
            End If
            If N < NI Then
                X0 = XN
                GoTo 200
            End If
        End If

        ' lamda :friction factor (without dimension)
        'headloss: m
120:    Dim rho As Single = 1000 ' Water density Kg/m3
        'vitesse= velocity in m/sec
        'diam: pipe diameter in mm
        'length : pipe length in m


        DW_headloss = (((lamda * rho * (vitesse * vitesse) * length)) / (2 * (diam / 1000))) / 10000 'divided by 10-4 to transform pa to m

    End Function
End Class
