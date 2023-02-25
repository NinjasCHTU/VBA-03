Attribute VB_Name = "Work04_Rotation2"

Function W_Redbook2023MakeArr()
    W_Redbook2023MakeArr = Array("Alfa Romeo  ", " Aston Martin  ", " Audi  ", " Bentley  ", " BMW  ", " BYD  ", " Chery  ", " Chevrolet  ", " Chrysler  ", _
                                                        " Citroen  ", " Daewoo  ", " Daihatsu  ", " DFM  ", " DFSK  ", " Ferrari  ", " Fiat  ", " FOMM  ", " Ford  ", " Foton  ", " Haval  ", " Honda  ", " Hummer  ", _
                                                        " Hyundai  ", " Isuzu  ", " Jaguar  ", " Jeep  ", " Kia  ", " Lamborghini  ", " Land Rover  ", " Lexus  ", " Lotus  ", " Maserati  ", " Maxus  ", " Mazda  ", " McLaren  ", _
                                                        " Mercedes-Benz  ", " MG  ", " MINE  ", " Mini  ", " Mitsubishi  ", " Mitsuoka  ", " Neta  ", " Nissan  ", " Opel  ", " ORA  ", " Peugeot  ", " Polarsun  ", " Porsche  ", _
                                                        " Proton  ", " Renault  ", " Rolls-Royce  ", " Rover  ", " Saab  ", " Seat  ", " Skoda  ", " Smart  ", " Spyker  ", " Ssangyong  ", " Subaru  ", " Suzuki  ", " Tata  ", _
                                                        " Tesla  ", " Toyota  ", " TR  ", " Volkswagen  ", " Volt  ", " Volvo  ")

End Function

Function W_Redbook2023Dict()
    Dim makeModelDict As Object
    Set makeModelDict = CreateObject("Scripting.Dictionary")



End Function

Sub scratch01()
    Dim dict As Object
  Set dict = CreateObject("Scripting.Dictionary")

  dict.Add "Key1", Array(1, 2, 3)
  dict.Add "Key2", "Value2"
  dict.Add "Key3", "Value3"
  MsgBox (dict("Key2"))
End Sub
