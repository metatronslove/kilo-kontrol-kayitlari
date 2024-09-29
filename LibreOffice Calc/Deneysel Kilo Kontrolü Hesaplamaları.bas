'Kilo Takibi için yeni veri satırına isim Girerek başladığımda "Kayıt Giriş" tarihini doldursun
've değişme potansiyeli az olan hücrelere bir önceki girilen değerleri geçsin.
'SheetChange() fonksiyonunu çalıştırmak için Sayfa sekmesi sağ tıklanır ve "Sayfa Olayları..."
'"içerik değişti" seçeneğine makro ataması yapılır.
Sub SheetChange(oEvent)
	Dim oSheet As Object
	Dim birth As Date
	oSheet = ThisComponent.CurrentController.getActiveSheet()
	If oEvent.CellAddress.Row >= 1 Then
		oSheet.getCellByPosition(0, oEvent.CellAddress.Row).setString(DateSerial(year(now()), month(now()), day(now())))
		If oEvent.CellAddress.Column = 1 Then
			Rows = oEvent.CellAddress.Row - 1
			Do Until oSheet.getCellByPosition(1, Rows).getString() = oSheet.getCellByPosition(1, oEvent.CellAddress.Row).getString() Or Rows < 2
				Rows = Rows - 1
			Loop
			If oSheet.getCellByPosition(1, Rows).getString() = oSheet.getCellByPosition(1, oEvent.CellAddress.Row).getString() Then names = "ok"
			If Rows >= 1 And names = "ok" Then
				birth = oSheet.getCellByPosition(2, Rows).getString()
				oSheet.getCellByPosition(2, oEvent.CellAddress.Row).setString(DateSerial(year(birth), month(birth), day(birth)))
				oSheet.getCellByPosition(3, oEvent.CellAddress.Row).setValue(oSheet.getCellByPosition(3, Rows).getValue())
				oSheet.getCellByPosition(5, oEvent.CellAddress.Row).setString(oSheet.getCellByPosition(5, Rows).getString())
				oSheet.getCellByPosition(9, oEvent.CellAddress.Row).setString(oSheet.getCellByPosition(9, Rows).getString())
			Else
			EndIf
		Else
		EndIf
		If oEvent.CellAddress.Column = 2 Then oSheet.getCellByPosition(6, oEvent.CellAddress.Row).setFormula("=INT((A" & oEvent.CellAddress.Row + 1 & "-C" & oEvent.CellAddress.Row + 1 & ")/365)")
		If oEvent.CellAddress.Column = 3 Then oSheet.getCellByPosition(10, oEvent.CellAddress.Row).setFormula("=OKG(E" & oEvent.CellAddress.Row + 1 & ";D" & oEvent.CellAddress.Row + 1 & ";IF(F" & oEvent.CellAddress.Row + 1 & "=" & Chr(34) & "KADIN" & Chr(34) & ";0;1);J" & oEvent.CellAddress.Row + 1 & ")")
		If oEvent.CellAddress.Column = 4 Then
			oSheet.getCellByPosition(7, oEvent.CellAddress.Row).setFormula("=VKE(E" & oEvent.CellAddress.Row + 1 & ";D" & oEvent.CellAddress.Row + 1 & ")")
			oSheet.getCellByPosition(8, oEvent.CellAddress.Row).setFormula("=TMB(E" & oEvent.CellAddress.Row + 1 & ";D" & oEvent.CellAddress.Row + 1 & ";G" & oEvent.CellAddress.Row + 1 & ";IF(F" & oEvent.CellAddress.Row + 1 & "=" & Chr(34) & "KADIN" & Chr(34) & ";0;1))")
		Else
		EndIf
	Else
	EndIf
End Sub
'Vücut Kitle Endeksi Fonksiyonu
'VKE(KİLO, BOY)
'Sonuç birimi kg/m²
'Kaynak https://www.acibadem.com.tr/ilgi-alani/vucut-kitle-indeksi-hesaplama/
'Tarih 28 Eylül 2024 Cumartesi
Function vke(ByVal kg As Double, Optional cm As Double) As Double
	vke = kg/(cm*cm/10000)
End Function
'Temel Metabolizma Hızı Fonksiyonu (Harris-Benedict formülü ile TMB)
'TMB(KİLO, BOY, YAŞ, SEX)
'Sonuç birimi cal/gün
'Kaynak https://calcuonline.com/hesaplamak/bazal-metabolizma-hesaplama/
'Tarih 28 Eylül 2024 Cumartesi
Function tmb(ByVal kg As Double, Optional cm As Double, Optional age As Long, Optional sex As Boolean) As Double
	Dim Root, Weight, Height, Alive As Double
	Select Case sex
	'Multipliers for female
	Case 0
		Root = 655.0955
		Weight = 9.5634
		Height = 1.8449
		Alive = 4.6756
	'Multipliers for male
	Case Else
		Root = 66.473
		Weight = 13.7516
		Height = 5.0033
		Alive = 6.755
	End Select
	tmb = Root + (Weight * kg) + (Height * cm) - (Alive * age)
End Function
'Optimal Kilogram Fonksiyonu
'OKG(KİLO, BOY, SEX, YÖNTEM)
'Sonuç birimi kg
'Kaynak https://calcuonline.com/hesaplamak/ideal-kilo-hesaplama/
'Tarih 29 Eylül 2024 Pazar
Function okg(ByVal kg As Double, Optional cm As Double, Optional sex As Boolean, Optional method As Variant) As Variant
	Dim min, max As Double
	Dim feet As Double
	feet = cm * 0.032808
	Select Case method
	Case "Vücut Kitle Endeksi"
		min = 18.5*(cm*cm/10000)
		max = 25*(cm*cm/10000)
		okg = min & " kg - " & max & " kg"
	Case "Miller"
		Select Case sex
		Case 0
			okg = 53.1 + (12 * (feet - 5) * 1.36) & " kg"
		Case Else
			okg = 56.2 + (12 * (feet - 5) * 1.41) & " kg"
		End Select
	Case "Robinson"
		Select Case sex
		Case 0
			okg = 49 + (12 * (feet - 5) * 1.7) & " kg"
		Case Else
			okg = 52 + (12 * (feet - 5) * 1.9) & " kg"
		End Select
	Case "Hamwi"
		Select Case sex
		Case 0
			okg = 45.35923 + (12 * (feet - 5) * 2.267962) & " kg"
		Case Else
			okg = 48.08079 + (12 * (feet - 5) * 2.721554) & " kg"
		End Select
	Case "Devine"
		Select Case sex
		Case 0
			okg = 45.5 + (12 * (feet - 5) * 2.3) & " kg"
		Case Else
			okg = 50 + (12 * (feet - 5) * 2.3) & " kg"
		End Select
	Case "Lorentz"
		Select Case sex
		Case 0
			okg = cm - 100 - ((cm - 150) / 4) & " kg"
		Case Else
			okg = cm - 100 - ((cm - 150) / 2) & " kg"
		End Select
	Case Else
	End Select
End Function