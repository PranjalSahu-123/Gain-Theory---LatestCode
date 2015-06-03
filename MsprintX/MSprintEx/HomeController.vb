Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
'Imports System.Web.Mvc
'Imports newapp1.Models
Imports System.Data
Imports System.Collections
Imports System.IO
Imports System.Data.OleDb
Imports System.Text
'Imports System.Collections.Generic


Namespace newapp1.Controllers
	Public Class HomeController
        'Inherits Controller

        'Public Function Index(RunID As [String], UserId As [String]) As ActionResult

        '	Return View()
        'End Function
        'Public Function About() As ActionResult
        '	Return View()

        'End Function

        'Public Function cal() As ActionResult

        '	Dim dt As New DataTable()
        '	Dim output As New DataTable()
        '	dt.Columns.Add("tcp", GetType(Decimal))
        '	dt.Columns.Add("val1", GetType(Decimal))
        '	dt.Columns.Add("val2", GetType(Decimal))
        '	dt.Columns.Add("val3", GetType(Decimal))
        '	dt.Columns.Add("val4", GetType(Decimal))
        '	dt.Columns.Add("val5", GetType(Decimal))
        '	dt.Columns.Add("val6", GetType(Decimal))
        '	dt.Columns.Add("val7", GetType(Decimal))
        '	dt.Columns.Add("val8", GetType(Decimal))
        '	dt.Columns.Add("val9", GetType(Decimal))
        '	dt.Columns.Add("val10", GetType(Decimal))
        '	dt.Columns.Add("val11", GetType(Decimal))
        '	dt.Columns.Add("val12", GetType(Decimal))
        '	dt.Columns.Add("val13", GetType(Decimal))
        '	dt.Columns.Add("val14", GetType(Decimal))
        '	dt.Columns.Add("val15", GetType(Decimal))
        '	dt.Columns.Add("val16", GetType(Decimal))
        '	dt.Columns.Add("val17", GetType(Decimal))
        '	dt.Columns.Add("val18", GetType(Decimal))
        '	dt.Columns.Add("val19", GetType(Decimal))
        '	dt.Columns.Add("val20", GetType(Decimal))
        '	Dim dr As DataRow = dt.NewRow()
        '	dr("tcp") = "503.8300"
        '	dr("val1") = "70.45"
        '	dr("val2") = "59.56"
        '	dr("val3") = "50.29"
        '	dr("val4") = "42.85"
        '	dr("val5") = "35.79"
        '	dr("val6") = "29.05"
        '	dr("val7") = "24.22"
        '	dr("val8") = "21.39"
        '	dr("val9") = "18.3"
        '	dr("val10") = "16.09"
        '	dr("val11") = "12.18"
        '	dr("val12") = "9.68"
        '	dr("val13") = "6.86"
        '	dr("val14") = "5.88"
        '	dr("val15") = "4.36"
        '	dr("val16") = "3.36"
        '	dr("val17") = "2.07"
        '	dr("val18") = "2.07"
        '	dr("val19") = "1.78"
        '	dr("val20") = "0"
        '	dt.Rows.Add(dr)
        '	dt.AcceptChanges()
        '	Dim dr1 As DataRow = dt.NewRow()
        '	dr1("tcp") = "143.47"
        '	dr1("val1") = "41.43"
        '	dr1("val2") = "25.77"
        '	dr1("val3") = "14.92"
        '	dr1("val4") = "10.1"
        '	dr1("val5") = "8.73"
        '	dr1("val6") = "6.45"
        '	dr1("val7") = "5.67"
        '	dr1("val8") = "3.41"
        '	dr1("val9") = "3.11"
        '	dr1("val10") = "2.69"
        '	dr1("val11") = "1.39"
        '	dr1("val12") = "0.55"
        '	dr1("val13") = "0.55"
        '	dr1("val14") = "0.55"
        '	dr1("val15") = "0.55"
        '	dr1("val16") = "0.55"
        '	dr1("val17") = "0.29"
        '	dr1("val18") = "0.29"
        '	dr1("val19") = "0.29"
        '	dr1("val20") = "0.29"
        '	dt.Rows.Add(dr1)
        '	dt.AcceptChanges()
        '	Dim dr2 As DataRow = dt.NewRow()
        '	dr2("tcp") = "26.11"
        '	dr2("val1") = "12.65"
        '	dr2("val2") = "5.26"
        '	dr2("val3") = "2.84"
        '	dr2("val4") = "1.46"
        '	dr2("val5") = "0.29"
        '	dr2("val6") = "0.29"
        '	dr2("val7") = "0.29"
        '	dr2("val8") = "0.29"
        '	dr2("val9") = "0.29"
        '	dr2("val10") = "0"
        '	dr2("val11") = "0"
        '	dr2("val12") = "0"
        '	dr2("val13") = "0"
        '	dr2("val14") = "0"
        '	dr2("val15") = "0"
        '	dr2("val16") = "0"
        '	dr2("val17") = "0"
        '	dr2("val18") = "0"
        '	dr2("val19") = "0"
        '	dr2("val20") = "0"
        '	dt.Rows.Add(dr2)
        '	dt.AcceptChanges()
        '	Dim dr3 As DataRow = dt.NewRow()
        '	dr3("tcp") = "480.26"
        '	dr3("val1") = "66.65"
        '	dr3("val2") = "55.74"
        '	dr3("val3") = "46.5"
        '	dr3("val4") = "40.25"
        '	dr3("val5") = "33.99"
        '	dr3("val6") = "27.6"
        '	dr3("val7") = "23.12"
        '	dr3("val8") = "20.46"
        '	dr3("val9") = "17.86"
        '	dr3("val10") = "14.69"
        '	dr3("val11") = "11.52"
        '	dr3("val12") = "9.5"
        '	dr3("val13") = "7.4"
        '	dr3("val14") = "5.98"
        '	dr3("val15") = "5.5"
        '	dr3("val16") = "4.26"
        '	dr3("val17") = "3.68"
        '	dr3("val18") = "2.86"
        '	dr3("val19") = "2.67"
        '	dr3("val20") = "2.29"
        '	dt.Rows.Add(dr3)
        '	dt.AcceptChanges()
        '	Dim dr4 As DataRow = dt.NewRow()
        '	dr4("tcp") = "181.39"
        '	dr4("val1") = "36.74"
        '	dr4("val2") = "28.23"
        '	dr4("val3") = "21.28"
        '	dr4("val4") = "16.86"
        '	dr4("val5") = "12.74"
        '	dr4("val6") = "9.93"
        '	dr4("val7") = "7.89"
        '	dr4("val8") = "5.85"
        '	dr4("val9") = "4.33"
        '	dr4("val10") = "2.96"
        '	dr4("val11") = "2.09"
        '	dr4("val12") = "1.32"
        '	dr4("val13") = "1.02"
        '	dr4("val14") = "0.78"
        '	dr4("val15") = "0.5"
        '	dr4("val16") = "0.22"
        '	dr4("val17") = "0.22"
        '	dr4("val18") = "0"
        '	dr4("val19") = "0"
        '	dr4("val20") = "0"
        '	dt.Rows.Add(dr4)
        '	dt.AcceptChanges()
        '	Dim dr5 As DataRow = dt.NewRow()
        '	dr5("tcp") = "33.52"
        '	dr5("val1") = "12.15"
        '	dr5("val2") = "8.04"
        '	dr5("val3") = "3.92"
        '	dr5("val4") = "2.9"
        '	dr5("val5") = "1.42"
        '	dr5("val6") = "0.93"
        '	dr5("val7") = "0.4"
        '	dr5("val8") = "0.27"
        '	dr5("val9") = "0.2"
        '	dr5("val10") = "0.2"
        '	dr5("val11") = "0.09"
        '	dr5("val12") = "0.09"
        '	dr5("val13") = "0"
        '	dr5("val14") = "0"
        '	dr5("val15") = "0"
        '	dr5("val16") = "0"
        '	dr5("val17") = "0"
        '	dr5("val18") = "0"
        '	dr5("val19") = "0"
        '	dr5("val20") = "0"
        '	dt.Rows.Add(dr5)
        '	dt.AcceptChanges()

        '	output = nbdmain(dt)
        '	Dim dtfinal As DataTable = output.Copy()

        '	Return View()


        'End Function


		Public Function nbdmain(dt As DataTable) As DataTable
			Dim dtfinal As DataTable = dt.Clone()

			'DataTable dsTak = new DataTable();


			Dim flArray As Decimal() = New Decimal(dt.Columns.Count - 2) {}

			Dim outputArray As Decimal() = New Decimal(dt.Columns.Count - 2) {}

			For Each drnew As DataRow In dt.Rows
				Dim tp As Decimal = CDec(drnew("tcp"))
				For i As Integer = 1 To dt.Columns.Count - 1
					flArray(i - 1) = CDec(drnew(i))
				Next
				outputArray = nbd(tp, flArray)

				Dim drout As DataRow = dtfinal.NewRow()
				For x As Integer = 0 To outputArray.Length - 1
					drout(x) = outputArray(x)
				Next
				dtfinal.Rows.Add(drout)
				dtfinal.AcceptChanges()
			Next

			Return dtfinal
		End Function
		Public Function nbd(tpval As Decimal, fcttemp As Decimal()) As Decimal()
			Dim [error] As Decimal() = New Decimal() {CDec(-1)}
			Try
				Dim tp As Decimal = tpval

				Dim fct As Decimal() = New Decimal(fcttemp.Length) {}
				fct(0) = 0

				If fcttemp.Contains(CDec(100)) Then
					Return [error]
				Else
					For T As Integer = 1 To fcttemp.Length - 1
						fct(T) = fcttemp(T - 1) / 100
					Next
					fct(0) = 1 - fct(1)


					Dim fc As Decimal() = New Decimal(fct.Length - 1) {}
					fc(0) = fct(0)

					For T As Integer = 1 To fct.Length - 2
						fc(T) = fct(T) - fct(T + 1)
					Next
					fc(fct.Length - 1) = fct(fct.Length - 1)

					Dim tc As Decimal = CDec(0)

					For T As Integer = 1 To fc.Length - 1
						tc = tc + T * fc(T)
					Next

					tc = tc * 100

					Dim c As Decimal = tc / (100 * CDec(Math.Log(CDbl(fc(0)))))

					If c >= -1 Then
						Return [error]
					Else

						Dim a As Decimal = -2 * (1 + c)
						' decimal a = (decimal)5.0529863205884977;
						Dim nbd_a As Decimal = nbdparams(a, c)

						Dim k As Decimal = tc / (100 * nbd_a)

						Dim ap As Decimal = nbd_a * tp / tc

						Dim pc As Decimal() = New Decimal(fc.Length - 1) {}
						Dim pp As Decimal() = New Decimal(fc.Length - 1) {}
						Dim fp As Decimal() = New Decimal(fc.Length - 1) {}

						pc(0) = fc(0)
						pp(0) = CDec(Math.Pow(CDbl(1 / CDbl(1 + ap)), CDbl(k)))
						fp(0) = pp(0)

						Dim mysum As Decimal = fp(0)

						Dim anew As Decimal = nbd_a / (1 + nbd_a)
						Dim apnew As Decimal = ap / (1 + ap)

						For T As Integer = 1 To fp.Length - 2
							Dim x As Decimal = (k + T - 1) / T
							pc(T) = x * anew * pc(T - 1)
							pp(T) = x * apnew * pp(T - 1)
							fp(T) = pp(T) + fc(T) - pc(T)
							mysum = mysum + fp(T)
						Next
						fp(fp.Length - 1) = 1 - mysum

						Dim fpf As Decimal() = New Decimal(fp.Length - 1) {}

						fpf(fp.Length - 1) = fp(fp.Length - 1)

						For S As Integer = fp.Length - 2 To 1 Step -1
							fpf(S) = fp(S) + fpf(S + 1)
						Next
						fpf(0) = 1 - fpf(1)

						For i As Integer = 0 To fpf.Length - 1
							fpf(i) = CDec(Math.Round(CDec(fpf(i) * 100)))
						Next

						Return fpf.ToArray()
					End If
				End If
			Catch e As Exception
				Return [error]
			End Try

		End Function
		Public Function nbdparams(a As Decimal, c As Decimal) As Decimal
			Dim b As Decimal = a
			Dim a1 As Decimal = CDec(Math.Log(CDbl(1 + a)))
			Dim atmp As Decimal = c * (a - (1 + a) * a1) / (1 + a + c)
			If Math.Abs(CDbl(b - atmp)) < 0.001 Then
				Return CDec(atmp)
			Else
				Return nbdparams(CDec(atmp), c)
			End If
		End Function
	End Class
End Namespace
