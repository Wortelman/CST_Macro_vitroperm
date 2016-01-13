'-----------------------------------------------------------------------------------------------------------------------------
' History Of Changes
'-----------------------------------------------------------------------------------------------------------------------------
' 13-jan-2016 Niek Moonen: Added Core material creation and option for 1, 2 or 3 phases.
' 11-Jan-2016 Niek Moonen: Adding beginning and end angle of the windings.
' 05-Feb-2014 ube/fwe: dialog changes: inner and outer radius more intuitive, repaired picture, symmetric terminals always on
' 02-Jul-2012 fhi: symmetrical terminals at outer side of the windings
' 27-Apr-2012 fhi: option for wire-radius= 0 (creates curves), correcting initial point for curve
' 29-Sep-2011 ube: include picture, allow multiple execution for multiple coils
' 28-Sep-2011 jwa: hide dialog
' 18-Jan-2011 fwe: Initial version
'-----------------------------------------------------------------------------------------------------------------------------
'#Language "WWB-COM"

Option Explicit

Sub Main
Dim cst_core_r As Double, cst_core_w As Double, cst_core_h As Double, cst_core_i As Integer
Dim cst_core_x As Double, cst_core_y As Double, cst_core_z As Double, cst_wire_r As Double
Dim cst_wire_N As Integer, cst_result As Integer, cst_symm_term As Integer
Dim scst_core_ri As String, scst_core_ra As String, scst_core_h As String, scst_core_i As String
Dim scst_core_x As String, scst_core_y As String, scst_core_z As String, scst_wire_r As String
Dim scst_wire_N As String, scst_symm_term As Integer, scst_core_ang As String
Dim cst_core_ang As Double, cst_core_off As Double, scst_core_off As String
Dim cst_lead As Double, scst_lead As String
Dim cst_kern As Integer, scst_kern As String
Dim cst_torrus_x As Double, cst_torrus_y As Double, cst_torrus_z As Double, cst_torrus_i As Integer
Dim cst_core_ra As Double, cst_core_ri As Double
Dim cst_phases_N As Double, phases_i As Integer





BeginHide

	Begin Dialog UserDialog 640,330,"Create 3D Toroidal Coil, rectangular core.  " ' %GRID:10,7,1,1
		Text 30,25,120,14,"Core inner radius",.Text1
		Text 30,55,130,14,"Core outer radius",.Text2
		Text 30,85,90,14,"Core height",.Text3
		Text 30,115,90,14,"Wire radius",.Text4
		Text 30,145,90,14,"Turns",.Text5
		Text 30,175,90,14,"Angle in rad",.Text6
		Text 30,205,90,14,"Angle Offset",.Text7
		Text 30,235,90,14,"Lead",.Text8
		Text 30,265,90,14,"Number of Phases",.Text9
		TextBox 170,20,90,21,.ri
		TextBox 170,50,90,21,.ra
		TextBox 170,80,90,21,.h
		TextBox 170,110,90,21,.wr
		TextBox 170,140,90,21,.n
		TextBox 170,170,90,21,.ang
		TextBox 170,200,90,21,.off
		TextBox 170,230,90,21,.ld
		TextBox 170,260,90,21,.ph
		OptionGroup .option_kern
			OptionButton 310,210,160,21,"Core Off",.option_kern_off
			OptionButton 310,231,190,14,"Core On",.option_kern_on

		Picture 290,7,340,168,GetInstallPath + "\Library\Macros\Construct\Coils\3D Toroidal Coil - rectangular core.bmp",0,.Picture1
		OKButton 30,292,90,21
		CancelButton 130,292,90,21
		'CheckBox 30,168,210,14,"Symmetric Terminals",.symm_term
	End Dialog
	Dim dlg As UserDialog
	dlg.ri = "17.725"
	dlg.ra = "26.75"
	dlg.h = "23.3"
	dlg.wr = "0.3"
	dlg.n = "14"
	dlg.ang = "pi*0.44"
	dlg.off = "0.5*pi"
	dlg.ld = "2"
	dlg.ph = "3"
	'dlg.symm_term = 1
	dlg.option_kern = 0
	cst_result = Dialog(dlg)



    assign "cst_result"
    If (cst_result = 0) Then Exit All

  scst_core_ri = dlg.ri  ' Core radius
  scst_core_ra = dlg.ra	' core width
  scst_core_h = dlg.h	' core height
  scst_wire_r = dlg.wr	' wire radius
  scst_wire_N = dlg.n	' number of turns
  scst_core_i = dlg.h	' core height
  scst_core_ang = dlg.ang ' angle of windings
  scst_core_off = dlg.off ' offset of angle
  scst_symm_term = 1  ' former dlg.symm_term	'symm terminals
  scst_lead = dlg.ld
  cst_kern = Cint(dlg.option_kern)


  assign "scst_core_ri"       ' writes e.g. "cst_core_r = 0.1"     into history list
  assign "scst_core_ra"
  assign "scst_core_h"
  assign "scst_wire_r"
  assign "scst_wire_N"
  assign "scst_symm_term"

  'assing "scst_core_ang"
  EndHide

        cst_result = Evaluate(cst_result)
        If (cst_result =0) Then Exit All   ' if cancel/help is clicked, exit all
        If (cst_result =1) Then Exit All
        cst_core_r       = 0.5 * (Evaluate(scst_core_ri) + Evaluate(scst_core_ra))
		cst_core_w       = Evaluate(scst_core_ra) - Evaluate(scst_core_ri)
		cst_core_h       = Evaluate(scst_core_h)
		cst_wire_r       = Evaluate(scst_wire_r)
		cst_wire_N       = Evaluate(scst_wire_N)
		cst_core_ang	 = Evaluate(scst_core_ang)
		cst_core_off	 = Evaluate(scst_core_off)
		cst_lead	 = Evaluate(scst_lead)
		cst_symm_term = cint(scst_symm_term)
		cst_core_ra = Evaluate(scst_core_ra)
		cst_core_ri = Evaluate(scst_core_ri)
		cst_phases_N = Evaluate(dlg.ph)



Debug.Print(cst_kern)

On Error Resume Next
 Curve.DeleteCurve "core_curve"
 Curve.DeleteCurve "wire_crosssection"
 Curve.DeleteCurve "torrus_curve"
On Error GoTo 0


If cst_kern = 1 Then 'check if core needs to be created
	 'start with core creation
	 Curve.NewCurve "torrus_curve"
With Layer
     .Reset
     .Name "Ferrite"
     .FrqType "hf"
     .Type "Normal"
     .Epsilon "1.0"
     .Mue "1.0"
     .Kappa "0.0"
     .TanD "0.0"
     .TanDFreq "0.0"
     .TanDGiven "False"
     .TanDModel "ConstTanD"
     .KappaM "0.0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .DispModelEps "None"
     .DispModelMue "None"
     .Rho "0.0"
     .Colour "0.501961", "0.501961", "0.501961"
     .Wireframe "False"
     .Transparency "0"
     .Create
 End With

  With Polygon3D
     .Reset
     .Name "torrus_3dpolygon"
     .Curve "torrus_curve"
     For cst_torrus_i = 0 To 360 'full circle!
     	cst_torrus_x = cst_core_r*Cos(cst_torrus_i*pi/180)
     	cst_torrus_y = cst_core_r*Sin(cst_torrus_i*pi/180)
     	cst_torrus_z = 0
      .Point cst_torrus_x, cst_torrus_y, cst_torrus_z
     Next cst_torrus_i
     .Create
 End With

Curve.NewCurve "torrus_crosssection"


With Rectangle
     .Reset
     .Name "rect1"
     .Curve "torrus_crosssection"
     .Xrange cst_core_ri+cst_wire_r, cst_core_ra-cst_wire_r
     .Yrange -cst_core_h/2+cst_wire_r, cst_core_h/2-cst_wire_r
     .Create
End With

With SweepCurve
     .Reset
     .Name Solid.GetNextFreeName
     .Component "Ferrite"
     .Material "Ferrite"
     .Twistangle "0.0"
     .Taperangle "0.0"
     .ProjectProfileToPathAdvanced "True"
     .Path "torrus_curve:torrus_3dpolygon"
     .Curve "torrus_crosssection:rect1"
     .Create
End With




End If

For phases_i = 1 To cst_phases_N

	If phases_i = 2 Then
		cst_core_off = cst_core_off + 0.66*pi
	ElseIf phases_i = 3 Then
		cst_core_off = cst_core_off + 0.66*pi


	End If

With Polygon3D
     .Reset
     .Name "core_3dpolygon"
     .Curve "core_curve"

     cst_core_i = 0

     cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
      .Point cst_core_x, cst_core_y, cst_core_z



      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=0.5*cst_core_h+cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z



     For cst_core_i = 1 To cst_wire_N-1


      'cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang) 'original
      'cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang) 'original
      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)	'with offset
      cst_core_z=0.5*cst_core_h+cst_wire_r
      .Point cst_core_x, cst_core_y, cst_core_z



      cst_core_x=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r+0.5*cst_core_w+cst_wire_r)*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=-0.5*cst_core_h-cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

      cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
      cst_core_z=0.5*cst_core_h+cst_wire_r
     .Point cst_core_x, cst_core_y, cst_core_z

     Debug.Print(cst_core_x)
	Debug.Print(cst_core_y)
	Debug.Print(cst_core_z)

     If cst_core_i = cst_wire_N-1 Then
     cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
     cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i+1)/cst_wire_N*cst_core_ang+cst_core_off)
     cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead
     .Point cst_core_x, cst_core_y, cst_core_z
     End If

   Next cst_core_i
   cst_core_x=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Cos((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
   cst_core_y=(cst_core_r-(0.5*cst_core_w+cst_wire_r))*Sin((cst_core_i)/cst_wire_N*cst_core_ang+cst_core_off)
   cst_core_z=0.5*cst_core_h+cst_wire_r+cst_lead

      'cst_core_x= cst_core_r+1.*cst_wire_r+.5*cst_core_w
      'cst_core_y=0
      'cst_core_z=0.5*cst_core_h+cst_wire_r

     '.Point cst_core_x, cst_core_y, cst_core_z

    .Create



End With


'@ new curve: wire_crosssection
If cst_wire_r > 0 Then
	Curve.NewCurve "wire_crosssection"
End If

'@ store picked point: 1

Pick.NextPickToDatabase "1"
Pick.PickCurveEndpointFromId "core_curve:core_3dpolygon", "1"



If cst_wire_r > 0 Then

'@ define curve circle: wire_crosssection:circle1

With Circle
     .Reset
     .Name "circle1"
     .Curve "wire_crosssection"
     .Radius cst_wire_r
     .Xcenter cst_core_x 'cst_core_r+0.5*cst_core_w+cst_wire_r
     .Ycenter cst_core_y '"0"
     .Segments "0"
     .Create
End With

 With Transform
     .Reset
     .Name "wire_crosssection:circle1"
     .Origin "ShapeCenter"
     .Center "0", "0", "0"
     .Angle "0", "90", "0"
     .MultipleObjects "False"
     .GroupObjects "False"
     .Repetitions "1"
     .MultipleSelection "False"
     .Transform "Curve", "Rotate"
 End With

 With Transform
     .Reset
     .Name "wire_crosssection:circle1"
     .Vector "0", "0", cst_core_h/2+cst_wire_r
     .UsePickedPoints "False"
     .InvertPickedPoints "False"
     .MultipleObjects "False"
     .GroupObjects "False"
     .Repetitions "1"
     .MultipleSelection "False"
     .Transform "Curve", "Translate"
 End With



With Material
	If Not .Exists("Wire_material") Then
		.Reset
		.Name "Wire_material"
		.FrqType "hf"
		.Type "Pec"
		.Rho "0.0"
		.Colour "1", "0.501961", "0"
		.Wireframe "False"
		.Transparency "0"
		.Reflection "True"
		.Create
	End If
End With


'@ define sweepprofile: core:wire

With SweepCurve
     .Reset
     .Name Solid.GetNextFreeName
     .Component "Wire_material"
     .Material "Wire_material"
     .Twistangle "0.0"
     .Taperangle "0.0"
     .ProjectProfileToPathAdvanced "True"
     .Path "core_curve:core_3dpolygon"
     .Curve "wire_crosssection:circle1"
     .Create
End With

End If
Next phases_i



End Sub



