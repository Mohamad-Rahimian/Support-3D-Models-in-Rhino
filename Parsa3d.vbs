Option Explicit
'Script written by <Mohamad Rahimian>
'Script copyrighted by <www.Rahimian-jewelry.com  parsa-3d>
'Script version Sunday, July 30, 2017 1:45:42 PM
Dim groupButtom,gcenter,gbet0,gbet1
Dim PointOrg
Dim groupintCenter
Dim groupintbet0
Dim groupintbet1
Dim planex ,planey 
planex = 96
planey = 60
Dim address,address1,address2,address3,drive

drive = "c:\"

address = drive & "Parsa3d-B.txt"
address1 = drive & "Parsa3d-C.txt"
address2 = drive & "Parsa3d-B1.txt"
address3 = drive & "Parsa3d-B2.txt"


Call main()

Sub main()
	groupButtom = "Bottom_support"
	gcenter = "Gcenter_support"
	gbet0 = "gbet0_support"
	gbet1 = "gbet1_support" 
	groupintbet0 = 1
	groupintbet1 = 1
	groupintCenter = 1
End Sub


Sub OnTop()
	Dim strSurface,strobj, strPoint:Dim ar(3):Dim add:Dim spz(2):Dim sp3
	Dim sp :Dim arrPlane:Dim VecNormal:Dim arrDomain:Dim vecnormal2
	Dim getradtop:Dim getraddown:Dim arr3
	Dim arrPick, arrCP, arrTypeIndex, intIndex
	Dim cir:Dim lin:Dim cap:Dim opt:Dim dent,trad,brad,blrad,blh,rc,sthe
	
	Const optionget = 3
	dent = Array(optionget, "Dent", 0.1, "New value Dent", 0.1, 1.0)
	trad = Array(optionget, "TopRad", 0.25, "New value TopRad", 0.15, 1.0)
	brad = Array(optionget, "ButtomRad", 0.5, "New value ButtomRad", 0.25, 2.0)
	blrad = Array(optionget, "BRad", 1.5, "New value BRad", 1.0, 3.0)
	blh = Array(optionget, "BHeight", 0.25, "New value BHeight", 0.2, 1.0)	
	sthe = Array(optionget, "STHeight", 2, "New value STHeight", 1.0, 5.0)	
	opt = Array(dent, trad, brad, blrad, blh, sthe)
	
	strobj = Rhino.GetObjectEx("Pick a surface", 16, True)
	If IsNull(strobj) Then 
		Exit Sub 
	End If
	arrCP = Rhino.BrepClosestPoint(strobj(0), strobj(3))
	arrTypeIndex = arrCP(2)
	intIndex = arrTypeIndex(1)		
	strSurface = Rhino.ExtractSurface(strobj(0), CInt(intIndex))
	Do
		rc = Rhino.GetOption("Scripting options", opt)
		If IsNull(rc) Then
			Exit Do
		End If
		If isArray(rc) Then
			opt(0)(2) = rc(0):opt(1)(2) = rc(1):opt(2)(2) = rc(2):opt(3)(2) = rc(3):opt(4)(2) = rc(4):opt(5)(2) = rc(5):
		End If
		Do
			If Not IsNull(strobj) Then	
				rhino.SelectObject strsurface
				strPoint = Rhino.GetPointOnSurface(strSurface, "Point on surface")	
				If Not IsNull(strPoint) Then	
					arrDomain = Rhino.surfaceclosestPoint(strSurface, strPoint)				
					VecNormal = Rhino.surfacenormal(strSurface, arrDomain)				
					VecNormal2 = Rhino.surfacenormal(strSurface, arrDomain)				
					vecNormal = Rhino.vectorscale(Rhino.vectorunitize(vecNormal), rc(5))				
					vecnormal2 = Rhino.VectorScale(rhino.VectorUnitize(vecnormal2), -rc(0))
					sp = Rhino.PointAdd(strPoint, vecNormal)
					sp3 = rhino.PointAdd(strPoint, vecnormal2)
					spz(0) = sp(0):spz(1) = sp(1):spz(2) = 0
					ar(0) = sp3:ar(1) = strPoint:ar(2) = sp:ar(3) = spz:
					add = Rhino.AddCurve(ar, 3)	
					Rhino.AddPipe add, Array(0, 1), Array(rc(1), rc(2)), 1, 1, 0
					cir = Rhino.AddCircle(spz, rc(3))
					lin = Rhino.AddLine(spz, Array(spz(0), spz(1), rc(4)))
					cap = Rhino.ExtrudeCurve(cir, lin)
					Rhino.CapPlanarHoles cap
					rhino.DeleteObject lin
					rhino.DeleteObject cir
					rhino.DeleteObject add					
				Else 
					Exit Do
				End If
			Else
				Exit Do
			End If		
		Loop
		If Not IsArray(rc) Then			
			Exit Do
		End If
	Loop
	rhino.JoinSurfaces Array(strobj(0), strsurface), True
	Rhino.UnselectAllObjects
	
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////


Sub OnPoint()
	Dim obj,sur,line,opt,inx,inx1,inx2,a,b,c,d,e,f,ptsur,surclp,surnorm,vec,vec2
	Dim pt1,pt2,crv,dom1,dom2,crv2,crv3,crv4,pipe,dom3,dom4,opt2
	a = Array(3, "DentTop", -0.1, "New value ", -0.5, 1.0)
	b = Array(3, "HeightTop", -0.2, "New value ", -1.0, 2.0)
	c = Array(3, "TopRadius", 0.25, "New value ", 0.1, 2.0)
	d = Array(3, "HeightDown", 1, "New value ", 0.1, 2.0)
	e = Array(3, "ButtomRadius", 0.4, "New value ", 0.1, 2.0)
	f = Array(a, b, c, d, e)
	obj = rhino.GetObjectEx("get Surface : ", 8 + 16)
	If IsNull(obj) Then Exit Sub
	inx = rhino.BrepClosestPoint(obj(0), obj(3))
	inx1 = inx(2)
	inx2 = CInt(inx1(1))
	sur = rhino.ExtractSurface(obj(0), inx2)
	line = rhino.GetObject("get line ", 4)
	Do
		opt = Rhino.GetOption("Scripting options", f)
		If IsNull(opt)Then Exit Do
		f(0)(2) = opt(0)
		f(1)(2) = opt(1)
		f(2)(2) = opt(2)
		f(3)(2) = opt(3)
		f(4)(2) = opt(4)
		opt2 = opt(4)
		rhino.SelectObject sur
		Do
			ptsur = rhino.GetPointOnSurface(sur, "Get point on surface ")
			If IsNull(ptsur)Then Exit Do
			surclp = rhino.SurfaceClosestPoint(sur, ptsur)
			surnorm = rhino.SurfaceNormal(sur, surclp)
			vec = Rhino.vectorscale(Rhino.vectorunitize(surnorm), opt(1))
			vec2 = rhino.VectorScale(rhino.VectorUnitize(surnorm), opt(0))
			pt1 = rhino.PointAdd(ptsur, vec)
			pt2 = rhino.PointAdd(ptsur, vec2)
			crv = rhino.AddCurve(Array(pt1, pt2), 3)
			dom1 = rhino.CurveDomain(line)
			dom2 = rhino.CurveDomain(crv)
			crv2 = rhino.AddBlendCurve(Array(line, crv), Array(dom1(1), dom2(0)), Array(False, False), Array(2, 2))
			dom3 = Rhino.CurveDomain(crv2)
			rhino.AddPipe crv2, ArraY(dom3(1), dom3(0)), Array(opt(2), opt(4)), 1, 1, 0
			rhino.DeleteObject crv
			rhino.DeleteObject crv2
		Loop
	Loop
	If Not IsNull(line)Then
		dom4 = rhino.CurveDomain(line)
		If Not IsNull(opt2)Then	
			rhino.AddPipe line, Array(dom4(0), dom4(1)), Array(opt2, opt2), 1, 1, 0
		End If
	End If
	rhino.JoinSurfaces Array(obj(0), sur), True
	Rhino.UnselectAllObjects

End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////


Sub OnPoint1()
	Dim obj,sur,line,opt,inx,inx1,inx2,a,b,c,d,e,f,ptsur,surclp,surnorm,vec,vec2
	Dim pt1,pt2,crv,dom1,dom2,crv2,crv3,crv4,pipe,dom3,dom4,opt2
	a = Array(3, "DentTop", -0.1, "New value ", -0.5, 1.0)
	b = Array(3, "HeightTop", 1, "New value ", 0.1, 2.0)
	c = Array(3, "TopRadius", 0.2, "New value ", 0.1, 2.0)
	d = Array(2, "CAP", 1, " 1 OR 2 ", 1.0, 2.0)
	e = Array(3, "ButtomRadius", 0.2, "New value ", 0.1, 2.0)
	f = Array(a, b, c, d, e)
	obj = rhino.GetObjectEx("get Surface : ", 8 + 16)
	If IsNull(obj) Then Exit Sub
	inx = rhino.BrepClosestPoint(obj(0), obj(3))
	inx1 = inx(2)
	inx2 = CInt(inx1(1))
	sur = rhino.ExtractSurface(obj(0), inx2)
	line = rhino.GetPointOnSurface(obj(0), "Pick up Point on Surface ")
	Do
		opt = Rhino.GetOption("Scripting options", f)
		If IsNull(opt)Then Exit Do
		f(0)(2) = opt(0)
		f(1)(2) = opt(1)
		f(2)(2) = opt(2)
		f(3)(2) = opt(3)
		f(4)(2) = opt(4)
		opt2 = opt(4)
		rhino.SelectObject sur
		Do
			ptsur = rhino.GetPointOnSurface(sur, "Get point on surface ")
			If IsNull(ptsur)Then Exit Do
			surclp = rhino.SurfaceClosestPoint(sur, ptsur)
			surnorm = rhino.SurfaceNormal(sur, surclp)
			vec = Rhino.vectorscale(Rhino.vectorunitize(surnorm), opt(1))
			vec2 = rhino.VectorScale(rhino.VectorUnitize(surnorm), opt(0))
			pt1 = rhino.PointAdd(ptsur, vec)
			pt2 = rhino.PointAdd(ptsur, vec2)
			crv2 = Array(line, pt1, pt2)
			crv = rhino.AddCurve(crv2, 3)
			dom3 = rhino.CurveDomain(crv)
			rhino.AddPipe crv, ArraY(dom3(1), dom3(0)), Array(opt(2), opt(4)), 1, opt(3), 0
			rhino.DeleteObject crv
		Loop
	Loop
	rhino.JoinSurfaces Array(obj(0), sur), True
	Rhino.UnselectAllObjects
End Sub



'///////////////////////////////////////////////////////////////////////////////////////////////////////////////


Sub OnPoint2()
	Dim obj,sur,line,opt,inx,inx1,inx2,a,b,c,d,e,f,ptsur,surclp,surnorm,vec,vec2
	Dim pt1,pt2,crv,dom1,dom2,crv2,crv3,crv4,pipe,dom3,dom4,opt2
	a = Array(3, "DentTop", -0.1, "New value ", -0.5, 1.0)
	b = Array(3, "HeightTop", 1, "New value ", 0.1, 2.0)
	c = Array(3, "TopRadius", 0.25, "New value ", 0.1, 2.0)
	d = Array(2, "CAP", 1, " 1 OR 2 ", 1.0, 2.0)
	e = Array(3, "ButtomRadius", 0.4, "New value ", 0.1, 2.0)
	f = Array(a, b, c, d, e)
	obj = rhino.GetObjectEx("get Surface : ", 8 + 16)
	If IsNull(obj) Then Exit Sub
	inx = rhino.BrepClosestPoint(obj(0), obj(3))
	inx1 = inx(2)
	inx2 = CInt(inx1(1))
	sur = rhino.ExtractSurface(obj(0), inx2)
	line = rhino.GetPoint("Pick up Point")
	Do
		opt = Rhino.GetOption("Scripting options", f)
		If IsNull(opt)Then Exit Do
		f(0)(2) = opt(0)
		f(1)(2) = opt(1)
		f(2)(2) = opt(2)
		f(3)(2) = opt(3)
		f(4)(2) = opt(4)
		opt2 = opt(4)
		rhino.SelectObject sur
		Do
			ptsur = rhino.GetPointOnSurface(sur, "Get point on surface ")
			If IsNull(ptsur)Then Exit Do
			surclp = rhino.SurfaceClosestPoint(sur, ptsur)
			surnorm = rhino.SurfaceNormal(sur, surclp)
			vec = Rhino.vectorscale(Rhino.vectorunitize(surnorm), opt(1))
			vec2 = rhino.VectorScale(rhino.VectorUnitize(surnorm), opt(0))
			pt1 = rhino.PointAdd(ptsur, vec)
			pt2 = rhino.PointAdd(ptsur, vec2)
			crv2 = Array(line, pt1, pt2)
			crv = rhino.AddCurve(crv2, 3)
			dom3 = rhino.CurveDomain(crv)
			rhino.AddPipe crv, ArraY(dom3(1), dom3(0)), Array(opt(2), opt(4)), 1, opt(3), 0
			rhino.DeleteObject crv
		Loop
	Loop
	rhino.JoinSurfaces Array(obj(0), sur), True
	Rhino.UnselectAllObjects
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////


Sub OnPoint3()
	Dim obj,opt,a,b,c,d,e,f,ptsur
	Dim crv,dom1,dom2,crv2,crv3,crv4,pipe,dom3,dom4,opt2,point1,point2
	a = Array(3, "DentTop", -0.1, "New value ", -0.5, 0.0)
	b = Array(3, "HeightTop", 1, "New value ", 0.1, 4.0)
	c = Array(3, "TopRadius", 0.25, "New value ", 0.1, 2.0)
	d = Array(3, "DentDown", -0.1, " 1 OR 2 ", -0.5, 0.0)
	e = Array(3, "ButtomRadius", 0.25, "New value ", 0.1, 2.0)
	f = Array(a, b, c, d, e)
	obj = rhino.GetObjectEx("Pick Up Object  ", 16)
	If IsNull(obj) Then Exit Sub
	Do
		opt = Rhino.GetOption("Scripting options", f)
		If IsNull(opt)Then Exit Do
		f(0)(2) = opt(0)
		f(1)(2) = opt(1)
		f(2)(2) = opt(2)
		f(3)(2) = opt(3)
		f(4)(2) = opt(4)
		opt2 = opt(4)
		Do
			point1 = rhino.GetPointOnSurface(obj(0), "Pick Up First Point")
			If IsNull(point1)Then 
				Exit Do
			End If
			point2 = rhino.GetPointOnSurface(obj(0), "Pick Up Scound Point")
			If isNull(point2)Then 
				Exit Do
			End If
			crv2 = Array(point1, point2)
			crv = rhino.AddCurve(crv2, 3)
			dom3 = rhino.CurveDomain(crv)
			rhino.AddPipe crv, ArraY(dom3(1), dom3(0)), Array(opt(2), opt(4)), 1, 2, 0
			rhino.DeleteObject crv
			point1 = Null:point2 = Null
		Loop
	Loop
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////


Sub OnPoint4()
	Dim obj,point,scp,x,y,z,arrP,arrC,cr,pr,arrD,sph,arrp2,arrC2,cr2,arrD2,pr2
	Dim trad,brad,getobt,opt,mrad,Lrad,Frad
	Dim arrobj
	Const optionget = 3
	
	Dim a,b,c,d,e,f
	Dim strSrf, arrPt, arrParam, arrNormal, arrScale, arrAdd
	Dim ex,exm,arrScale1, arrAdd1,ii
	
	Dim Point1,Point2,Point3
	Dim Point11,Point12,Point13
	Dim Curve1,Curve2,Curve3
	Dim pipeDo1,pipeDo2,pipeDo3
	Dim bottom,Up,Mid,Input,LenghtUp,Mid1
	Dim group,pip1,pip2,pip3,objname
	
	Dim er,ar
	
	b = Array(3, "TopRadius", 0.2, "New value ", 0.15, 3.0)
	c = Array(3, "MediumRadius", 0.3, "New value ", 0.15, 3.0)
	a = Array(3, "ButtomRadius", 0.2, "New value ", 0.15, 3.0)
	d = Array(3, "Lenght", 0.75, "New value ", 0.15, 3.0)
	e = Array(3, "Input", 0.05, "New value ", 0.15, 1.0)	
	
	opt = Array(b, c, a, d, e)
	er = rhino.ReadTextFile(address2)
	
	If Not isnull(er) Then
		opt(0)(2) = er(0)
		opt(1)(2) = er(1)
		opt(2)(2) = er(2)
		opt(3)(2) = er(3)
		opt(4)(2) = er(4)
	End If		
	
	Do
		getobt = Rhino.GetOption("Scripting options", opt)
		
		If Not isnull(getobt) Then
			opt(0)(2) = getobt(0)
			opt(1)(2) = getobt(1)
			opt(2)(2) = getobt(2)
			opt(3)(2) = getobt(3)
			opt(4)(2) = getobt(4)
			
			
			input = getobt(4)
			LenghtUp = getobt(3)
			bottom = getobt(2)
			Mid = getobt(1)
			Up = getobt(0)
			
			ar = array(getobt(0), getobt(1), getobt(2), getobt(3), getobt(4))
			Call Rhino.WriteTextFile(address2, ar)
			Do	
			  
				ii = Rhino.GetObjectex("select Point ", 8 + 16 + 32)				
				If isnull(ii) Then Exit Do
				objname = ii(0) & gbet0
				
				Rhino.EnableRedraw(False)
				If rhino.IsMesh(ii(0))Then 

					ex = rhino.MeshClosestPoint(ii(0), ii(3))
					exm = rhino.MeshFaceNormals(ii(0))		
					arrScale = Rhino.VectorScale(exm(ex(1)), LenghtUp)
					arrScale1 = Rhino.VectorScale(exm(ex(1)), -Input)
					Mid1 = (Mid * 10) / 100
					
					Point2 = Rhino.VectorAdd(arrScale, ii(3))
					Point1 = Rhino.VectorAdd(arrScale1, ii(3))	
					Point3 = Array(Point2(0), Point2(1), 0)			
						
					Curve1 = rhino.AddCurve(Array(Point1, Point2))
		
					pipeDo1 = rhino.CurveDomain(Curve1)
			
					pip1 = Rhino.AddPipe(Curve1, Array(pipeDo1(0), pipeDo1(1)), Array(Up, Mid - Mid1), 1, 1, 1)
		
					rhino.DeleteObject(Curve1)
					Rhino.EnableRedraw(True)
					
					ii = Rhino.GetObjectex("select Point ", 8 + 16 + 32)
					If isnull(ii) Then Exit Do
					Rhino.EnableRedraw(False)
					
					ex = rhino.MeshClosestPoint(ii(0), ii(3))
					exm = rhino.MeshFaceNormals(ii(0))		
					arrScale = Rhino.VectorScale(exm(ex(1)), LenghtUp)
					arrScale1 = Rhino.VectorScale(exm(ex(1)), -Input)
			
					Point12 = Rhino.VectorAdd(arrScale, ii(3))
					Point11 = Rhino.VectorAdd(arrScale1, ii(3))		
						
					Curve2 = rhino.AddCurve(Array(Point11, Point12))
					Curve3 = rhino.AddCurve(Array(Point2, Point12))
			
					pipeDo2 = rhino.CurveDomain(Curve2)
					pipeDo3 = rhino.CurveDomain(Curve3)
			
					pip2 = Rhino.AddPipe(Curve2, Array(pipeDo2(0), pipeDo2(1)), Array(Up, Mid - Mid1), 1, 2, 1)					
					pip3 = Rhino.AddPipe(Curve3, Array(pipeDo3(0), pipeDo3(1)), Array(Mid, Mid), 1, 1, 1)
					
					
					'Rhino.AddObjectsToGroup Array(pip1(0), pip2(0), pip3(0)), Group
					'	group = objname & groupintbet0
					'	rhino.AddGroup(group)
					'	Rhino.AddObjectsToGroup Array(pip1(0), pip2(0), pip3(0)), Group
					rhino.BooleanUnion Array(pip1(0), pip2(0), pip3(0)), True
					'	groupintbet0 = groupintbet0 + 1
					
					
					rhino.DeleteObject(Curve1)
					rhino.DeleteObject(Curve3)
					rhino.DeleteObject(Curve2)
					Rhino.EnableRedraw(True)
				Else
					ex = rhino.BrepClosestPoint(ii(0), ii(3))
					arrScale = Rhino.VectorScale(ex(3), LenghtUp)
					arrScale1 = Rhino.VectorScale(ex(3), -Input)
					Mid1 = (Mid * 10) / 100
					
					Point2 = Rhino.VectorAdd(arrScale, ii(3))
					Point1 = Rhino.VectorAdd(arrScale1, ii(3))	
					Point3 = Array(Point2(0), Point2(1), 0)			
						
					Curve1 = rhino.AddCurve(Array(Point1, Point2))
		
					pipeDo1 = rhino.CurveDomain(Curve1)
			
					pip1 = Rhino.AddPipe(Curve1, Array(pipeDo1(0), pipeDo1(1)), Array(Up, Mid - Mid1), 1, 2, 1)
		
					rhino.DeleteObject(Curve1)
					Rhino.EnableRedraw(True)
					
					ii = Rhino.GetObjectex("select Point ", 8 + 16 + 32)
					If isnull(ii) Then Exit Do
					Rhino.EnableRedraw(False)
					
					ex = rhino.BrepClosestPoint(ii(0), ii(3))
					arrScale = Rhino.VectorScale(ex(3), LenghtUp)
					arrScale1 = Rhino.VectorScale(ex(3), -Input)
			
					Point12 = Rhino.VectorAdd(arrScale, ii(3))
					Point11 = Rhino.VectorAdd(arrScale1, ii(3))		
						
					Curve2 = rhino.AddCurve(Array(Point11, Point12))
					Curve3 = rhino.AddCurve(Array(Point2, Point12))
			
					pipeDo2 = rhino.CurveDomain(Curve2)
					pipeDo3 = rhino.CurveDomain(Curve3)
			
					pip2 = Rhino.AddPipe(Curve2, Array(pipeDo2(0), pipeDo2(1)), Array(Up, Mid - Mid1), 1, 2, 1)					
					pip3 = Rhino.AddPipe(Curve3, Array(pipeDo3(0), pipeDo3(1)), Array(Mid, Mid), 1, 1, 1)
					
					'	group = objname & groupintbet0
					'	rhino.AddGroup(group)
					'	Rhino.AddObjectsToGroup Array(pip1(0), pip2(0), pip3(0)), Group
					rhino.BooleanUnion Array(pip1(0), pip2(0), pip3(0)), True
					'	groupintbet0 = groupintbet0 + 1
					
					'Rhino.AddObjectsToGroup Array(pip1(0), pip2(0), pip3(0)), Group
					
					rhino.DeleteObject(Curve1)
					rhino.DeleteObject(Curve3)
					rhino.DeleteObject(Curve2)
					Rhino.EnableRedraw(True)
				End If
				
						
			Loop
		Else Exit Do
		End If
			
	Loop


End Sub



'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sub OnPoint5()
	Dim obj,obj1,point,scp,x,y,z,arrP,arrC,cr,pr,arrD,sph,arrp2,arrC2,cr2,arrD2,pr2
	Dim trad,brad,getobt,opt,mrad,Lrad,Frad
	Dim arrobj
	Const optionget = 3
	
	Dim a,b,c,d,e,f
	Dim strSrf, arrPt, arrParam, arrNormal, arrScale, arrAdd
	Dim ex,exm,arrScale1, arrAdd1,ii
	
	Dim Point1,Point2,Point3
	Dim Point11,Point12,Point13
	Dim Curve1,Curve2,Curve3
	Dim pipeDo1,pipeDo2,pipeDo3
	Dim bottom,Up,Mid,Input,LenghtUp,Mid1
	Dim group,pip1,pip2,pip3,objname
	Dim er,ar
	
	group = "Between2_Support"
	
	b = Array(3, "TopRadius", 0.2, "New value ", 0.15, 3.0)
	c = Array(3, "MediumRadius", 0.3, "New value ", 0.15, 3.0)	
	d = Array(3, "Lenght", 0.75, "New value ", 0.15, 3.0)
	e = Array(3, "Input", 0.05, "New value ", 0.15, 1.0)
	opt = Array(b, c, d, e)
	er = rhino.ReadTextFile(address3)
	
	If Not isnull(er) Then
		opt(0)(2) = er(0)
		opt(1)(2) = er(1)
		opt(2)(2) = er(2)
		opt(3)(2) = er(3)
	End If		
	
	obj = rhino.GetObjectEx("Select First Point", 8 + 16 + 32)
	If  isnull(obj) Then Exit Sub
	objname = obj(0) & gbet1
	Point3 = obj(3)
	Do
		getobt = Rhino.GetOption("Scripting options", opt)			
		If  isnull(getobt) Then Exit Sub
		
		opt(0)(2) = getobt(0)
		opt(1)(2) = getobt(1)
		opt(2)(2) = getobt(2)
		opt(3)(2) = getobt(3)
			
			
		input = getobt(3)
		LenghtUp = getobt(2)
		Mid = getobt(1)
		Up = getobt(0)	
		
		ar = array(getobt(0), getobt(1), getobt(2), getobt(3))
		Call Rhino.WriteTextFile(address3, ar)
		
		Do
			ii = rhino.GetObjectEx("Select Point ", 8 + 16 + 32)
			If isnull(ii) Then Exit Do
			
			Rhino.EnableRedraw(False)
			If rhino.IsMesh(ii(0))Then 

				ex = rhino.MeshClosestPoint(ii(0), ii(3))
				exm = rhino.MeshFaceNormals(ii(0))		
				arrScale = Rhino.VectorScale(exm(ex(1)), LenghtUp)
				arrScale1 = Rhino.VectorScale(exm(ex(1)), -Input)
				Mid1 = (Mid * 10) / 100
				
				Point2 = Rhino.VectorAdd(arrScale, ii(3))
				Point1 = Rhino.VectorAdd(arrScale1, ii(3))	
				Point3 = obj(3)		
						
				Curve1 = rhino.AddCurve(Array(Point1, Point2))
				Curve2 = rhino.AddCurve(Array(Point2, Point3))
			
				pipeDo1 = rhino.CurveDomain(Curve1)
				pipeDo2 = rhino.CurveDomain(Curve1)
			
				pip1 = Rhino.AddPipe(Curve1, Array(pipeDo1(0), pipeDo1(1)), Array(Up, Mid - Mid1), 1, 2, 1)
				pip2 = Rhino.AddPipe(Curve2, Array(pipeDo2(0), pipeDo2(1)), Array(Mid, Mid), 1, 1, 1)
				
				'group = objname & groupintbet1
				'	rhino.AddGroup(group)
				'	Rhino.AddObjectsToGroup Array(pip1(0), pip2(0)), Group
				rhino.BooleanUnion Array(pip1(0), pip2(0)), True
				'	groupintbet1 = groupintbet1 + 1
				
				'Rhino.AddObjectsToGroup Array(pip1(0), pip2(0)), Group
				rhino.DeleteObject(Curve1)
				rhino.DeleteObject(Curve2)
				Rhino.EnableRedraw(True)
			
			Else
				ex = rhino.BrepClosestPoint(ii(0), ii(3))
				arrScale = Rhino.VectorScale(ex(3), LenghtUp)
				arrScale1 = Rhino.VectorScale(ex(3), -Input)
				Mid1 = (Mid * 10) / 100
				
				Point2 = Rhino.VectorAdd(arrScale, ii(3))
				Point1 = Rhino.VectorAdd(arrScale1, ii(3))	
				Point3 = obj(3)		
						
				Curve1 = rhino.AddCurve(Array(Point1, Point2))
				Curve2 = rhino.AddCurve(Array(Point2, Point3))
			
				pipeDo1 = rhino.CurveDomain(Curve1)
				pipeDo2 = rhino.CurveDomain(Curve1)
			
				pip1 = Rhino.AddPipe(Curve1, Array(pipeDo1(0), pipeDo1(1)), Array(Up, Mid - Mid1), 1, 2, 1)
				pip2 = Rhino.AddPipe(Curve2, Array(pipeDo2(0), pipeDo2(1)), Array(Mid, Mid), 1, 1, 1)
			
				'	group = objname & groupintbet1
				'	rhino.AddGroup(group)
				'	Rhino.AddObjectsToGroup Array(pip1(0), pip2(0)), Group
				rhino.BooleanUnion Array(pip1(0), pip2(0)), True
				'			 groupintbet1 = groupintbet1 + 1
				
				'Rhino.AddObjectsToGroup Array(pip1(0), pip2(0)), Group
				rhino.DeleteObject(Curve1)
				rhino.DeleteObject(Curve2)
				Rhino.EnableRedraw(True)
				'arrAdd = Rhino.VectorAdd(arrScale, ii(3))
			
				'Call Rhino.AddLine(ii(3), arrAdd)
			End If
		
		
		Loop
	Loop
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////



Sub Bottom()
	Dim obj,point,scp,x,y,z,arrP,arrC,cr,pr,arrD,sph,arrp2,arrC2,cr2,arrD2,pr2
	Dim trad,brad,getobt,opt,mrad,Lrad,Frad
	Dim arrobj
	Const optionget = 3
	Dim strSrf, arrPt, arrParam, arrNormal, arrScale, arrAdd
	Dim ex,exm,arrScale1, arrAdd1,ii
	
	Dim Point1,Point2,Point3,Point4,Mid1
	Dim Curve1,Curve2
	Dim pipeDo1,pipeDo2
	Dim bottom,Up,Mid,Input,LenghtUp
	Dim pip1,pip2,cone1
	Dim groupc
	Dim arrBase, arrHeight, dblRadius,strFile,ar,er
	'   Input = 0.05
	'   LenghtUp = 0.75
	'	bottom = 0.3
	'   Mid = 0.2
	'	Up = 0.1
	groupc = 1 
	trad = Array(optionget, "TopRad", 0.25, "New value TopRad", 0.15, 1.0)
	mrad = Array(optionget, "Midium", 0.30, "New value Midium", 0.15, 2.0)
	brad = Array(optionget, "ButtomRad", 0.5, "New value ButtomRad", 0.25, 2.0)
	Lrad = Array(optionget, "Lenght", 1.0, "New value Lenght", 0.2, 2.0)
	Frad = Array(optionget, "Input", 0.05, "New value Input", 0.05, 1.0)
	er = rhino.ReadTextFile(address)
	opt = Array(trad, mrad, brad, Lrad, Frad)
	
	If Not isnull(er) Then
		opt(0)(2) = er(0)
		opt(1)(2) = er(1)
		opt(2)(2) = er(2)
		opt(3)(2) = er(3)
		opt(4)(2) = er(4)
	End If
	
	Do
		getobt = Rhino.GetOption("Scripting options", opt)
		
		If Not isnull(getobt) Then
			
			
			opt(0)(2) = getobt(0)
			opt(1)(2) = getobt(1)
			opt(2)(2) = getobt(2)
			opt(3)(2) = getobt(3)
			opt(4)(2) = getobt(4)
			
			input = getobt(4)
			LenghtUp = getobt(3)
			bottom = getobt(2)
			Mid = getobt(1)
			Up = getobt(0)
			
			ar = array(getobt(0), getobt(1), getobt(2), getobt(3), getobt(4))
			Call Rhino.WriteTextFile(address, ar)
			
			Do	
				'	Rhino.DisplayOleAlerts False
				ii = Rhino.GetObjectex("select Point ", 8 + 16 + 32, False, False)
				
				If isnull(ii) Then 
					Exit Do
				End If
			
			
				Rhino.EnableRedraw(False)
				If rhino.IsMesh(ii(0))Then 

					ex = rhino.MeshClosestPoint(ii(0), ii(3))
					exm = rhino.MeshFaceNormals(ii(0))		
					arrScale = Rhino.VectorScale(exm(ex(1)), LenghtUp)
					arrScale1 = Rhino.VectorScale(exm(ex(1)), -Input)
					Mid1 = (Mid * 10) / 100
					
					Point2 = Rhino.VectorAdd(arrScale, ii(3))
					Point1 = Rhino.VectorAdd(arrScale1, ii(3))	
					Point3 = Array(Point2(0), Point2(1), 0)			
					Point4 = Array(Point2(0), Point2(1), Point2(2) + Mid)
					
					
					
					Curve1 = rhino.AddCurve(Array(Point1, Point2), 1)
					Curve2 = rhino.AddCurve(Array(Point4, Point3), 1)
			
					pipeDo1 = rhino.CurveDomain(Curve1)
					pipeDo2 = rhino.CurveDomain(Curve2)
			
					pip1 = Rhino.AddPipe(Curve1, Array(pipeDo1(0), pipeDo1(1)), Array(Up, Mid - Mid1), 1, 1, 1)
					pip2 = Rhino.AddPipe(Curve2, Array(pipeDo2(0), pipeDo2(1)), Array(Mid, bottom), 1, 1, 1)
					
					dblRadius = bottom * 2.5
					arrBase = Array(Point3(0), Point3(1), Point3(2))
					arrHeight = Array(Point3(0), Point3(1), Point3(2) + 1.5)
					cone1 = Rhino.AddCone(arrBase, arrHeight, dblRadius)
						
					groupButtom = ii(0) & "support" & groupc 
					
					'	rhino.AddGroup(groupButtom)
					'	Rhino.AddObjectsToGroup Array(pip1(0), pip2(0), cone1), groupButtom
					rhino.BooleanUnion Array(pip1(0), pip2(0), cone1), True
					'	Groupc = Groupc + 1
					rhino.DeleteObject(Curve1)
					rhino.DeleteObject(Curve2)
					Rhino.EnableRedraw(True)
			
				Else
					ex = rhino.BrepClosestPoint(ii(0), ii(3))
					arrScale = Rhino.VectorScale(ex(3), LenghtUp)
					arrScale1 = Rhino.VectorScale(ex(3), -Input)
					Mid1 = (Mid * 10) / 100
					
					Point2 = Rhino.VectorAdd(arrScale, ii(3))
					Point1 = Rhino.VectorAdd(arrScale1, ii(3))	
					Point3 = Array(Point2(0), Point2(1), 0)			
					Point4 = Array(Point2(0), Point2(1), Point2(2) + Mid)
					
					
					Curve1 = rhino.AddCurve(Array(Point1, Point2), 1)
					Curve2 = rhino.AddCurve(Array(Point4, Point3), 1)
					
					pipeDo1 = rhino.CurveDomain(Curve1)
					pipeDo2 = rhino.CurveDomain(Curve2)
					
					pip1 = Rhino.AddPipe(Curve1, Array(pipeDo1(0), pipedo1(1)), Array(Up, Mid), 1, 1, 1) 			
					pip2 = Rhino.AddPipe(Curve2, Array(pipeDo2(0), pipeDo2(1)), Array(Mid, bottom), 1, 1, 1)
					
					dblRadius = bottom * 2.5
					arrBase = Array(Point3(0), Point3(1), Point3(2))
					arrHeight = Array(Point3(0), Point3(1), Point3(2) + 1.5)
					cone1 = Rhino.AddCone(arrBase, arrHeight, dblRadius)
					
					groupButtom = ii(0) & "support" & groupc 
					
					'	rhino.AddGroup(groupButtom)
					'	Rhino.AddObjectsToGroup Array(pip1(0), pip2(0), cone1), groupButtom
					rhino.BooleanUnion Array(pip1(0), pip2(0), cone1), True
					'	Groupc = Groupc + 1
				
					rhino.DeleteObject(Curve1)
					rhino.DeleteObject(Curve2)
					Rhino.EnableRedraw(True)

				End If												
			Loop
			
		Else 
			Exit Do			
		End If
			
	Loop
	
	
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////



Sub plane()
	Dim a,b,c,d
	Dim e,f,g,h
	Dim arr
	a = Array(-planex, -planey, 0)
	b = Array(+planex, -planey, 0)
	c = Array(+planex, +planey, 0)
	d = Array(-planex, +planey, 0)
	e = rhino.AddLine(a, b)
	f = rhino.AddLine(b, c)
	g = rhino.AddLine(c, d)
	h = rhino.AddLine(d, a)
	arr = Array(e, f, g, h)
	rhino.JoinCurves arr, True
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub OnPoint6()
	Dim obj,point,scp,x,y,z,arrP,arrC,cr,pr,arrD,sph,arrp2,arrC2,cr2,arrD2,pr2
	Dim trad,brad,getobt,opt
	Dim arrobj,zl,z2,z3,z4,z5,z6,s1,s2,s3,z7,z8,z9,z10,z11
	Dim zb,zt,a,b,c,d,e,f,g
	
	
	b = Array(3, "TopRadius", 0.4, "New value ", 0.2, 5.0)
	a = Array(3, "ButtomRadius", 0.75, "New value ", 0.2, 5.0)
	c = Array(3, "Height", 7, "New value ", 2, 45)
	f = Array(a, b, c)
	Do
		opt = Rhino.GetOption("Scripting options", f)
		If opt(0) = f(0)(2) And f(1)(2) = opt(1) And f(2)(2) = opt(2) Then Exit Do
		f(0)(2) = opt(0)
		f(1)(2) = opt(1)
		f(2)(2) = opt(2)
	Loop
		
		
	point = Array(0, 0, 0)
	z7 = Array(0, 0, 1)
	zl = opt(2)
	z2 = Array(0, 0, zl)
	z3 = Array(point, z2)
	z8 = Array(point, z7)
	z4 = rhino.AddCurve(z3, 2)
	z9 = rhino.AddCurve(z8, 2)
	z5 = rhino.CurveDomain(z4)
	z10 = rhino.CurveDomain(z9)
	z6 = rhino.AddPipe(z4, Array(z5(0), z5(1)), Array(opt(0), opt(1)), 0, 1)
	z11 = rhino.AddPipe(z9, Array(z10(0), z10(1)), Array(opt(0) * 2, opt(0) * 2), 0, 1)
	rhino.DeleteObject(z4)
	rhino.DeleteObject(z9)
	s1 = rhino.AddSphere(z2, opt(1) * 2)
	PointOrg = s1
	rhino.BooleanUnion(Array(z6(0), s1, z11(0)))
	
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub GoZero()
	Dim a,b,c,xmid,ymid	
	a = rhino.GetObjects("Please Select Objects  ", 0, True, True, True)
	If Isnull(a)Then Exit Sub		
	b = rhino.BoundingBox(a)	
	xmid = (b(0)(0) + b(2)(0)) / 2
	ymid = (b(0)(1) + b(2)(1)) / 2
	rhino.MoveObjects a, Array(xmid, ymid, b(0)(2)), Array(0, 0, 2.5)	
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////


Sub Center()
	Dim obj,point,scp,x,y,z,arrP,arrC,cr,pr,arrD,sph,arrp2,arrC2,cr2,arrD2,pr2
	Dim trad,brad,getobt,opt,mrad,Lrad,Frad
	Dim arrobj,zl,z2,z3,z4,z5,z6,s1,s2,s3
	Const optionget = 3
	Dim Point1,Point2,Point3
	Dim Curve1,Curve2
	Dim pipeDo1,pipeDo2
	Dim bottom,Up,Mid,Input,LenghtUp,Mid1
	Dim ex,exm,arrScale1, arrAdd1,ii
	Dim strSrf, arrPt, arrParam, arrNormal, arrScale, arrAdd
	Dim Group,Groupint
	Dim objname
	Dim pip1,pip2
	Dim er,ar
	
	
	trad = Array(optionget, "TopRad", 0.2, "New value TopRad", 0.15, 1.0)
	mrad = Array(optionget, "Midium", 0.28, "New value Midium", 0.15, 2.0)
	brad = Array(optionget, "ButtomRad", 0.3, "New value ButtomRad", 0.25, 2.0)
	Lrad = Array(optionget, "Lenght", 0.75, "New value Lenght", 0.2, 2.0)
	Frad = Array(optionget, "Input", 0.05, "New value Input", 0.05, 1.0)
	opt = Array(trad, mrad, brad, Lrad, Frad)
	
	er = rhino.ReadTextFile(address1)
	
	If Not isnull(er) Then
		opt(0)(2) = er(0)
		opt(1)(2) = er(1)
		opt(2)(2) = er(2)
		opt(3)(2) = er(3)
		opt(4)(2) = er(4)
	End If	
	
	If Not isNull(obj) Then 
		
		obj = rhino.GetObjectEx("Please Select Center Point On Polysurface ", 8 + 16 + 32)
		If isNull(obj) Then Exit Sub
		objname = obj(0) & gcenter
		
		point = obj(3)
		zl = rhino.GetInteger("Please Set Lenght Of Line 2 to 12    ", 7, 2, 12)
		z2 = Array(point(0), point(1), point(2) + zl)
		z3 = Array(point, z2)
		z4 = rhino.AddCurve(z3, 2)
		z5 = rhino.CurveDomain(z4)
		z6 = rhino.AddPipe(z4, Array(z5(0), z5(1)), Array(0.6, 0.45), 0, 1)
		rhino.DeleteObject(z4)
		s1 = rhino.AddSphere(z2, 1)
		PointOrg = s1
		s2 = rhino.AddSphere(point, 0.65)
		
		group = objname & groupintCenter
		rhino.AddGroup(group)
		Rhino.AddObjectsToGroup Array(z6(0), s1, s2), Group
		groupintCenter = groupintCenter + 1

	End If
	
	
	If  isNull(obj) Then Exit Sub
	Do
		getobt = Rhino.GetOption("Scripting options", opt)
		If  isnull(getobt) Then Exit Do
		opt(0)(2) = getobt(0)
		opt(1)(2) = getobt(1)
		opt(2)(2) = getobt(2)
		input = getobt(4)
		LenghtUp = getobt(3)
		bottom = getobt(2)
		Mid = getobt(1)
		Up = getobt(0)
		
		ar = array(getobt(0), getobt(1), getobt(2), getobt(3), getobt(4))
		Call Rhino.WriteTextFile(address1, ar)
			
		Do	
		
			ii = Rhino.GetObjectex("select Point ", 8 + 16 + 32)
			If isnull(ii) Then Exit Do
		
			Rhino.EnableRedraw(False)
			
			If rhino.IsMesh(ii(0))Then 

				ex = rhino.MeshClosestPoint(ii(0), ii(3))
				exm = rhino.MeshFaceNormals(ii(0))		
				arrScale = Rhino.VectorScale(exm(ex(1)), LenghtUp)
				arrScale1 = Rhino.VectorScale(exm(ex(1)), -Input)
				Mid1 = (Mid * 10) / 100
				
				Point2 = Rhino.VectorAdd(arrScale, ii(3))
				Point1 = Rhino.VectorAdd(arrScale1, ii(3))	
				Point3 = z2		
						
				Curve1 = rhino.AddCurve(Array(Point1, Point2))
				Curve2 = rhino.AddCurve(Array(Point2, Point3))
			
				pipeDo1 = rhino.CurveDomain(Curve1)
				pipeDo2 = rhino.CurveDomain(Curve1)
		
				pip1 = Rhino.AddPipe(Curve1, Array(pipeDo1(0), pipeDo1(1)), Array(Up, Mid - Mid1), 1, 2, 1)
				pip2 = Rhino.AddPipe(Curve2, Array(pipeDo2(0), pipeDo2(1)), Array(Mid, bottom), 1, 1, 1)
				
				'group = objname & groupintCenter
				'	rhino.AddGroup(group)
				'	Rhino.AddObjectsToGroup Array(pip1(0), pip2(0)), Group
				rhino.BooleanUnion Array(pip1(0), pip2(0)), True
				'	groupintCenter = groupintCenter + 1
				
				rhino.DeleteObject(Curve1)
				rhino.DeleteObject(Curve2)
				
				Rhino.EnableRedraw(True)
		
			Else
				ex = rhino.BrepClosestPoint(ii(0), ii(3))
				arrScale = Rhino.VectorScale(ex(3), LenghtUp)
				arrScale1 = Rhino.VectorScale(ex(3), -Input)
				Mid1 = (Mid * 10) / 100
				
				Point2 = Rhino.VectorAdd(arrScale, ii(3))
				Point1 = Rhino.VectorAdd(arrScale1, ii(3))	
				Point3 = z2			
						
				Curve1 = rhino.AddCurve(Array(Point1, Point2))
				Curve2 = rhino.AddCurve(Array(Point2, Point3))
			
				pipeDo1 = rhino.CurveDomain(Curve1)
				pipeDo2 = rhino.CurveDomain(Curve1)
			
				pip1 = Rhino.AddPipe(Curve1, Array(pipeDo1(0), pipeDo1(1)), Array(Up, Mid - Mid1), 1, 2, 1)
				pip2 = Rhino.AddPipe(Curve2, Array(pipeDo2(0), pipeDo2(1)), Array(Mid, bottom), 1, 1, 1)
				
				'	group = objname & groupintCenter
				'	rhino.AddGroup(group)
				'	Rhino.AddObjectsToGroup Array(pip1(0), pip2(0)), Group
				rhino.BooleanUnion Array(pip1(0), pip2(0)), True
				'	groupintCenter = groupintCenter + 1
			
				
				
				rhino.DeleteObject(Curve1)
				rhino.DeleteObject(Curve2)
			
				Rhino.EnableRedraw(True)
		
			End If
						
		Loop
		
	Loop
	
End Sub
	

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////


Sub SolidWeight()
  
	' Declare local constants
	Const RH_SURFACE = 8
	Const RH_POLYSRF = 16
	Const RH_MESH = 32
	Const RH_CM = 3
  
	' Declare local variables
	Dim strObject, intType
	Dim arrVolume, dblVolume, strVolume
	Dim arrElement, strElement
	Dim dblDensity, dblScale
	Dim dblWeight, strWeight
  
	' Prompt to select a solid object
	intType = RH_SURFACE + RH_POLYSRF + RH_MESH
	strObject = Rhino.GetObject("Select solid surface, polysurface, or mesh", intType, True)
	If IsNull(strObject) Then Exit Sub
	If Not Rhino.IsObjectSolid(strObject) Then Exit Sub

	' Calculate the volume of the selected object
	dblVolume = Null
	If Rhino.IsMesh(strObject) Then
		arrVolume = Rhino.MeshVolume(strObject)
		If IsArray(arrVolume) Then dblVolume = arrVolume(1)
		Else
			arrVolume = Rhino.SurfaceVolume(strObject)
		If IsArray(arrVolume) Then dblVolume = arrVolume(0)
	End If
  
	' Verify volume
	If IsNull(dblVolume) Then
		Call Rhino.Print("Unable to calculate volume.")
		Exit Sub
	End If

	' Prompt for the element type
	arrElement = Array("Gold 24K", "White Gold 18K", "Yellow Gold 18K", "Platinum", "Silver", "Copper", "AL")
	strElement = Rhino.ListBox(arrElement, "Select element type:", "Solid Weight")
	If IsNull(strElement) Then Exit Sub

	' Density in grams per cubic centimeter
	Select Case strElement
		Case "Gold 24K"      dblDensity = 19.32
		Case "Silver"    dblDensity = 10.5
		Case "Platinum"  dblDensity = 21.4
		Case "Yellow Gold 18K" dblDensity = 17.1
		Case "White Gold 18K" dblDensity = 17.3
		Case "Copper"    dblDensity = 8.94
		Case "AL"   dblDensity = 2.70
	End Select
  
	' Convert volume to cubic centimeters
	dblScale = Rhino.UnitScale(RH_CM)
	dblVolume = dblVolume * (dblScale ^ 3)
	strVolume = FormatNumber(dblVolume, 3)
	Call Rhino.Print("Volume = " & strVolume & " cubic centimeters")
  
	' Calculate the weight in grams, where
	' Weight = Volume x Density 
	dblWeight = dblVolume * dblDensity
	strWeight = FormatNumber(dblWeight, 3)
	Call Rhino.Print("Volume = " & strWeight & " grams")
  
End Sub


'//////////////////////////////////////////////////////////////////////////////////////

Sub Boxboundari
	Dim obj,box
	obj = rhino.GetObjects("Select Any Objects", 8 + 16 + 32)
	If isnull(obj) Then Exit Sub
	box = rhino.BoundingBox(obj)
	Rhino.AddBox box
End Sub