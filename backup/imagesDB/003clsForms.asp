<%
	class clsForms
		private intUploadPath
		
		private sub class_initialize
			intUploadPath = Application("intUploadPath")
			session.LCID = 1033
		end sub
		
		public function frmAditionalInfo
			dim objAI, id, title, desc, url, intStatus, erro
			
			id = getForm("id")
			title = getForm("title")
			desc = getForm("desc")
			url = getForm("url")
			intStatus = getForm("status")
			
			if title = "" then erro = erro & "Fill the Title.<br>"
			if desc = "" then erro = erro & "Fill the Description.<br>"
			
			if erro <> "" then
				frmAditionalInfo = erro
				exit function
			end if
			
			if intStatus = "1" then
				intStatus = true
			else
				intStatus = false
			end if
			
			set objAI = new clsAditionalInfo
			if ValidaNumero(id) then objAI.setID(id)
			objAI.setTitle(title)
			objAI.setDesc(desc)
			objAI.setURL(url)
			objAI.setStatus(intStatus)
			objAI.MngAditionalInfo
			set objAI = nothing
		end function
		
		public function frmUsefulLink
			dim objUL, id, title, desc, url, intStatus, intType, erro
			
			id = getForm("id")
			title = getForm("title")
			desc = getForm("desc")
			url = getForm("url")
			intStatus = getForm("status")
			intType = getForm("type")
			
			if title = "" then erro = erro & "Fill the Title.<br>"
			if desc = "" then erro = erro & "Fill the Description.<br>"
			
			if erro <> "" then
				frmUsefulLink = erro
				exit function
			end if
			
			if intStatus = "1" then
				intStatus = true
			else
				intStatus = false
			end if

			if intType = "1" then
				intType = true
			else
				intType = false
			end if
			
			set objUL = new clsUsefulLink
			if ValidaNumero(id) then objUL.setID(id)
			objUL.setTitle(title)
			objUL.setDesc(desc)
			objUL.setURL(url)
			objUL.setStatus(intStatus)
			objUL.setType(intType)
			objUL.MngUsefulLink
			set objUL = nothing
		end function

		public function frmPropertyType
			dim objPT, id, intName, erro
			
			id = getForm("id")
			intName = getForm("name")
			
			if intName = "" then erro = erro & "Fill the Name.<br>"
			
			if erro <> "" then
				frmPropertyType = erro
				exit function
			end if
			
			set objPT = new clsPropertyType
			if ValidaNumero(id) then objPT.setID(id)
			objPT.setName(intName)
			objPT.MngPropertyType
			set objPT = nothing
		end function
		
		public function frmProperty
			dim objP, id, cit_id, pt_id, cli_id, intName, address, intNumber, desc, price
			dim txt1, txt2, arrImg, arrImage, sellDate, intStatus, upload, i, erro, aKeys
			dim vtourUrl, vtourDesc
			
			Set Upload = New FreeASPUpload
			Upload.Save(intUploadPath)
			
			If Err.Number <> 0 then 
				frmProperty = "Falha no Upload.<br>"
				Exit function
			end if

			id = trim(upload.form("id"))
			cit_id = trim(upload.form("cit_id"))
			pt_id = trim(upload.form("pt_id"))
			cli_id = trim(upload.form("cli_id"))
			intName = trim(upload.form("name"))
			address = trim(upload.form("address"))
			intNumber = trim(upload.form("number"))
			desc = trim(upload.form("desc"))
			price = trim(upload.form("price"))
			txt1 = trim(upload.form("txt1"))
			txt2 = trim(upload.form("txt2"))
			intStatus = trim(upload.form("status"))
			sellDate = trim(upload.form("syear")) & "/" & trim(upload.form("smonth")) & "/" & trim(upload.form("sday"))
			vtourURL = trim(upload.form("vtourURL"))
			vtourDesc = trim(upload.form("vtourDesc"))
			erro = ""
			
			redim arrImg(1)
			aKeys = Upload.UploadedFiles.Keys
			For i = 0 To Upload.UploadedFiles.Count -1
				'response.write aKeys(i)
				if trim(aKeys(i)) = "file1" then
					arrImg(0) = Upload.UploadedFiles.Item(aKeys(i)).FileName
				elseif trim(aKeys(i)) = "file2" then
					arrImg(1) = Upload.UploadedFiles.Item(aKeys(i)).FileName
				end if
			Next
			set upload = nothing
			
			if not validaNumero(cit_id) then erro = erro & "Select the City.<br>"
			if not validaNumero(pt_id) then erro = erro & "Select Property Type.<br>"
			if not validaNumero(cli_id) then erro = erro & "Client not Identified.<br>"
			if intName = "" then erro = erro & "Fill the Name.<br>"
			if not isDate(sellDate) then sellDate = ""
			
			if erro <> "" then
				frmProperty = erro
				exit function
			end if

			set objP = new clsProperty
			if validanumero(id) then objp.setID(id)
			objP.setCityID(cit_id)
			objP.setPropertyTypeID(pt_id)
			objP.setClientID(cli_id)
			objP.setname(intName)
			objP.setAddress(address)
			objP.setNumber(intNumber)
			objP.setDesc(desc)
			objP.setPrice(price)
			objP.setTxt1(txt1)
			objP.setTxt2(txt2)
			if arrImg(0) <> "" then objP.setImg1(arrImg(0))
			if arrImg(1) <> "" then objP.setImg2(arrImg(1))
			objP.setSellDate(sellDate)
			objP.setStatus(intStatus)
			objP.setVtourURL(vtourURL)
			objP.setVtourDesc(vtourDesc)
			objP.mngProperty
			set objP = nothing
		end function
		
		public function frmPropertyImage
			dim obj, id, pro_id, title, desc, img, arrImage, main, aerial, intStatus, upload, erro
			
			Set Upload = New FreeASPUpload
			Upload.Save(intUploadPath)
			
			If Err.Number <> 0 then 
				frmProperty = "Falha no Upload.<br>"
				Exit function
			end if
			
			id = trim(upload.form("id"))
			pro_id = trim(upload.form("pro_id"))
			title = trim(upload.form("title"))
			desc = trim(upload.form("desc"))
			main = trim(upload.form("main"))
			aerial = trim(upload.form("aerial"))
			intStatus = trim(upload.form("status"))
			erro = ""
			
			redim arrImage(0)
			for each img in Upload.UploadedFiles.Items
				arrImage(0) = img.fileName
			next
			set upload = nothing
			
			if not validaNumero(pro_id) then erro = erro & "Property not identified.<br>"
			if title = "" then erro = erro & "Fill the title.<br>"
			
			if erro <> "" then
				frmPropertyImage = erro
				exit function
			end if
			
			if main = "1" then
				main = true
			else
				main = false
			end if
			
			if aerial = "1" then
				aerial = true
			else
				aerial = false
			end if
			
			if intStatus = "1" then 
				intStatus = true
			else
				intStatus = false
			end if
			
			set obj = new clsPropertyImage
			if validaNumero(id) then obj.setID(id)
			obj.setPropertyID(pro_id)
			obj.setTitle(title)
			obj.setDesc(desc)
			obj.setImage(arrImage(0))
			obj.setMain(main)
			obj.setAerial(aerial)
			obj.setStatus(intStatus)
			obj.mngPropertyImage
			set obj = nothing
		end function
		
		public function frmClient
			dim objC, id, cit_id, intName, email, address, zipCode, areaCode, phone, erro
			
			id = getForm("id")
			cit_id = getForm("cit_id")
			intName = getForm("name")
			email = getForm("email")
			address = getForm("address")
			zipCode = getForm("zipCode")
			areaCode = getForm("areaCode")
			phone = getForm("phone")
			erro = ""
			
			if not validaNumero(cit_id) then erro = erro & "Select the City.<br>"
			if intName = "" then erro = erro & "Fill the Name.<br>"
			if email <> "" and not ValidaEmail(email) then erro = erro & "E-mail invalid format.<br>"
			if zipCode <> "" and not ValidaNumero(zipCode) then erro = erro & "Zip Code: only numbers.<br>"
			if areaCode <> "" or phone <> "" then
				if not ValidaNumero(areaCode) then erro = erro & "Area Code: only numbers.<br>"
				if not ValidaNumero(phone) then erro = erro & "Phone: only numbers.<br>"
			end if
			
			if erro <> "" then
				frmClient = erro
				exit function
			end if
			
			set objC = new clsClient
			if validaNUmero(id) then objC.setID(id)
			objC.setCityID(cit_id)
			objC.setName(intName)
			objC.setEmail(email)
			objC.setAddress(address)
			objC.setZipCode(zipCode)
			objC.setAreaCode(areaCode)
			objC.setPhone(phone)
			objC.MngClient
			set objC = nothing
		end function

		public function frmTestimonial
			dim objT, id, intName, desc, image(0), arrImage, intStatus, erro, Upload
			
			Set Upload = New FreeASPUpload
			Upload.Save(intUploadPath)
			
			If Err.Number <> 0 then 
				frmTestimonial = "Falha no Upload.<br>"
				Exit function
			end if

			id = upload.form("id")
			intName = upload.form("name")
			desc = upload.form("desc")
			intStatus = upload.form("status")
			
			for each arrImage in Upload.UploadedFiles.Items
				image(0) = arrImage.fileName
			next
			set upload = nothing
			
			if intName = "" then erro = erro & "Fill the Name.<br>"
			if desc = "" then erro = erro & "Fill the Description.<br>"
			
			if erro <> "" then
				frmTestimonial = erro
				exit function
			end if
			
			if intStatus = "1" then
				intStatus = true
			else
				intStatus = false
			end if

			set objT = new clsTestimonial
			if validaNumero(id) then objT.setID(id)
			objT.setName(intName)
			objT.setDesc(desc)
			objT.setImage(image(0))
			objT.setStatus(intStatus)
			objT.MngTestimonial
			set objT = nothing
		end function
		
		public function frmSampleAd
			dim objT, id, intTitle, desc, image, arrImage, url, intStatus, erro, Upload, aKeys, i
			
			Set Upload = New FreeASPUpload
			Upload.Save(intUploadPath)
			
			If Err.Number <> 0 then 
				frmSampleAd = "Falha no Upload.<br>"
				Exit function
			end if

			id = trim(upload.form("id"))
			intTitle = trim(upload.form("title"))
			desc = trim(upload.form("desc"))
			url = trim(upload.form("url"))
			intStatus = trim(upload.form("status"))
			
			redim arrImage(1)
			aKeys = Upload.UploadedFiles.Keys
			For i = 0 To Upload.UploadedFiles.Count -1
				if trim(aKeys(i)) = "image" then
					arrImage(0) = Upload.UploadedFiles.Item(aKeys(i)).FileName
				elseif trim(aKeys(i)) = "document" then
					arrImage(1) = Upload.UploadedFiles.Item(aKeys(i)).FileName
				end if
			Next
			set upload = nothing
			
			if intTitle = "" then erro = erro & "Fill the Title.<br>"
			if url = "" then erro = erro & "Fill the URL.<br>"
			if desc = "" then erro = erro & "Fill the Description.<br>"
			
			if erro <> "" then
				frmSampleAd = erro
				exit function
			end if
			
			if intStatus = "1" then
				intStatus = true
			else
				intStatus = false
			end if

			set objT = new clsSampleAd
			if validaNumero(id) then objT.setID(id)
			objT.setTitle(intTitle)
			objT.setDesc(desc)
			if arrImage(0) <> "" then objT.setImage(arrImage(0))
			if arrImage(1) <> "" then objT.setDocument(arrImage(1))
			objT.setURL(url)
			objT.setStatus(intStatus)
			objT.MngSampleAd
			set objT = nothing
		end function


		public function frmFeature
			dim obj, id, intName
			
			id = getForm("id")
			intName = getForm("name")
			
			if intName = "" then
				frmFeature = "Fill the feature name.<br>"
				exit function
			end if
			
			set obj = new clsFeature
			if validaNumero(id) then obj.setID(id)
			obj.setName(intName)
			obj.mngFeature
			set obj = nothing
		end function
		
		public function frmPropertyFeature
			dim obj, pf_id, pro_id, desc, intStatus, erro
			
			pf_id = getForm("pf_id")
			pro_id = getForm("pro_id")
			desc = getForm("desc")
			intStatus = getForm("status")
			
			if not validaNumero(pf_id) then erro = erro & "Select feature.<br>"
			if not validaNumero(pro_id) then erro = erro & "Property ID not identified.<br>"
			if desc = "" then erro = erro & "Fill the description.<br>"
			
			if erro <> "" then 
				frmPropertyFeature = erro
				exit function
			end if
			
			if intStatus = "1" then
				intStatus = true
			else
				intStatus = false
			end if
			
			set obj = new clsPropertyFeature
			obj.setPropertyID(pro_id)
			obj.setFeatureID(pf_id)
			obj.setDesc(desc)
			obj.setStatus(intStatus)
			obj.mngPropertyFeature
			set obj = nothing
		end function
		
		public function frmState
			dim objS, id, intName, abreviation, erro
			
			id = getForm("id")
			intName = getForm("name")
			abreviation = getForm("abreviation")
			
			if intName = "" then erro = erro & "Fill the Name of State.<br>"
			if abreviation = "" then erro = erro & "Fille the Abreviation.<br>"
			
			if erro <> "" then
				frmState = erro
				exit function
			end if
			
			set objS = new clsState
			if validaNumero(id) then objS.setID(id)
			objS.setName(intName)
			objS.setAbreviation(abreviation)
			objS.MngState
			set objS = nothing
		end function

		public function frmCity
			dim obj, id, stt_id, intName, erro
			
			id = getForm("id")
			stt_id = getForm("stt_id")
			intName = getForm("name")

			
			if not ValidaNumero(stt_id) then erro = erro & "Select the State.<br>"
			if intName = "" then erro = erro & "Fill the Name of City.<br>"
		
			if erro <> "" then
				frmCity = erro
				exit function
			end if
			
			set obj = new clsCity
			if validaNumero(id) then obj.setID(id)
			obj.setName(intName)
			obj.setStateID(stt_id)
			obj.MngCity
			set obj = nothing
		end function

		public function frmSecText
			dim obj, id, section, title, desc, img, arrImage, footer, intStatus, erro, upload, url

			Set Upload = New FreeASPUpload
			Upload.Save(intUploadPath)
			
			If Err.Number <> 0 then 
				frmSecText = "Falha no Upload.<br>"
				Exit function
			end if

			id = upload.form("id")
			section = upload.form("section")
			title = upload.form("title")
			desc = upload.form("desc")
			footer = upload.form("footer")
			intStatus = upload.form("status")
			url = upload.form("url")
			
			redim arrImage(0)
			for each img in Upload.UploadedFiles.Items
				arrImage(0) = img.fileName
			next
			set upload = nothing
			
			if not validanumero(section) then erro = erro & "Failure post data.<br>"
			if title = "" then erro = erro & "Fill the Title.<br>"
			if desc = "" then erro = erro & "Fill the Description.<br>"
			
			if erro <> "" then
				frmSecText = erro
				exit function
			end if
			
			if intStatus = "1" then
				intStatus = true
			else
				intStatus = false
			end if

			set obj = new clsSecText
			if validanumero(id) then obj.setID(id)
			obj.setSectionID(section)
			obj.setTitle(title)
			obj.setDesc(desc)
			obj.setImage(arrImage(0))
			obj.setFooter(footer)
			obj.setStatus(intStatus)
			obj.setURL(url)
			obj.mngSecText
			set obj = nothing
		end function
		
		public function frmUser
			dim obj, id, intName, login, pass1, pass2, intStatus, erro
			
			id = getForm("id")
			intName = getForm("name")
			login = getForm("login")
			pass1 = getForm("pass1")
			pass2 = getForm("pass2")
			intStatus = getForm("status")
			erro = ""
			
			if not validaNumero(id) then id = -1
			if intName = "" then erro = erro & "Fill the name.<br>"
			if len(login) < 4 then erro = erro & "Login between 4 and 20 carachters.<br>"
			if id = -1 and pass1 = "" then erro = erro & "Fill the password.<br>"
			if pass1 <> "" and (len(pass1) < 4 or len(pass1) > 8) then erro = erro & "Password betwess 4 and 8 carachters.<br>"
			if pass1 <> pass2 then erro = erro & "The password fields are invalid.<br>"
			
			if erro <> "" then 
				frmUser = erro
				exit function
			end if
			
			set obj = new clsUser
			obj.setID(id)
			if obj.ExistsLogin(login) then
				frmUser = "Login already exists.<br>"
				set obj = nothing
				exit function
			end if
			
			obj.setName(intName)
			obj.setLogin(login)
			obj.setPassword(pass1)
			obj.setStatus(intStatus)
			obj.mngUser
			set obj = nothing
		end function
		
		public function frmUserChangePassword
			dim obj, id, pass1, pass2, pass3, erro
			
			'id = getForm("id")
			pass1 = getForm("pass1")
			pass2 = getForm("pass2")
			pass3 = getForm("pass3")
			
			if pass1 = "" then erro = erro & "Fill Atual password field.<br>"
			if len(pass2) < 4 then erro = erro & "New password is invalid. <br>It must have between 4 and 8 caracters.<br>"
			if pass2 <> pass3 then erro = erro & "The New password and Confirm new password fiels are invalid."
			
			if erro <> "" then
				frmUserChangePassword = erro
				exit function
			end if
			
			set obj = new clsUser
			obj.setId(session("usr_id"))
			obj.fndUser
			if obj.chkUser(obj.getLogin , pass1) then
				obj.changeNewPassword(pass2)
			else
				set obj = nothing
				frmUserChangePassword = "Your password is invalid. Please, try again."
				exit function
			end if
			set obj = nothing
		end function

		public function frmUserAccess
			dim obj, usr_id, acs_id
			
			usr_id = getForm("usr_id")
			acs_id = getForm("acs_id")
			
			if not validaNumero(usr_id) and not validaNumero(acs_id) then
				frmUserAccess = "Invalida post data.<br>"
				exit function
			end if
			
			set obj = new clsUser
			obj.setID(usr_id)
			obj.mngUserAccess(acs_id)
			set obj = nothing
		end function
		
		public function frmNews
			dim obj, id, title, desc, reference, intStatus, endDate, erro
			
			id = getForm("id")
			title = getForm("title")
			desc = getForm("desc")
			reference = getForm("reference")
			intStatus = getForm("status")
			endDate = getForm("year") & "/" & getForm("month") & "/" & getForm("day")
			erro = ""
			
			if title = "" then erro = erro & "Fill the Title.<br>"
			if desc = "" then erro = erro & "Fille the description.<br>"
			if not isDate(endDate) then endDate = ""
			
			if erro <> "" then 
				frmNews = erro
				exit function
			end if
			
			if intStatus = "1" then
				intStatus = true
			else
				intStatus = false
			end if
			
			set obj = new clsNews
			if validaNumero(id) then obj.setID(id)
			obj.setTitle(title)
			obj.setDesc(desc)
			obj.setReference(reference)
			obj.setStatus(intStatus)
			obj.setEndDate(endDate)
			obj.mngNews
			set obj = nothing
		end function
end class
%>