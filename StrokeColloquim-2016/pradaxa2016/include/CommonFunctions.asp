<%
''**************************************************************
'*
'*	GetVal(strName) : for Form field value : strName- Name of the field
'*	GetQVal(strName) : for QueryString field value	: strName- Name of QS
'*
'***
'*	Following All returns Combobox Options
'***
'*	ListNumbers(N,s,e) :for listing numbers : N- selected values,s-Starting no, e- Ending no
'*	ListDate(jdDate) : for List of Date : jdDate- selected Date
'*	ListMonth(jdMonth) : for List of Month : jdMonth - selected Month
'*	ListYear(N,s,e) : for List of Years :N- selected Year, s- Starting Year, e- Ending Year
'*	ListGender(N) : for List of Gender : N- selected Gender
'*	ListTitle(N) : for List of Title : N- selected Title
'*	ListCountry(N) : for List of Country : N- selected Country
'*
'***************************************************************

Function GetVal(sFFieldName)
	Dim tmpStr
	tmpStr=Trim(Request.Form(sFFieldName))
	if tmpStr<>"" then
		tmpStr=Replace(tmpStr,"'","''")
	End if
	GetVal=tmpStr
End Function

Function GetQVal(sFFieldName)
	Dim tmpStr
	tmpStr=Trim(Request.QueryString(sFFieldName))
	'if tmpStr<>"" then
	'	tmpStr=Replace(tmpStr,"'","''")
	'End if
	GetQVal=tmpStr
End Function

Function ListNumbers(N,s,e)
	Dim NN
	for i=s to e
		NN=NN& "<option value='"&i&"'"
		if strComp(i,N)=0 then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">"&i&"</Option>"
	next
	ListNumbers=NN
End Function

Function IsEmail(e)
	if instr(e,"@")>0 and instrrev(e,".")>0 then
		isEmail="True"
	else
		isEmail="False"
	End if
End Function

Function ListDate(N)
	Dim NN
	for i=0 to 31
		NN=NN& "<option value='"&i&"'"
		if i=N then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">"&i&"</Option>"
	next
	ListDate=NN
End Function


Function dispdate(dt)
  if dt<>"" then
     dt=day(dt)&"-"&monthname(month(dt),true)&"-"&year(dt)
  else
     dt=""
  end if      
  dispdate=dt
End Function


Function ListYear(N,s,e)
	Dim NN
	for i=s to e
		NN=NN& "<option value='"&i&"'"
		if strComp(i,N)=0 then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">"&i&"</Option>"
	next
	ListYear=NN
End Function


Function ListMonth(N)
	Dim NN
		NN=NN& "<option value='1'"
		if N="1" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Jan</Option>"

		NN=NN& "<option value='2'"
		if N="2" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Feb</Option>"

		NN=NN& "<option value='3'"
		if N="3" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Mar</Option>"

		NN=NN& "<option value='4'"
		if N="4" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Apr</Option>"

		NN=NN& "<option value='5'"
		if N="5" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">May</Option>"

		NN=NN& "<option value='6'"
		if N="6" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Jun</Option>"

		NN=NN& "<option value='7'"
		if N="7" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Jly</Option>"

		NN=NN& "<option value='8'"
		if N="8" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Aug</Option>"

		NN=NN& "<option value='9'"
		if N="9" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Sep</Option>"

		NN=NN& "<option value='10'"
		if N="10" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Oct</Option>"

		NN=NN& "<option value='11'"
		if N="11" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Nov</Option>"

		NN=NN& "<option value='12'"
		if N="12" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Dec</Option>"
	ListMonth=NN
End Function

Function ListGender(N)
	Dim NN
		NN=NN& "<option value='Male'"
		if N="Male" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Male</Option>"

		NN=NN& "<option value='FeMale'"
		if N="FeMale" then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Female</Option>"
	ListGender=NN
End Function

Function ListTitle(N)
	Dim NN
		NN=NN& "<option value='Mr.'"
		if N="Mr." then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Mr.</Option>"
		
		NN=NN& "<option value='Mrs.'"
		if N="Mrs." then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Mrs.</Option>"
		
		NN=NN& "<option value='Ms.'"
		if N="Ms." then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Ms.</Option>"
		
		NN=NN& "<option value='Dr.'"
		if N="Dr." then
			NN=NN &"  SELECTED "
		End if
		NN=NN & ">Dr.</Option>"

	ListTitle=NN
End Function
Function ListCountry (C)
Dim OP(251)

 OP( 0 ) = "Select Country"
 OP( 1 ) = "USA"
 OP( 2 ) = "Afghanistan"
 OP( 3 ) = "Albania"
 OP( 4 ) = "Algeria"
 OP( 5 ) = "American Samoa"
 OP( 6 ) = "Andorra"
 OP( 7 ) = "Angola"
 OP( 8 ) = "Anguilla"
 OP( 9 ) = "Antarctica"
 OP( 10 ) = "Antigua and Barbuda"
 OP( 11 ) = "Argentina"
 OP( 12 ) = "Armenia"
 OP( 13 ) = "Aruba"
 OP( 14 ) = "Australia"
 OP( 15 ) = "Austria"
 OP( 16 ) = "Azerbaijan"
 OP( 17 ) = "Bahamas"
 OP( 18 ) = "Bahrain"
 OP( 19 ) = "Bangladesh"
 OP( 20 ) = "Barbados"
 OP( 21 ) = "Belarussia"
 OP( 22 ) = "Belgium"
 OP( 23 ) = "Belize"
 OP( 24 ) = "Benin"
 OP( 25 ) = "Bermuda"
 OP( 26 ) = "Bhutan"
 OP( 27 ) = "Bolivia"
 OP( 28 ) = "Bosnia and Herzegovina"
 OP( 30 ) = "Botswana"
 OP( 31 ) = "Bouvet Island"
 OP( 32 ) = "Brazil"
 OP( 33 ) = "British Indian Ocean Territory"
 OP( 34 ) = "Brunei Darussalam"
 OP( 35 ) = "Bulgaria"
 OP( 36 ) = "Burkina Faso"
 OP( 37 ) = "Burma"
 OP( 38 ) = "Burundi"
 OP( 39 ) = "Cambodia"
 OP( 40 ) = "Cameroon"
 OP( 41 ) = "Canada"
 OP( 42 ) = "Canton and Enderbury Islands"
 OP( 43 ) = "Cape Verde"
 OP( 44 ) = "Cayman Islands"
 OP( 45 ) = "Central African Republic"
 OP( 46 ) = "Chad"
 OP( 47 ) = "Chile"
 OP( 48 ) = "China"
 OP( 49 ) = "Christmas Island"
 OP( 50 ) = "Cocos (Keeling) Islands"
 OP( 51 ) = "Colombia"
 OP( 52 ) = "Comoros"
 OP( 53 ) = "Congo"
 OP( 54 ) = "Cook Islands"
 OP( 55 ) = "Costa Rica"
 OP( 56 ) = "Cote D'Ivoire"
 OP( 57 ) = "Croatia (Hrvatska)"
 OP( 58 ) = "Cuba"
 OP( 59 ) = "Cyprus"
 OP( 60 ) = "Czech Republic"
 OP( 61 ) = "Democratic Yemen"
 OP( 62 ) = "Denmark"
 OP( 63 ) = "Djibouti"
 OP( 64 ) = "Dominica"
 OP( 65 ) = "Dominican Republic"
 OP( 66 ) = "Dronning Maud Land"
 OP( 67 ) = "East Timor"
 OP( 68 ) = "Ecuador"
 OP( 69 ) = "Egypt"
 OP( 70 ) = "El Salvador"
 OP( 71 ) = "Equatorial Guinea"
 OP( 72 ) = "Eritrea"
 OP( 73 ) = "Estonia"
 OP( 74 ) = "Ethiopia"
 OP( 75 ) = "Falkland Islands (Malvinas)"
 OP( 76 ) = "Faroe Islands"
 OP( 77 ) = "Fiji"
 OP( 78 ) = "Finland"
 OP( 79 ) = "France"
 OP( 80 ) = "France, Metropolitan"
 OP( 81 ) = "French Guiana"
 OP( 82 ) = "French Polynesia"
 OP( 83 ) = "French Southern Territories"
 OP( 84 ) = "Gabon"
 OP( 85 ) = "Gambia"
 OP( 86 ) = "Georgia"
 OP( 87 ) = "Germany"
 OP( 88 ) = "Ghana"
 OP( 89 ) = "Gibraltar"
 OP( 90 ) = "Greece"
 OP( 91 ) = "Greenland"
 OP( 92 ) = "Grenada"
 OP( 93 ) = "Guadeloupe"
 OP( 94 ) = "Guam"
 OP( 95 ) = "Guatemala"
 OP( 96 ) = "Guinea"
 OP( 97 ) = "Guinea-bisseu"
 OP( 98 ) = "Guyana"
 OP( 99 ) = "Haiti"
 OP( 100 ) = "Heard and Mc Donald Islands"
 OP( 101 ) = "Honduras"
 OP( 102 ) = "Hong Kong"
 OP( 103 ) = "Hungary"
 OP( 104 ) = "Iceland"
 OP( 105 ) = "India"
 OP( 106 ) = "Indonesia"
 OP( 106 ) = "Iran(Islamic Republic of)"
 OP( 107 ) = "Iraq"
 OP( 108 ) = "Ireland"
 OP( 109 ) = "Israel"
 OP( 110 ) = "Italy"
 OP( 111 ) = "Ivory Coast"
 OP( 112 ) = "Jamaica"
 OP( 113 ) = "Japan"
 OP( 114 ) = "Johnston Island"
 OP( 115 ) = "Jordan"
 OP( 116 ) = "Kazakstan"
 OP( 117 ) = "Kenya"
 OP( 118 ) = "Kiribati"
 OP( 119 ) = "Korea"
 OP( 120 ) = "Kuwait"
 OP( 121 ) = "Kyrgyzstan"
 OP( 123 ) = "Latvia"
 OP( 124 ) = "Lebanon"
 OP( 125 ) = "Lesotho"
 OP( 126 ) = "Liberia"
 OP( 127 ) = "Libyan Arab Jamahiriya"
 OP( 128 ) = "Liechtenstein"
 OP( 129 ) = "Lithuania"
 OP( 130 ) = "Luxembourg"
 OP( 131 ) = "Macau"
 OP( 132 ) = "Macedonia"
 OP( 133 ) = "Madagascar"
 OP( 134 ) = "Malawi"
 OP( 135 ) = "Malaysia"
 OP( 136 ) = "Maldives"
 OP( 137 ) = "Mali"
 OP( 138 ) = "Malta"
 OP( 139 ) = "Marshall Islands"
 OP( 140 ) = "Martinique"
 OP( 141 ) = "Mauritania"
 OP( 142 ) = "Mauritius"
 OP( 143 ) = "Mayotte"
 OP( 144 ) = "Mexico"
 OP( 146 ) = "Midway Islands"
 OP( 147 ) = "Moldova, Republic of"
 OP( 148 ) = "Monaco"
 OP( 149 ) = "Mongolia"
 OP( 150 ) = "Montserrat"
 OP( 151 ) = "Morocco"
 OP( 152 ) = "Mozambique"
 OP( 153 ) = "Myanmar"
 OP( 154 ) = "Namibia"
 OP( 155 ) = "Nauru"
 OP( 156 ) = "Nepal"
 OP( 157 ) = "Netherlands"
 OP( 158 ) = "Netherlands Antilles"
 OP( 159 ) = "Neutral Zone"
 OP( 160 ) = "New Calidonia"
 OP( 161 ) = "New Zealand"
 OP( 162 ) = "Nicaragua"
 OP( 163 ) = "Niger"
 OP( 164 ) = "Nigeria"
 OP( 165 ) = "Niue"
 OP( 166 ) = "Norfolk Island"
 OP( 167 ) = "Northern Mariana Islands"
 OP( 168 ) = "Norway"
 OP( 169 ) = "Oman"
 OP( 170 ) = "Pacific Islands"
 OP( 171 ) = "Pakistan"
 OP( 172 ) = "Palau"
 OP( 173 ) = "Panama"
 OP( 174 ) = "Papua New Guinea"
 OP( 175 ) = "Paraguay"
 OP( 176 ) = "Peru"
 OP( 177 ) = "Philippines"
 OP( 178 ) = "Pitcairn Island"
 OP( 179 ) = "Poland"
 OP( 180 ) = "Portugal"
 OP( 181 ) = "Puerto Rico"
 OP( 182 ) = "Qatar"
 OP( 183 ) = "Reunion"
 OP( 184 ) = "Romania"
 OP( 185 ) = "Russian Federation"
 OP( 186 ) = "Rwanda"
 OP( 187 ) = "S. Georgia and S. Sandwich Isls."
 OP( 188 ) = "Saint Lucia"
 OP( 189 ) = "Saint Vincent/Grenadines"
 OP( 190 ) = "Samoa"
 OP( 191 ) = "San Marino"
 OP( 192 ) = "Sao Tome and Principe"
 OP( 193 ) = "Saudi Arabia"
 OP( 194 ) = "Senegal"
 OP( 195 ) = "Seychelles"
 OP( 196 ) = "Sierra Leone"
 OP( 197 ) = "Singapore"
 OP( 198 ) = "Slovakia (Slovak Republic)"
 OP( 199 ) = "Slovenia"
 OP( 200 ) = "Solomon Islands"
 OP( 201 ) = "Somalia"
 OP( 202 ) = "South Africa"
 OP( 203 ) = "Spain"
 OP( 204 ) = "Sri Lanka"
 OP( 205 ) = "St. Helena"
 OP( 206 ) = "St. Kitts Nevis Anguilla"
 OP( 207 ) = "St. Pierre and Miquelon"
 OP( 208 ) = "Sudan"
 OP( 209 ) = "Suriname"
 OP( 210 ) = "Svalbard and Jan Mayen Islands"
 OP( 211 ) = "Swaziland"
 OP( 212 ) = "Sweden"
 OP( 213 ) = "Switzerland"
 OP( 214 ) = "Syran Arab Republic"
 OP( 215 ) = "Taiwan "
 OP( 216 ) = "Tajikistan"
 OP( 217 ) = "Tanzania, United Republic of "
 OP( 218 ) = "Thailand"
 OP( 219 ) = "Togo"
 OP( 220 ) = "Tokelau"
 OP( 221 ) = "Tonga"
 OP( 222 ) = "Trinidad and Tobago"
 OP( 223 ) = "Tunisia"
 OP( 224 ) = "Turkey"
 OP( 225 ) = "Turkmenistan"
 OP( 226 ) = "Turks and Caicos Islands"
 OP( 227 ) = "Tuvalu"
 OP( 228 ) = "US Minor Outlying Islands"
 OP( 229 ) = "Uganda"
 OP( 230 ) = "Ukraine"
 OP( 231 ) = "United Arab Emirates"
 OP( 232 ) = "United Kingdom"
 OP( 233 ) = "United States"
 OP( 234 ) = "United States Pacific Islands"
 OP( 235 ) = "Upper Volta"
 OP( 236 ) = "Uruguay"
 OP( 237 ) = "Uzbekistan"
 OP( 238 ) = "Vanuatu"
 OP( 239 ) = "Vatican City State"
 OP( 240 ) = "Venezuela"
 OP( 241 ) = "VietNam"
 OP( 242 ) = "Virgin Islands, British"
 OP( 243 ) = "Virgin Islands, Unites States "
 OP( 244 ) = "Wake Island"
 OP( 245 ) = "Wallis and Futuma Islands"
 OP( 246 ) = "Western Sahara"
 OP( 247 ) = "Yemen"
 OP( 248 ) = "Yugoslavia"
 OP( 249 ) = "Zaire"
 OP( 250 ) = "Zambia"
 OP( 251 ) = "Zimbabwe"
	Dim CC
	for i=0 to 251
		If OP(i)<>"" then
			CC=CC& "<option value='"&i&"'"
			if strComp(i,C)=0 then
				CC=CC &"  SELECTED "
			End if
			CC=CC & ">"& OP(i) &"</Option>"
		End IF	
	next
	ListCountry=CC
End Function
Function CountryName(c)
 Dim OP(251)
 OP( 0 ) = "Select Country"
 OP( 1 ) = "USA"
 OP( 2 ) = "Afghanistan"
 OP( 3 ) = "Albania"
 OP( 4 ) = "Algeria"
 OP( 5 ) = "American Samoa"
 OP( 6 ) = "Andorra"
 OP( 7 ) = "Angola"
 OP( 8 ) = "Anguilla"
 OP( 9 ) = "Antarctica"
 OP( 10 ) = "Antigua and Barbuda"
 OP( 11 ) = "Argentina"
 OP( 12 ) = "Armenia"
 OP( 13 ) = "Aruba"
 OP( 14 ) = "Australia"
 OP( 15 ) = "Austria"
 OP( 16 ) = "Azerbaijan"
 OP( 17 ) = "Bahamas"
 OP( 18 ) = "Bahrain"
 OP( 19 ) = "Bangladesh"
 OP( 20 ) = "Barbados"
 OP( 21 ) = "Belarussia"
 OP( 22 ) = "Belgium"
 OP( 23 ) = "Belize"
 OP( 24 ) = "Benin"
 OP( 25 ) = "Bermuda"
 OP( 26 ) = "Bhutan"
 OP( 27 ) = "Bolivia"
 OP( 28 ) = "Bosnia and Herzegovina"
 OP( 30 ) = "Botswana"
 OP( 31 ) = "Bouvet Island"
 OP( 32 ) = "Brazil"
 OP( 33 ) = "British Indian Ocean Territory"
 OP( 34 ) = "Brunei Darussalam"
 OP( 35 ) = "Bulgaria"
 OP( 36 ) = "Burkina Faso"
 OP( 37 ) = "Burma"
 OP( 38 ) = "Burundi"
 OP( 39 ) = "Cambodia"
 OP( 40 ) = "Cameroon"
 OP( 41 ) = "Canada"
 OP( 42 ) = "Canton and Enderbury Islands"
 OP( 43 ) = "Cape Verde"
 OP( 44 ) = "Cayman Islands"
 OP( 45 ) = "Central African Republic"
 OP( 46 ) = "Chad"
 OP( 47 ) = "Chile"
 OP( 48 ) = "China"
 OP( 49 ) = "Christmas Island"
 OP( 50 ) = "Cocos (Keeling) Islands"
 OP( 51 ) = "Colombia"
 OP( 52 ) = "Comoros"
 OP( 53 ) = "Congo"
 OP( 54 ) = "Cook Islands"
 OP( 55 ) = "Costa Rica"
 OP( 56 ) = "Cote D'Ivoire"
 OP( 57 ) = "Croatia (Hrvatska)"
 OP( 58 ) = "Cuba"
 OP( 59 ) = "Cyprus"
 OP( 60 ) = "Czech Republic"
 OP( 61 ) = "Democratic Yemen"
 OP( 62 ) = "Denmark"
 OP( 63 ) = "Djibouti"
 OP( 64 ) = "Dominica"
 OP( 65 ) = "Dominican Republic"
 OP( 66 ) = "Dronning Maud Land"
 OP( 67 ) = "East Timor"
 OP( 68 ) = "Ecuador"
 OP( 69 ) = "Egypt"
 OP( 70 ) = "El Salvador"
 OP( 71 ) = "Equatorial Guinea"
 OP( 72 ) = "Eritrea"
 OP( 73 ) = "Estonia"
 OP( 74 ) = "Ethiopia"
 OP( 75 ) = "Falkland Islands(Malvinas)"
 OP( 76 ) = "Faroe Islands"
 OP( 77 ) = "Fiji"
 OP( 78 ) = "Finland"
 OP( 79 ) = "France"
 OP( 80 ) = "France, Metropolitan"
 OP( 81 ) = "French Guiana"
 OP( 82 ) = "French Polynesia"
 OP( 83 ) = "French Southern Territories"
 OP( 84 ) = "Gabon"
 OP( 85 ) = "Gambia"
 OP( 86 ) = "Georgia"
 OP( 87 ) = "Germany"
 OP( 88 ) = "Ghana"
 OP( 89 ) = "Gibraltar"
 OP( 90 ) = "Greece"
 OP( 91 ) = "Greenland"
 OP( 92 ) = "Grenada"
 OP( 93 ) = "Guadeloupe"
 OP( 94 ) = "Guam"
 OP( 95 ) = "Guatemala"
 OP( 96 ) = "Guinea"
 OP( 97 ) = "Guinea-bisseu"
 OP( 98 ) = "Guyana"
 OP( 99 ) = "Haiti"
 OP( 100 ) = "Heard and Mc Donald Islands"
 OP( 101 ) = "Honduras"
 OP( 102 ) = "Hong Kong"
 OP( 103 ) = "Hungary"
 OP( 104 ) = "Iceland"
 OP( 105 ) = "India"
 OP( 106 ) = "Indonesia"
 OP( 106 ) = "Iran(Islamic Republic of)"
 OP( 107 ) = "Iraq"
 OP( 108 ) = "Ireland"
 OP( 109 ) = "Israel"
 OP( 110 ) = "Italy"
 OP( 111 ) = "Ivory Coast"
 OP( 112 ) = "Jamaica"
 OP( 113 ) = "Japan"
 OP( 114 ) = "Johnston Island"
 OP( 115 ) = "Jordan"
 OP( 116 ) = "Kazakstan"
 OP( 117 ) = "Kenya"
 OP( 118 ) = "Kiribati"
 OP( 119 ) = "Korea"
 OP( 120 ) = "Kuwait"
 OP( 121 ) = "Kyrgyzstan"
 OP( 123 ) = "Latvia"
 OP( 124 ) = "Lebanon"
 OP( 125 ) = "Lesotho"
 OP( 126 ) = "Liberia"
 OP( 127 ) = "Libyan Arab Jamahiriya"
 OP( 128 ) = "Liechtenstein"
 OP( 129 ) = "Lithuania"
 OP( 130 ) = "Luxembourg"
 OP( 131 ) = "Macau"
 OP( 132 ) = "Macedonia"
 OP( 133 ) = "Madagascar"
 OP( 134 ) = "Malawi"
 OP( 135 ) = "Malaysia"
 OP( 136 ) = "Maldives"
 OP( 137 ) = "Mali"
 OP( 138 ) = "Malta"
 OP( 139 ) = "Marshall Islands"
 OP( 140 ) = "Martinique"
 OP( 141 ) = "Mauritania"
 OP( 142 ) = "Mauritius"
 OP( 143 ) = "Mayotte"
 OP( 144 ) = "Mexico"
 OP( 146 ) = "Midway Islands"
 OP( 147 ) = "Moldova, Republic of"
 OP( 148 ) = "Monaco"
 OP( 149 ) = "Mongolia"
 OP( 150 ) = "Montserrat"
 OP( 151 ) = "Morocco"
 OP( 152 ) = "Mozambique"
 OP( 153 ) = "Myanmar"
 OP( 154 ) = "Namibia"
 OP( 155 ) = "Nauru"
 OP( 156 ) = "Nepal"
 OP( 157 ) = "Netherlands"
 OP( 158 ) = "Netherlands Antilles"
 OP( 159 ) = "Neutral Zone"
 OP( 160 ) = "New Calidonia"
 OP( 161 ) = "New Zealand"
 OP( 162 ) = "Nicaragua"
 OP( 163 ) = "Niger"
 OP( 164 ) = "Nigeria"
 OP( 165 ) = "Niue"
 OP( 166 ) = "Norfolk Island"
 OP( 167 ) = "Northern Mariana Islands"
 OP( 168 ) = "Norway"
 OP( 169 ) = "Oman"
 OP( 170 ) = "Pacific Islands"
 OP( 171 ) = "Pakistan"
 OP( 172 ) = "Palau"
 OP( 173 ) = "Panama"
 OP( 174 ) = "Papua New Guinea"
 OP( 175 ) = "Paraguay"
 OP( 176 ) = "Peru"
 OP( 177 ) = "Philippines"
 OP( 178 ) = "Pitcairn Island"
 OP( 179 ) = "Poland"
 OP( 180 ) = "Portugal"
 OP( 181 ) = "Puerto Rico"
 OP( 182 ) = "Qatar"
 OP( 183 ) = "Reunion"
 OP( 184 ) = "Romania"
 OP( 185 ) = "Russian Federation"
 OP( 186 ) = "Rwanda"
 OP( 187 ) = "S. Georgia and S. Sandwich Isls."
 OP( 188 ) = "Saint Lucia"
 OP( 189 ) = "Saint Vincent/Grenadines"
 OP( 190 ) = "Samoa"
 OP( 191 ) = "San Marino"
 OP( 192 ) = "Sao Tome and Principe"
 OP( 193 ) = "Saudi Arabia"
 OP( 194 ) = "Senegal"
 OP( 195 ) = "Seychelles"
 OP( 196 ) = "Sierra Leone"
 OP( 197 ) = "Singapore"
 OP( 198 ) = "Slovakia (Slovak Republic)"
 OP( 199 ) = "Slovenia"
 OP( 200 ) = "Solomon Islands"
 OP( 201 ) = "Somalia"
 OP( 202 ) = "South Africa"
 OP( 203 ) = "Spain"
 OP( 204 ) = "Sri Lanka"
 OP( 205 ) = "St. Helena"
 OP( 206 ) = "St. Kitts Nevis Anguilla"
 OP( 207 ) = "St. Pierre and Miquelon"
 OP( 208 ) = "Sudan"
 OP( 209 ) = "Suriname"
 OP( 210 ) = "Svalbard and Jan Mayen Islands"
 OP( 211 ) = "Swaziland"
 OP( 212 ) = "Sweden"
 OP( 213 ) = "Switzerland"
 OP( 214 ) = "Syran Arab Republic"
 OP( 215 ) = "Taiwan "
 OP( 216 ) = "Tajikistan"
 OP( 217 ) = "Tanzania, United Republic of "
 OP( 218 ) = "Thailand"
 OP( 219 ) = "Togo"
 OP( 220 ) = "Tokelau"
 OP( 221 ) = "Tonga"
 OP( 222 ) = "Trinidad and Tobago"
 OP( 223 ) = "Tunisia"
 OP( 224 ) = "Turkey"
 OP( 225 ) = "Turkmenistan"
 OP( 226 ) = "Turks and Caicos Islands"
 OP( 227 ) = "Tuvalu"
 OP( 228 ) = "US Minor Outlying Islands"
 OP( 229 ) = "Uganda"
 OP( 230 ) = "Ukraine"
 OP( 231 ) = "United Arab Emirates"
 OP( 232 ) = "United Kingdom"
 OP( 233 ) = "United States"
 OP( 234 ) = "United States Pacific Islands"
 OP( 235 ) = "Upper Volta"
 OP( 236 ) = "Uruguay"
 OP( 237 ) = "Uzbekistan"
 OP( 238 ) = "Vanuatu"
 OP( 239 ) = "Vatican City State"
 OP( 240 ) = "Venezuela"
 OP( 241 ) = "VietNam"
 OP( 242 ) = "Virgin Islands, British"
 OP( 243 ) = "Virgin Islands, Unites States "
 OP( 244 ) = "Wake Island"
 OP( 245 ) = "Wallis and Futuma Islands"
 OP( 246 ) = "Western Sahara"
 OP( 247 ) = "Yemen"
 OP( 248 ) = "Yugoslavia"
 OP( 249 ) = "Zaire"
 OP( 250 ) = "Zambia"
 OP( 251 ) = "Zimbabwe"
CountryName = OP(c)
End Function

Function IsEuropeanCountry(C)
IsEuropeanCountry=False
Dim OP(41)
 OP( 0 ) = "Albania"
 OP( 1 ) = "Andorra"
 OP( 2 ) = "Austria"
 OP( 3 ) = "Belarus"
 OP( 4 ) = "Belgium"
 OP( 5 ) = "Bosnia and Herzegovina"
 OP( 6 ) = "Bulgaria"
 OP( 7 ) = "Croatia (Hrvatska)"
 OP( 8 ) = "Czech Republic"
 OP( 9 ) = "Denmark"
 OP( 10 ) = "Estonia"
 OP( 11 ) = "Finland"
 OP( 12 ) = "France"
 OP( 13 ) = "Georgia"
 OP( 14 ) = "Germany"
 OP( 15 ) = "Greece"
 OP( 16 ) = "Hungary"
 OP( 17 ) = "Iceland"
 OP( 18 ) = "Ireland"
 OP( 19 ) = "Italy"
 OP( 20 ) = "Latvia"
 OP( 21 ) = "Liechtenstein"
 OP( 22 ) = "Lithuania"
 OP( 23 ) = "Luxembourg"
 OP( 24 ) = "Macedonia"
 OP( 25 ) = "Malta"
 OP( 26 ) = "Moldova, Republic of"
 OP( 27 ) = "Monaco"
 OP( 28 ) = "Netherlands"
 OP( 29 ) = "Norway"
 OP( 30 ) = "Poland"
 OP( 31 ) = "Portugal"
 OP( 32 ) = "Romania"
 OP( 33 ) = "San Marino"
 OP( 34 ) = "Slovakia (Slovak Republic)"
 OP( 35 ) = "Slovenia"
 OP( 36 ) = "Spain"
 OP( 37 ) = "Sweden"
 OP( 38 ) = "Switzerland"
 OP( 39 ) = "Turkey"
 OP( 40 ) = "Ukraine"
 OP( 41 ) = "Vatican City State"
	 for i=0 to 41
		if strComp(OP(i),C)=0 then
			IsEuropeanCountry=True
		End if
	next
End Function

Function GetCountryID(CC)
Dim OP(251)

 OP( 0 ) = "Select Country"
 OP( 1 ) = "USA"
 OP( 2 ) = "Afghanistan"
 OP( 3 ) = "Albania"
 OP( 4 ) = "Algeria"
 OP( 5 ) = "American Samoa"
 OP( 6 ) = "Andorra"
 OP( 7 ) = "Angola"
 OP( 8 ) = "Anguilla"
 OP( 9 ) = "Antarctica"
 OP( 10 ) = "Antigua and Barbuda"
 OP( 11 ) = "Argentina"
 OP( 12 ) = "Armenia"
 OP( 13 ) = "Aruba"
 OP( 14 ) = "Australia"
 OP( 15 ) = "Austria"
 OP( 16 ) = "Azerbaijan"
 OP( 17 ) = "Bahamas"
 OP( 18 ) = "Bahrain"
 OP( 19 ) = "Bangladesh"
 OP( 20 ) = "Barbados"
 OP( 21 ) = "Belarussia"
 OP( 22 ) = "Belgium"
 OP( 23 ) = "Belize"
 OP( 24 ) = "Benin"
 OP( 25 ) = "Bermuda"
 OP( 26 ) = "Bhutan"
 OP( 27 ) = "Bolivia"
 OP( 28 ) = "Bosnia and Herzegovina"
 OP( 30 ) = "Botswana"
 OP( 31 ) = "Bouvet Island"
 OP( 32 ) = "Brazil"
 OP( 33 ) = "British Indian Ocean Territory"
 OP( 34 ) = "Brunei Darussalam"
 OP( 35 ) = "Bulgaria"
 OP( 36 ) = "Burkina Faso"
 OP( 37 ) = "Burma"
 OP( 38 ) = "Burundi"
 OP( 39 ) = "Cambodia"
 OP( 40 ) = "Cameroon"
 OP( 41 ) = "Canada"
 OP( 42 ) = "Canton and Enderbury Islands"
 OP( 43 ) = "Cape Verde"
 OP( 44 ) = "Cayman Islands"
 OP( 45 ) = "Central African Republic"
 OP( 46 ) = "Chad"
 OP( 47 ) = "Chile"
 OP( 48 ) = "China"
 OP( 49 ) = "Christmas Island"
 OP( 50 ) = "Cocos (Keeling) Islands"
 OP( 51 ) = "Colombia"
 OP( 52 ) = "Comoros"
 OP( 53 ) = "Congo"
 OP( 54 ) = "Cook Islands"
 OP( 55 ) = "Costa Rica"
 OP( 56 ) = "Cote D'Ivoire"
 OP( 57 ) = "Croatia (Hrvatska)"
 OP( 58 ) = "Cuba"
 OP( 59 ) = "Cyprus"
 OP( 60 ) = "Czech Republic"
 OP( 61 ) = "Democratic Yemen"
 OP( 62 ) = "Denmark"
 OP( 63 ) = "Djibouti"
 OP( 64 ) = "Dominica"
 OP( 65 ) = "Dominican Republic"
 OP( 66 ) = "Dronning Maud Land"
 OP( 67 ) = "East Timor"
 OP( 68 ) = "Ecuador"
 OP( 69 ) = "Egypt"
 OP( 70 ) = "El Salvador"
 OP( 71 ) = "Equatorial Guinea"
 OP( 72 ) = "Eritrea"
 OP( 73 ) = "Estonia"
 OP( 74 ) = "Ethiopia"
 OP( 75 ) = "Falkland Islands(Malvinas)"
 OP( 76 ) = "Faroe Islands"
 OP( 77 ) = "Fiji"
 OP( 78 ) = "Finland"
 OP( 79 ) = "France"
 OP( 80 ) = "France, Metropolitan"
 OP( 81 ) = "French Guiana"
 OP( 82 ) = "French Polynesia"
 OP( 83 ) = "French Southern Territories"
 OP( 84 ) = "Gabon"
 OP( 85 ) = "Gambia"
 OP( 86 ) = "Georgia"
 OP( 87 ) = "Germany"
 OP( 88 ) = "Ghana"
 OP( 89 ) = "Gibraltar"
 OP( 90 ) = "Greece"
 OP( 91 ) = "Greenland"
 OP( 92 ) = "Grenada"
 OP( 93 ) = "Guadeloupe"
 OP( 94 ) = "Guam"
 OP( 95 ) = "Guatemala"
 OP( 96 ) = "Guinea"
 OP( 97 ) = "Guinea-bisseu"
 OP( 98 ) = "Guyana"
 OP( 99 ) = "Haiti"
 OP( 100 ) = "Heard and Mc Donald Islands"
 OP( 101 ) = "Honduras"
 OP( 102 ) = "Hong Kong"
 OP( 103 ) = "Hungary"
 OP( 104 ) = "Iceland"
 OP( 105 ) = "India"
 OP( 106 ) = "Indonesia"
 OP( 106 ) = "Iran(Islamic Republic of)"
 OP( 107 ) = "Iraq"
 OP( 108 ) = "Ireland"
 OP( 109 ) = "Israel"
 OP( 110 ) = "Italy"
 OP( 111 ) = "Ivory Coast"
 OP( 112 ) = "Jamaica"
 OP( 113 ) = "Japan"
 OP( 114 ) = "Johnston Island"
 OP( 115 ) = "Jordan"
 OP( 116 ) = "Kazakstan"
 OP( 117 ) = "Kenya"
 OP( 118 ) = "Kiribati"
 OP( 119 ) = "Korea"
 OP( 120 ) = "Kuwait"
 OP( 121 ) = "Kyrgyzstan"
 OP( 123 ) = "Latvia"
 OP( 124 ) = "Lebanon"
 OP( 125 ) = "Lesotho"
 OP( 126 ) = "Liberia"
 OP( 127 ) = "Libyan Arab Jamahiriya"
 OP( 128 ) = "Liechtenstein"
 OP( 129 ) = "Lithuania"
 OP( 130 ) = "Luxembourg"
 OP( 131 ) = "Macau"
 OP( 132 ) = "Macedonia"
 OP( 133 ) = "Madagascar"
 OP( 134 ) = "Malawi"
 OP( 135 ) = "Malaysia"
 OP( 136 ) = "Maldives"
 OP( 137 ) = "Mali"
 OP( 138 ) = "Malta"
 OP( 139 ) = "Marshall Islands"
 OP( 140 ) = "Martinique"
 OP( 141 ) = "Mauritania"
 OP( 142 ) = "Mauritius"
 OP( 143 ) = "Mayotte"
 OP( 144 ) = "Mexico"
 OP( 146 ) = "Midway Islands"
 OP( 147 ) = "Moldova, Republic of"
 OP( 148 ) = "Monaco"
 OP( 149 ) = "Mongolia"
 OP( 150 ) = "Montserrat"
 OP( 151 ) = "Morocco"
 OP( 152 ) = "Mozambique"
 OP( 153 ) = "Myanmar"
 OP( 154 ) = "Namibia"
 OP( 155 ) = "Nauru"
 OP( 156 ) = "Nepal"
 OP( 157 ) = "Netherlands"
 OP( 158 ) = "Netherlands Antilles"
 OP( 159 ) = "Neutral Zone"
 OP( 160 ) = "New Calidonia"
 OP( 161 ) = "New Zealand"
 OP( 162 ) = "Nicaragua"
 OP( 163 ) = "Niger"
 OP( 164 ) = "Nigeria"
 OP( 165 ) = "Niue"
 OP( 166 ) = "Norfolk Island"
 OP( 167 ) = "Northern Mariana Islands"
 OP( 168 ) = "Norway"
 OP( 169 ) = "Oman"
 OP( 170 ) = "Pacific Islands"
 OP( 171 ) = "Pakistan"
 OP( 172 ) = "Palau"
 OP( 173 ) = "Panama"
 OP( 174 ) = "Papua New Guinea"
 OP( 175 ) = "Paraguay"
 OP( 176 ) = "Peru"
 OP( 177 ) = "Philippines"
 OP( 178 ) = "Pitcairn Island"
 OP( 179 ) = "Poland"
 OP( 180 ) = "Portugal"
 OP( 181 ) = "Puerto Rico"
 OP( 182 ) = "Qatar"
 OP( 183 ) = "Reunion"
 OP( 184 ) = "Romania"
 OP( 185 ) = "Russian Federation"
 OP( 186 ) = "Rwanda"
 OP( 187 ) = "S. Georgia and S. Sandwich Isls."
 OP( 188 ) = "Saint Lucia"
 OP( 189 ) = "Saint Vincent/Grenadines"
 OP( 190 ) = "Samoa"
 OP( 191 ) = "San Marino"
 OP( 192 ) = "Sao Tome and Principe"
 OP( 193 ) = "Saudi Arabia"
 OP( 194 ) = "Senegal"
 OP( 195 ) = "Seychelles"
 OP( 196 ) = "Sierra Leone"
 OP( 197 ) = "Singapore"
 OP( 198 ) = "Slovakia (Slovak Republic)"
 OP( 199 ) = "Slovenia"
 OP( 200 ) = "Solomon Islands"
 OP( 201 ) = "Somalia"
 OP( 202 ) = "South Africa"
 OP( 203 ) = "Spain"
 OP( 204 ) = "Sri Lanka"
 OP( 205 ) = "St. Helena"
 OP( 206 ) = "St. Kitts Nevis Anguilla"
 OP( 207 ) = "St. Pierre and Miquelon"
 OP( 208 ) = "Sudan"
 OP( 209 ) = "Suriname"
 OP( 210 ) = "Svalbard and Jan Mayen Islands"
 OP( 211 ) = "Swaziland"
 OP( 212 ) = "Sweden"
 OP( 213 ) = "Switzerland"
 OP( 214 ) = "Syran Arab Republic"
 OP( 215 ) = "Taiwan "
 OP( 216 ) = "Tajikistan"
 OP( 217 ) = "Tanzania, United Republic of "
 OP( 218 ) = "Thailand"
 OP( 219 ) = "Togo"
 OP( 220 ) = "Tokelau"
 OP( 221 ) = "Tonga"
 OP( 222 ) = "Trinidad and Tobago"
 OP( 223 ) = "Tunisia"
 OP( 224 ) = "Turkey"
 OP( 225 ) = "Turkmenistan"
 OP( 226 ) = "Turks and Caicos Islands"
 OP( 227 ) = "Tuvalu"
 OP( 228 ) = "US Minor Outlying Islands"
 OP( 229 ) = "Uganda"
 OP( 230 ) = "Ukraine"
 OP( 231 ) = "United Arab Emirates"
 OP( 232 ) = "United Kingdom"
 OP( 233 ) = "United States"
 OP( 234 ) = "United States Pacific Islands"
 OP( 235 ) = "Upper Volta"
 OP( 236 ) = "Uruguay"
 OP( 237 ) = "Uzbekistan"
 OP( 238 ) = "Vanuatu"
 OP( 239 ) = "Vatican City State"
 OP( 240 ) = "Venezuela"
 OP( 241 ) = "VietNam"
 OP( 242 ) = "Virgin Islands, British"
 OP( 243 ) = "Virgin Islands, Unites States "
 OP( 244 ) = "Wake Island"
 OP( 245 ) = "Wallis and Futuma Islands"
 OP( 246 ) = "Western Sahara"
 OP( 247 ) = "Yemen"
 OP( 248 ) = "Yugoslavia"
 OP( 249 ) = "Zaire"
 OP( 250 ) = "Zambia"
 OP( 251 ) = "Zimbabwe"
	GetCountryID=0
	if Trim(cc) = "UK - England, Wales and Scotland excl. Highlands and Islands" then
		GetCountryID= 232
	end if
	For i=1 to 251
	 If	strComp(Trim(OP(i)),Trim(CC))=0 Then
	 	GetCountryID=i
	 End If
	Next
End Function

function ClearHTMLTags(strHTML, intWorkFlow)
		dim regEx, strTagLess
		strTagless = strHTML
		set regEx = New RegExp 
		regEx.IgnoreCase = True
		regEx.Global = True
		if intWorkFlow <> 1 then
			regEx.Pattern = "<[^>]*>"
			strTagLess = regEx.Replace(strTagLess, "")
		end if
		
		if intWorkFlow > 0 and intWorkFlow < 3 then
			regEx.Pattern = "[<]"
			'matches a single <
			strTagLess = regEx.Replace(strTagLess, "&lt;")
			regEx.Pattern = "[>]"
			'matches a single >
			strTagLess = regEx.Replace(strTagLess, "&gt;")
		end if
		set regEx = nothing
		
		ClearHTMLTags = strTagLess
end function


Function checkfile(smallimage)
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If trim(len(smallimage)>0) then
	        FolderPath = Server.MapPath("ProductImages/Thumbnail/")
	        filename=folderpath&"\"&smallimage
	      '  response.write "<br>"&filename
	       If(objFSO.FileExists(filename))=true then
	          simageexist="yes"
          Else
        		   simageexist="no"
           End if
		else
			simageexist="no"
		End if
  set objFSO=nothing
  checkfile=simageexist
End Function

Function checkfile1(Bigimage)
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If trim(len(Bigimage)>0) then
	        FolderPath = Server.MapPath("ProductImages/Largeimage")
	        filename=folderpath&"\"&Bigimage
	        'response.write "<br>"&filename
	       If(objFSO.FileExists(filename))=true then
		          bimageexist="yes"
           Else
        		   bimageexist="no"
           End if
		else
			bimageexist="no"
		End if
  set objFSO=nothing
  checkfile1=bimageexist
End Function

'*************************************admin panel****************************
Function checkfile_admin(smallimage)
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If trim(len(smallimage)>0) then
	        FolderPath = Server.MapPath("../../ProductImages/Thumbnail/")
	        filename=folderpath&"\"&smallimage
	        'response.write "<br>"&filename
	       If(objFSO.FileExists(filename))=true then
		          simageexist="yes"
           Else
        		   simageexist="no"
           End if
		else
			simageexist="no"
		End if
  set objFSO=nothing
  checkfile_admin=simageexist
End Function

Function checkfile1_admin(Bigimage)
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If trim(len(Bigimage)>0) then
	        FolderPath = Server.MapPath("../../ProductImages/Largeimage")
	        filename=folderpath&"\"&Bigimage
	        'response.write "<br>"&filename
	       If(objFSO.FileExists(filename))=true then
		          bimageexist="yes"
           Else
        		   bimageexist="no"
           End if
		else
			bimageexist="no"
		End if
  set objFSO=nothing
  checkfile1_admin=bimageexist
End Function



'****************************************************************************
Function checkBrandfile(Brandimage)
   		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If trim(len(Brandimage)>0) then
	        FolderPath = Server.MapPath("ProductImages/brand")
	        filename=folderpath&"\"&Brandimage
	        'response.write "<br>"&filename
	       If(objFSO.FileExists(filename))=true then
		          Brandimageexist="yes"
           Else
        		   Brandimageexist="no"
           End if
		else
			Brandimageexist="no"
		End if
  set objFSO=nothing
  checkBrandfile=Brandimageexist
End Function

Function checkAdminBrandfile(Brandimage)
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If trim(len(Brandimage)>0) then
	        FolderPath = Server.MapPath("../../ProductImages/brand")
	        filename=folderpath&"\"&Brandimage
	        'response.write "<br>"&filename
	       If(objFSO.FileExists(filename))=true then
		          Brandimageexist="yes"
           Else
        		   Brandimageexist="no"
           End if
		else
			Brandimageexist="no"
		End if
  set objFSO=nothing
  checkAdminBrandfile=Brandimageexist
End Function


Function sqlgen (catyid)
set rs11=server.createobject("adodb.recordset")
set rs4=server.createobject("adodb.recordset")
set rs6=server.createobject("adodb.recordset")
set rs7=server.createobject("adodb.recordset")
set rs8=server.createobject("adodb.recordset")

sql="select categoryid from category where pid='"&catyid&"'"
if rs11.state=1 then rs11.close
rs11.open sql,con,1,3
'response.write rs11.recordcount
'	response.write "<br>"
while not rs11.eof 
	'categoryid=categoryid&rs11("categoryid")&","
	if rs4.state=1 then rs4.close
	sql1="select * from category where pid="&rs11("categoryid")&""
'	response.write "<br>child:"&sql1
	rs4.open sql1,con,1,3
		if rs4.recordcount then
			while not rs4.eof 
				sql2="select * from category where pid="&rs4("categoryid")&""
				
				'response.write "<br>subchild:<br>"&sql2
				if rs6.state=1 then rs6.close
				rs6.open sql2,con,1,3
				if rs6.recordcount then
						while not rs6.eof 
							sql3="select * from category where pid="&rs6("categoryid")&""
							if rs7.state=1 then rs7.close
							rs7.open sql3,con,1,3
							if rs7.recordcount then
								'child=child&rs7("categoryid")
								while not rs7.eof 
									sql4="select * from category where pid="&rs7("categoryid")&"'"
									if rs8.state=1 then rs8.close
									rs8.open sql4,con,1,3
									if rs8.recordcount then
										child=child&rs8("categoryid")&","
									else
										child=child&rs7("categoryid")&","
									end if
								rs7.movenext
								wend
							else
								child=child&rs6("categoryid")&","	
							end if
						rs6.movenext
						wend
					'child=child&rs6("categoryid")&","
					'response.write "Subchild true"
				else
					child=child&rs4("categoryid")&","
				end if
		rs4.movenext
		wend 	
	else
		child=child&rs11("categoryid")&","
	end if
rs11.movenext
wend

'	response.write "<br>"

 spcategory=Split(child,",")
  for i=0 to UBound(spcategory)
  
	nwcategory=nwcategory & "'" & spcategory(i) & "',"
	count=count+1
  Next
   category_id=LEFT(nwcategory,Len(nwcategory)-1)
'	response.write category_id
'	response.write "<BR>"
	category_id=category_id&",'"&id&"'"
	'response.write category_id&"*********"
'	sqlproduct="select * from product where categoryid in ("&category_id&") order by productname"
	sqlproduct="select * from product where categoryid in ("&category_id&") and disable='0' and featuredproduct='1' order by productname"
sqlgen=sqlproduct	
'	response.write sqlproduct
end function






Function GetCountrycodeforCC(Countname)

Select case Countname
		Case "Algeria"
			name="DZA"
		Case "Afghanistan"
			name="AFG"
		Case "Albania"
			name="ALB"		
		Case "Andorra"
			name="AND"
		Case "Angola"
			name="AGO"
		Case "Anguilla"
			name="AIA"
		Case "Argentina"
			name="ARG"		
		Case "Armenia"
			name="ARM"

		Case "Aruba"
			name="ABW"
		Case "Australia"
			name="AUS"
		Case "Austria"
			name="AUT"		
		Case "Azerbaijan"
			name="AZE"


		Case "Bahamas"
			name="BHS"
		Case "Bahrain"
			name="BHR"
		Case "Bangladesh"
			name="BGD"		
		Case "Barbados"
			name="BRB"


		Case "Belarus"
			name="BLR"
		Case "Belgium"
			name="BEL"
		Case "Belize"
			name="BLZ"		
		Case "Benin"
			name="BEN"

		Case "Bermuda"
			name="BMU"
		Case "Bhutan"
			name="BTN"
		Case "Bolivia"
			name="BOL"		
		Case "Bosnia and Herzegovina"
			name="BIH"

		Case "Botswana"
			name="BWA"
		Case "Brazil"
			name="BRA"
		Case "British Virgin Islands"
			name="VGB"		
		Case "Brunei"
			name="BRN"
		
		Case "Bulgaria"
			name="BGR"
		Case "Burkina Faso"
			name="BFA"
		Case "Burundi"
			name="BDI"		
		Case "Cambodia"
			name="KHM"		
			
		Case "Cameroon"
			name="CMR"
		Case "Canada"
			name="CAN"
		Case "Cape Verde"
			name="CPV"		
		Case "Central African Republic"
			name="CAF"		

		Case "Chad"
			name="TCD"
		Case "Chile"
			name="CHL"
		Case "China"
			name="CHN"		
		Case "Colombia"
			name="COL"	
				
		Case "Comoros"
			name="COM"
		Case "Congo"
			name="COG"
		Case "Costa Rica"
			name="CRI"		
		Case "Croatia"
			name="HRV"	

		Case "Cuba"
			name="CUB"
		Case "Cyprus"
			name="CYP"
		Case "Czech Republic"
			name="CZE"		
		Case "Côte d'Ivoire"
			name="CIV"	
						
						
		Case "Denmark"
			name="DNK"
		Case "Djibouti"
			name="DJI"
		Case "Dominica"
			name="DMA"		
		Case "Dominican Republic"
			name="DOM"	

		Case "East Timor"
			name="TMP"
		Case "Ecuador"
			name="ECU"
		Case "Egypt"
			name="EGY"		
		Case "El Salvador"
			name="SLV"				
						
		Case "Equatorial Guinea"
			name="GNQ"
		Case "Eritrea"
			name="ERI"
		Case "Estonia"
			name="EST"		
		Case "Ethiopia"
			name="ETH"	
				
				
		Case "Fiji"
			name="FJI"
		Case "Finland"
			name="FIN"
		Case "France"
			name="FRA"		
		Case "French Guiana"
			name="GUF"	
				
		Case "French Polynesia"
			name="PYF"
		Case "French Southern Territories"
			name="ATF"
		Case "Gabon"
			name="GAB"		
		Case "Gambia"
			name="GMB"					

		Case "Georgia"
			name="GEO"
		Case "Germany"
			name="DEU"
		Case "Ghana"
			name="GHA"		
		Case "Greece"
			name="GRC"					

		Case "Guadeloupe"
			name="GLP"
		Case "Guatemala"
			name="GTM"
		Case "Guinea"
			name="GIN"		
		Case "Guinea-Bissau"
			name="GNB"	
				
		Case "Guyana"
			name="GUY"
		Case "Haiti"
			name="HTI"
		Case "Honduras"
			name="HND"		
		Case "Hong Kong"
			name="HKG"					
			
		Case "Hungary"
			name="HUN"
		Case "Iceland"
			name="ISL"
		Case "India"
			name="IND"		
		Case "Indonesia"
			name="IDN"	

		Case "Iran"
			name="IRN"
		Case "Iraq"
			name="IRQ"
		Case "Ireland"
			name="IRL"		
		Case "Israel"
			name="ISR"	
			
		Case "Italy"
			name="ITA"
		Case "Jamaica"
			name="JAM"
		Case "Japan"
			name="JPN"		
		Case "Jordan"
			name="JOR"				

		Case "Kazakhstan"
			name="KAZ"
		Case "Kenya"
			name="KEN"
		Case "Kiribati"
			name="KIR"		
		Case "Kuwait"
			name="KWT"	

		Case "Kyrgyzstan"
			name="KGZ"
		'Case "Laos"
			'name="LA"
		Case "Latvia"
			name="LVA"		
		Case "Lebanon"
			name="LBN"	

		Case "Lesotho"
			name="LSO"
		Case "Liberia"
			name="LBR"
		Case "Libya"
			name="LBY"		
		Case "Liechtenstein"
			name="LIE"
			
		Case "Lithuania"
			name="LTU"
		Case "Luxembourg"
			name="LUX"
		Case "Macedonia"
			name="MKD"		
		Case "Madagascar"
			name="MDG"	

		Case "Malaysia"
			name="MYS"
		Case "Mali"
			name="MLI"
		Case "Malta"
			name="MLT"		
		Case "Martinique"
			name="MTQ"	
			
		Case "Mauritania"
			name="MRT"
		Case "Mauritius"
			name="MUS"
		Case "Mayotte"
			name="MYT"		
		Case "Mexico"
			name="MEX"	
			
		Case "Micronesia"
			name="FSM"
		Case "Moldova"
			name="MDA"
		Case "Monaco"
			name="MCO"		
		Case "Mongolia"
			name="MNG"				
			
		Case "Montserrat"
			name="MSR"
		Case "Morocco"
			name="MAR"
		Case "Mozambique"
			name="MOZ"		
		Case "Myanmar"
			name="MMR"	
			
		Case "Namibia"
			name="NAM"
		Case "Nepal"
			name="NPL"
		Case "Netherlands"
			name="NLD"		
		Case "Netherlands Antilles"
			name="ANT"	
			
		Case "New Caledonia"
			name="NCL"
		Case "New Zealand"
			name="NZL"
		Case "Nicaragua"
			name="NIC"		
		Case "Niger"
			name="NER"				
			
		Case "Nigeria"
			name="NGA"
		Case "Niue"
			name="NIU"
		'Case "North Korea"
			'name="KP"		
		Case "Norway"
			name="NOR"	
			
		Case "Oman"
			name="OMN"
		Case "Pakistan"
			name="PAK"
		Case "Panama"
			name="PAN"		
		Case "Papua New Guinea"
			name="PNG"					
					

		Case "Paraguay"
			name="PRY"
		Case "Peru"
			name="PER"
		Case "Philippines"
			name="PHL"		
		Case "Poland"
			name="POL"		

		Case "Portugal"
			name="PRT"
		Case "Puerto Rico"
			name="PRI"
		Case "Qatar"
			name="QAT"		
		Case "Romania"
			name="ROM"	
			
			
		Case "Russia"
			name="RUS"
		Case "Rwanda"
			name="RWA"
		Case "Saudi Arabia"
			name="SAU"		
		Case "Senegal"
			name="SEN"	
			
		Case "Seychelles"
			name="SYC"
		Case "Sierra Leone"
			name="SLE"
		Case "Singapore"
			name="SGP"		
		Case "Slovakia"
			name="SVK"		
			
			
		Case "Slovenia"
			name="SVN"
		Case "Somalia"
			name="SOM"
		Case "South Africa"
			name="ZAF"		
		'Case "South Korea"
		'	name="ZAF"			
			
			
		Case "Spain"
			name="ESP"
		Case "Sri Lanka"
			name="LKA"
		Case "Sudan"
			name="SDN"		
		Case "Suriname"
			name="SUR"				
			
		Case "Swaziland"
			name="SWZ"
		Case "Sweden"
			name="SWE"
		Case "Switzerland"
			name="CHE"		
		Case "Syria"
			name="SYR"			
			
			
		Case "Taiwan"
			name="TWN"
		Case "Tajikistan"
			name="TJK"
		Case "Tanzania"
			name="TZA"		
		Case "Thailand"
			name="THA"		
			
		Case "Togo"
			name="TGO"
		Case "Tokelau"
			name="TKL"
		Case "Tonga"
			name="TON"		
		Case "Trinidad and Tobago"
			name="TTO"				

		Case "Tunisia"
			name="TUN"
		Case "Turkey"
			name="TUR"
		Case "Turkmenistan"
			name="TKM"		
		Case "U.S. Virgin Islands"
			name="VIR"	


		Case "Uganda"
			name="UGA"
		Case "Ukraine"
			name="UKR"
		Case "United Arab Emirates"
			name="ARE"		
		Case "United Kingdom"
			name="GBR"	
			
			
		Case "United States"
			name="USA"
		Case "Uruguay"
			name="URY"
		Case "Uzbekistan"
			name="UZB"		
		Case "Vanuatu"
			name="VUT"	
			
		Case "Vatican"
			name="VAT"
		Case "Venezuela"
			name="VEN"
		Case "Vietnam"
			name="VNM"		
		Case "Western Sahara"
			name="ESH"				
						
						
		Case "Yemen"
			name="YEM"
		Case "Yugoslavia"
			name="YUG"
		Case "Zambia"
			name="ZMB"		
		Case "Zimbabwe"
			name="ZWE"				

End Select

GetCountrycodeforCC=NAME
	
End Function





function sqlgen1_new(catlid)
set rs3_new=server.createobject("adodb.recordset")	'1
set rs4_new=server.createobject("adodb.recordset")	'2
set rs5_new=server.createobject("adodb.recordset")	'3
set rs6_new=server.createobject("adodb.recordset")	'4
set rs7_new=server.createobject("adodb.recordset")	'5
'prodcutcatid

'*****************First level*********************************
'response.write "first<BR>"
sql1="select categoryid from category where pid="&catlid&" and pid is not null"
if rs3_new.state=1 then rs3_new.close
rs3_new.open sql1,con,1,3
'response.write sql1&"<br>"
'response.write rs.recordcount&"<br>"
	if not rs3_new.eof then
		while not rs3_new.eof
		child=child&rs3_new("categoryid")&","
		rs3_new.movenext
		wend
 level1child_id=LEFT(child,Len(child)-1)  
    
'response.write level1child_id&"<br>"
'response.write child1&"<br>"
'response.end
prodcutcatid= level1child_id&","
else
prodcutcatid= catlid&","
end if
'*****************First level*********************************
'*****************second level*********************************
if level1child_id<>"" then
'response.write "2:<BR>"
		if rs4_new.state=1 then rs4_new.close
		sql2="select categoryid from category where pid in ("&level1child_id&")"
		'response.write sql2&"<br>"
		rs4_new.open sql2,con,1,3
		'response.write rs4_new.recordcount&"<br>"
			if not rs4_new.eof then
				while not rs4_new.eof
					child1=child1&rs4_new("categoryid")&","
				rs4_new.movenext
				wend

 level1child2_id=LEFT(child1,Len(child1)-1)
' level1child2_id=LEFT(child,Len(child)-1)
'response.write level1child2_id&"<br>"
'response.write child2&"<br>"
prodcutcatid=prodcutcatid&level1child2_id&","
end if
end if
'*****************second level*********************************
'*****************third level*********************************
if level1child2_id<>"" then
'response.write "3:<BR>"
		if rs5_new.state=1 then rs5_new.close
		sql3="select categoryid from category where pid in ("&level1child2_id&")"
'		response.write sql3&"<br>"
		rs5_new.open sql3,con,1,3
'		response.write rs5_new.recordcount&"<br>"
			if not rs5_new.eof then
				while not rs5_new.eof
					child2=child2&rs5_new("categoryid")&","
				rs5_new.movenext
				wend
 level1child3_id=LEFT(child2,Len(child2)-1)
'response.write level1child3_id&"<br>"
'response.write child2&"<br>"
prodcutcatid=prodcutcatid&level1child3_id&","
			end if			
end if	
'*****************Third level*********************************
'*****************Forth level*********************************
if level1child3_id<>"" then
'	response.write "4:<br>"
		if rs6_new.state=1 then rs6_new.close
		sql4="select categoryid from category where pid in ("&level1child3_id&")"
'		response.write sql4&"<br>"
		rs6_new.open sql4,con,1,3
'		response.write rs6_new.recordcount&"<br>"
			if not rs6_new.eof then
				while not rs6_new.eof
					child3=child3&rs6_new("categoryid")&","
				rs6_new.movenext
				wend
level1child4_id=LEFT(child3,Len(child3)-1)
'response.write level1child4_id&"<br>"
'response.write child3&"<br>"
prodcutcatid=prodcutcatid&level1child4_id
end if			
end if	
'*****************Forth level*********************************

'*****************Fifth level*********************************
if level1child4_id<>"" then
'	response.write "5:<br>"
		if rs7_new.state=1 then rs7_new.close
		sql5="select categoryid from category where pid in ("&level1child4_id&")"
'		response.write sql5&"<br>"
		rs7_new.open sql5,con,1,3
'		response.write rs7_new.recordcount&"<br>"
			if not rs7_new.eof then
				while not rs7_new.eof
					child4=child4&rs7_new("categoryid")&","
				rs7_new.movenext
				wend
level1child5_id=LEFT(child4,Len(child4)-1)
'response.write level1child5_id&"<br>"
'response.write child4&"<br>"
prodcutcatid=prodcutcatid&level1child5_id
end if			
end if	
'*****************Fifth level*********************************



'response.write "********************************<br>"
				prodcutcatid_id=LEFT(prodcutcatid,Len(prodcutcatid)-1)
	
'response.write prodcutcatid_id&"<br>"
'response.write "********************************<br>"
sqlgen1_new="select * from product where categoryid in ("&prodcutcatid_id&","&id&")and disable='0' and featuredproduct='1' order by productname"
'sqlgen1=prodcutcatid_id
'response.write "********************************<br>"
'response.write sqlgen1&"<br>"
sqlcatid=sqlgen1_new
'sqlcatid=prodcutcatid_id
'response.write "********************************<br>"

'response.end
end Function

Function SendEmailNew(toEmail, FromEmail, strSubject, strEmailBody)
	 Const cdoSendUsingPickup = 1
        Const cdoSendUsingPort = 2 'Must use this to use Delivery Notification
        Const cdoAnonymous = 0
        Const cdoBasic = 1 ' clear text
        Const cdoNTLM = 2 'NTLM
        'Delivery Status Notifications
        Const cdoDSNDefault = 0 'None
        Const cdoDSNNever = 1 'None
        Const cdoDSNFailure = 2 'Failure
        Const cdoDSNSuccess = 4 'Success
        Const cdoDSNDelay = 8 'Delay
        Const cdoDSNSuccessFailOrDelay = 14 'Success, failure or delay

         set objMsg = CreateObject("CDO.Message")
        set objConf = CreateObject("CDO.Configuration")

        Set objFlds = objConf.Fields
        With objFlds
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.ofizkart.com"
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "customerservice@ofizkart.com"
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Welcome!321"
            .Update
        End With

        With objMsg
            Set .Configuration = objConf
            .To = toEmail
            .BCC = "singh.guru@gmail.com, gamimanoj@gmail.com, right_stationery@yahoo.co.in, sanjeevbhasker@gmail.com"
            .From = FromEmail
            .Subject = strSubject
            .HTMLBody = strEmailBody
            .Fields.update
            .Send
        End With


End Function

%>