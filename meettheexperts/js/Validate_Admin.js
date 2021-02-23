//XML HTTP REQUEST
function getCommonHTTPResponse(url) {
    
    var currentTime = new Date();
    var timespan = currentTime.getMinutes()+"a"+currentTime.getMilliseconds();
    url = url + '&_dc=' + timespan;
    var xmlHTTP;   
    if (window.XMLHttpRequest) {              
    xmlHTTP = new XMLHttpRequest();              
    } else {                                  
    xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
    }
    
    if (xmlHTTP) {
         xmlHTTP.open("GET", url, false);                            
         xmlHTTP.send(null);
         return xmlHTTP.responseText;                                        
    } 
    else {
     return "";
    }                                            
}
//XML HTTP REQUEST POST
function POST_HTTPResponse(url, Params) {
    var xmlHTTP;   
    if (window.XMLHttpRequest) {              
    xmlHTTP = new XMLHttpRequest();              
    } else {                                  
    xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
    }
    
    if (xmlHTTP) {
         xmlHTTP.open("POST", url, false);     
         xmlHTTP.setRequestHeader('Content-Type','application/x-www-form-urlencoded');                       
         xmlHTTP.send(Params);
         return xmlHTTP.responseText;                                        
    } 
    else {
     return "";
    }                                            
}



function validate_admin()
{
	if (trim(document.form1.txtlogin.value)=="")
	{
			alert("Enter Login ID")
			document.form1.txtlogin.value="";
			document.form1.txtlogin.focus();
			return false;
	}
	/*var loginName=document.form1.txtlogin.value.toLowerCase();
	if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz0123456789.-_" ))
		{
			alert("User name has invalid characters");
			document.form1.txtlogin.focus();
			return false;
		}*/
	if (trim(document.form1.txtpass.value)=="")
	{
			alert("Enter Password")
			document.form1.txtpass.value="";
			document.form1.txtpass.focus();
			return false;
	}
	
}
function validate_Editadmin()
{
	if (trim(document.form1.txtlogin.value)=="")
	{
			alert("Enter Login ID")
			document.form1.txtlogin.value="";
			document.form1.txtlogin.focus();
			return false;
	}
	/*var loginName=document.form1.txtlogin.value.toLowerCase();
	if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz0123456789.-_" ))
		{
			alert("User name has invalid characters");
			document.form1.txtlogin.focus();
			return false;
		}*/
	if (trim(document.form1.txtnewpass.value)=="")
	{
			alert("Enter Password.")
			document.form1.txtnewpass.value="";
			document.form1.txtnewpass.focus();
			return false;
	}
	/*if (trim(document.form1.txtcpassword.value)=="")
	{
			alert("Enter Confirm Password")
			document.form1.txtcpassword.value="";
			document.form1.txtcpassword.focus();
			return false;
	}
	if (document.form1.txtnewpass.value != document.form1.txtcpassword.value)
	{
			alert("Password does not match")
			document.form1.txtcpassword.value="";
			document.form1.txtcpassword.focus();
			return false;
	}*/
		
	
}
function validate_AddJob()
{
	if (trim(document.form1.txtjob.value)=="")
	{
			alert("Enter Job Category")
			document.form1.txtjob.value="";
			document.form1.txtjob.focus();
			return false;
	}
	/*var loginName=document.form1.txtjob.value.toLowerCase();
	if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz ',-()" ))
		{
			alert("Job category has invalid characters");
			document.form1.txtjob.focus();
			return false;
		}*/
	
}
function validate_news()
{
	if (trim(document.form1.txtsubject.value)=="")
	{
			alert("Enter News Subject")
			document.form1.txtsubject.value="";
			document.form1.txtsubject.focus();
			return false;
	}
	if (trim(document.form1.txtdt1.value)=="")
	{
			alert("Enter News Date")
			document.form1.txtdt1.value="";
			document.form1.txtdt1.focus();
			return false;
	}
	
}

function validate_adminDefault()
{
	if (trim(document.form1.txtusername.value)=="")
	{
			alert("Enter Login ID")
			document.form1.txtusername.value="";
			document.form1.txtusername.focus();
			return false;
	}
	if (trim(document.form1.txtpass.value)=="")
	{
			alert("Enter Password")
			document.form1.txtpass.value="";
			document.form1.txtpass.focus();
			return false;
	}
}
function validate_memberAdd()
{
	if (trim(document.form1.txtfname.value)=="")
	{
			alert("Enter First Name")
			document.form1.txtfname.value="";
			document.form1.txtfname.focus();
			return false;
	}
	/*var loginName=document.form1.txtfname.value.toLowerCase();
	if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz' " ))
		{
			alert("First Name has invalid characters");
			document.form1.txtfname.focus();
			return false;
		}*/
	if (trim(document.form1.txtlname.value)=="")
	{
			alert("Enter Last Name")
			document.form1.txtlname.value="";
			document.form1.txtlname.focus();
			return false;
	}
	var loginName=document.form1.txtlname.value.toLowerCase();
	if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz' " ))
		{
			alert("Last Name has invalid characters");
			document.form1.txtlname.focus();
			return false;
		}
	if (isvalidemail(trim(document.form1.txtemail.value))==false)
		{
			return false;
		}
	if (document.form1.txtphone.value=="")
		{
			alert("Enter Phone")
			document.form1.txtphone.focus();
			return false;
		}
	if (isNaN(document.form1.txtphone.value))
		{
			alert("Enter valid Phone")
			document.form1.txtphone.focus();
			return false;	
		}
	if (trim(document.form1.txtusername.value)=="")
	{
			alert("Enter Username")
			document.form1.txtusername.value="";
			document.form1.txtusername.focus();
			return false;
	}
	var loginName=document.form1.txtusername.value.toLowerCase();
	if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz0123456789_" ))
		{
			alert("User Name has invalid characters");
			document.form1.txtusername.focus();
			return false;
		}
	if (trim(document.form1.txtpassword.value)=="")
	{
			alert("Enter Password")
			document.form1.txtpassword.value="";
			document.form1.txtpassword.focus();
			return false;
	}
	if (trim(document.form1.txtadd.value)=="")
	{
			alert("Enter Address")
			document.form1.txtadd.value="";
			document.form1.txtadd.focus();
			return false;
	}
	if (trim(document.form1.txtcity.value)=="")
	{
			alert("Enter City")
			document.form1.txtcity.value="";
			document.form1.txtcity.focus();
			return false;
	}
	if(document.form1.lstcountry.selectedIndex==1)
	{
		if (document.form1.cmbstate.selectedIndex==0)
			{
				alert("Select State")
				document.form1.cmbstate.focus();
				return false;
			}
	}	
	else
	{
		if (document.form1.txtotherstate.value=="")
		{
			alert("Enter State")
			document.form1.txtotherstate.focus();
			return false;
		}
	}
	if (trim(document.form1.txtzip.value)=="")
	{
			alert("Enter Zip")
			document.form1.txtzip.value="";
			document.form1.txtzip.focus();
			return false;
	}
	var Selected = false;
	var StdLdTime="membership";	
	for (k = 0; k < document.form1[StdLdTime].length; k++)
		{
			if (document.form1[StdLdTime][k].checked)
			{
				Selected = true;
				if (document.form1[StdLdTime][1].checked)
					{
						
						if (document.form1.txtusername.value ==  document.form1.txtuser1.value || document.form1.txtusername.value ==  document.form1.txtuser2.value || document.form1.txtusername.value ==  document.form1.txtuser3.value || document.form1.txtusername.value ==  document.form1.txtuser4.value)
								{
									alert("Please select other username.")
									return false
								}
							if (document.form1.txtuser1.value ==  document.form1.txtusername.value || document.form1.txtuser1.value ==  document.form1.txtuser2.value || document.form1.txtuser1.value ==  document.form1.txtuser3.value || document.form1.txtuser1.value ==  document.form1.txtuser4.value)
								{
									alert("Please select other username.")
									return false
								}
							if (document.form1.txtuser2.value ==  document.form1.txtusername.value || document.form1.txtuser2.value ==  document.form1.txtuser1.value || document.form1.txtuser2.value ==  document.form1.txtuser3.value || document.form1.txtuser2.value ==  document.form1.txtuser4.value)
								{
									alert("Please select other username.")
									return false
								}	
							if (document.form1.txtuser3.value ==  document.form1.txtusername.value || document.form1.txtuser3.value ==  document.form1.txtuser2.value || document.form1.txtuser3.value ==  document.form1.txtuser1.value || document.form1.txtuser3.value ==  document.form1.txtuser4.value)
								{
									alert("Please select other username.")
									return false
								}
							if (document.form1.txtuser4.value ==  document.form1.txtusername.value || document.form1.txtuser4.value ==  document.form1.txtuser2.value || document.form1.txtuser4.value ==  document.form1.txtuser1.value || document.form1.txtuser4.value ==  document.form1.txtuser3.value)
								{
									alert("Please select other username.")
									return false
								}
						
						
						for (i=1;i<=4;i++)
							{
								var username="txtuser"+i;
								var pass="txtpass"+i;
								var EmailOption1="EmailOption"+i;
								if (trim(document.form1[username].value)=="")
									{
										alert("Enter Username")
										document.form1[username].value="";
										document.form1[username].focus();
										return false;
									}
								var loginName=document.form1[username].value.toLowerCase();
								if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz0123456789_" ))
									{
										alert("User Name has invalid characters");
										document.form1[username].focus();
										return false;
									}
								if (trim(document.form1[pass].value)=="")
									{
										alert("Enter Password")
										document.form1[pass].value="";
										document.form1[pass].focus();
										return false;
									}
									//alert(EmailOption1)
									var emailselected=false
								for (j = 0; j < document.form1[EmailOption1].length; j++)
									{
										if (document.form1[EmailOption1][j].checked)
										{
											emailselected=true;
										}
									}
								if (!emailselected)	
									{
										alert("Select LDS Email Option \n");
										return (false);
									}		
										
							}  // end of internal for loop
					}
				
			}
		}
		if (!Selected)
		{
			alert("Select Membership Plan \n");
			return (false);
		}
	var Selected = false;
	var StdLdTime="EmailOption";	
	for (k = 0; k < document.form1[StdLdTime].length; k++)
		{
			if (document.form1[StdLdTime][k].checked)
			{
				Selected = true;
				
			}
		}
		if (!Selected)
		{
			alert("Select LDS email Option");
			return (false);
		} 	 
					
}


function validate_memberedit()
{
	if (trim(document.form1.txtfname.value)=="")
	{
			alert("Enter First Name")
			document.form1.txtfname.value="";
			document.form1.txtfname.focus();
			return false;
	}
	var loginName=document.form1.txtfname.value.toLowerCase();
	if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz' " ))
		{
			alert("First Name has invalid characters");
			document.form1.txtfname.focus();
			return false;
		}
	if (trim(document.form1.txtlname.value)=="")
	{
			alert("Enter Last Name")
			document.form1.txtlname.value="";
			document.form1.txtlname.focus();
			return false;
	}
	var loginName=document.form1.txtlname.value.toLowerCase();
	if (!isCharsInBag( loginName, "abcdefghijklmnopqrstuvwxyz' " ))
		{
			alert("Last Name has invalid characters");
			document.form1.txtlname.focus();
			return false;
		}
	if (isvalidemail(trim(document.form1.txtemail.value))==false)
		{
			return false;
		}
	if (document.form1.txtphone.value=="")
		{
			alert("Enter Phone")
			document.form1.txtphone.focus();
			return false;
		}
	if (isNaN(document.form1.txtphone.value))
		{
			alert("Enter valid Phone")
			document.form1.txtphone.focus();
			return false;	
		}
	if (trim(document.form1.txtpassword.value)=="")
	{
			alert("Enter Password")
			document.form1.txtpassword.value="";
			document.form1.txtpassword.focus();
			return false;
	}
	if (trim(document.form1.txtadd.value)=="")
	{
			alert("Enter Address")
			document.form1.txtadd.value="";
			document.form1.txtadd.focus();
			return false;
	}
	if (trim(document.form1.txtcity.value)=="")
	{
			alert("Enter City")
			document.form1.txtcity.value="";
			document.form1.txtcity.focus();
			return false;
	}
	if(document.form1.lstcountry.selectedIndex==1)
	{
		if (document.form1.cmbstate.selectedIndex==0)
			{
				alert("Select State")
				document.form1.cmbstate.focus();
				return false;
			}
	}	
	else
	{
		if (document.form1.txtotherstate.value=="")
		{
			alert("Enter State")
			document.form1.txtotherstate.focus();
			return false;
		}
	}
	if (trim(document.form1.txtzip.value)=="")
	{
			alert("Enter Zip")
			document.form1.txtzip.value="";
			document.form1.txtzip.focus();
			return false;
	}
}

function validate_selectMember()
{
	if (document.form1.DropDownList1.value=="Select")
		{
			alert("Select Member")
			document.form1.DropDownList1.focus();
			return false;
		}
	
}

function validateAddProfile()
{
	if (trim(document.form1.txtheadline.value)=="")
	{
			alert("Enter Headline")
			document.form1.txtheadline.value="";
			document.form1.txtheadline.focus();
			return false;
	}
	if (document.form1.cmbheight.value==0)
	{
			alert("Select Height")
			document.form1.cmbheight.focus();
			return false;
	}
	if (document.form1.cmbbuild.value==0)
	{
			alert("Select Build")
			document.form1.cmbbuild.focus();
			return false;
	}
	if (document.form1.cmbhair.value==0)
	{
			alert("Select Hair")
			document.form1.cmbhair.focus();
			return false;
	}
	if (document.form1.cmbeye.value==0)
	{
			alert("Select Eyes")
			document.form1.cmbeye.focus();
			return false;
	}
	if (document.form1.cmbme.value==0)
	{
			alert("Select Criteria")
			document.form1.cmbme.focus();
			return false;
	}
	/*if (document.form1.cmblooking.value==0)
	{
			alert("Select You are looking for")
			document.form1.cmblooking.focus();
			return false;
	}*/
	if (document.form1.cmbstatus.value==0)
	{
			alert("Select Status")
			document.form1.cmbstatus.focus();
			return false;
	}
	if (document.form1.cmbtemple.value==0)
	{
			alert("Select Temple Status")
			document.form1.cmbtemple.focus();
			return false;
	}
	if (document.form1.cmbchild.value=="")
	{
			alert("Select Children")
			document.form1.cmbchild.focus();
			return false;
	}
	if (document.form1.cmbchildathome.value=="")
	{
			alert("Select Children at Home")
			document.form1.cmbchildathome.focus();
			return false;
	}
	if (document.form1.cmbchurch.value==0)
	{
			alert("Select Church activity")
			document.form1.cmbchurch.focus();
			return false;
	}
	if (document.form1.cmbldsmission.value==0)
	{
			alert("Select LDS Mission")
			document.form1.cmbldsmission .focus();
			return false;
	}
	if (document.form1.cmbeducation.value==0)
	{
			alert("Select Education")
			document.form1.cmbeducation .focus();
			return false;
	}
	if (document.form1.cmbethnicity.value==0)
	{
			alert("Select Ethnicity")
			document.form1.cmbethnicity .focus();
			return false;
	}
	if(document.form1.cmbmonth.selectedIndex==0)
		{
			alert("Select Month")
			document.form1.cmbmonth.focus();
			return false;
		}
	if(document.form1.cmbdate.selectedIndex==0)
		{
			alert("Select Date")
			document.form1.cmbdate.focus();
			return false;
		}
	if(document.form1.cmbyear.selectedIndex==0)
		{
			alert("Select Year")
			document.form1.cmbyear.focus();
			return false;
		}
	if (document.form1.cmbmonth.selectedIndex==2 && document.form1.cmbdate.selectedIndex > daysInFebruary(document.form1.cmbyear.selectedIndex)) 	
	{
		alert("Enter a valid date")
		document.form1.cmbdate.focus();
		return false;
	}
	var daysInMonth = DaysArray(12)
	if (document.form1.cmbdate.selectedIndex > daysInMonth[document.form1.cmbmonth.selectedIndex])
	{
		alert("Enter a valid date")
		document.form1.cmbdate.focus();
		return false;
	}	
	if (trim(document.form1.txtoccup.value)=="")
	{
			alert("Enter Occupation")
			document.form1.txtoccup.value="";
			document.form1.txtoccup.focus();
			return false;
	}
	if (trim(document.form1.txtme.value)=="")
	{
			alert("Enter About Me")
			document.form1.txtme.value="";
			document.form1.txtme.focus();
			return false;
	}
	str=document.form1.txtme.value;
	if(str.length>=2000)
		{
		alert("Content should not exceeded by 2000 characters");
		document.form1.txtme.focus();
		return false;
	  } 
	  
	if (trim(document.form1.txtmymatch.value)=="")
	{
			alert("Enter About My Match")
			document.form1.txtmymatch.value="";
			document.form1.txtmymatch.focus();
			return false;
	}
	str=document.form1.txtmymatch.value;
	if(str.length>=2000)
		{
		alert("Content should not exceeded by 2000 characters");
		document.form1.txtmymatch.focus();
		return false;
	  }
	//cmbeducation
	
}








//------  COMMON FUNCTION-----------------------
function trim(str)
{
   return str.replace(/^\s*|\s*$/g,"");
}
function isCharsInBag (s, bag)
{
  var i;
  for (i = 0; i < s.length; i++)
  {
          var c = s.charAt(i);
          if (bag.indexOf(c) == -1) return false;
  }
  return true;
}
function isvalidemail(str)
{
		var AtTheRate= str.indexOf("@");
	    var DotSap= str.lastIndexOf(".");
		if (AtTheRate==-1 || DotSap ==-1)
		{
			alert("Enter Valid Email Address");
			document.form1.txtemail.focus();
			document.form1.txtemail.select();
			return false;
		}
		else
		{
			if( AtTheRate > DotSap )
			{
			alert("Enter Valid Email Address");
			document.form1.txtemail.focus();
			document.form1.txtemail.select();
			return false;
			}
		}

}

function isvalidemailoBJ(str1)
{
        var str = str1.value;
		var AtTheRate= str.indexOf("@");
	    var DotSap= str.lastIndexOf(".");
		if (AtTheRate==-1 || DotSap ==-1)
		{
			alert("Enter Valid Email Address");
			str1.focus();
			str1.select();
			return false;
		}
		else
		{
			if( AtTheRate > DotSap )
			{
			alert("Enter Valid Email Address");
			str1.focus();
			str1.select();
			return false;
			}
		}

}
function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this
}
//--------------------------------------------------
