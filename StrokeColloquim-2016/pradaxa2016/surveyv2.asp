<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- The above 3 meta tags *must* come first in the head; any other head content must come *after* these tags -->
    <title>Georg Nickenig - Survey</title>
    <!-- Bootstrap -->
    <link href="css/bootstrap.min.css" rel="stylesheet">
    <!--Google Fonts-->
    <link href='https://fonts.googleapis.com/css?family=Open+Sans:300' rel='stylesheet' type='text/css'>
    <link href="css/style.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
	<script src="js/Validate_Admin.js"></script>
<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','https://www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-80569436-1', 'auto');
  ga('send', 'pageview');

</script>
	<script>

	function validateMev1(){
	
		if(trim(document.getElementById("name").value) === "") {
			alert("Enter Name");
			document.getElementById("name").focus();
			return false;
		}


		
		if(trim(document.getElementById("email").value) === "") {
			alert("Enter Email Id");
			document.getElementById("email").focus();
			return false;
		}

			if(trim(document.getElementById("phone").value) === "") {
			alert("Enter Phone no.");
			document.getElementById("phone").focus();
			return false;
		}

		if(trim(document.getElementById("speciality").value) === "") {
			alert("Enter Speciality");
			document.getElementById("speciality").focus();
			return false;
		}

		if(trim(document.getElementById("live").value) === "") {
				alert("select Live");
				document.getElementById("live").focus();
				return false;
			}
    
	frm1 = document.getElementById('frm1');

	frm1.action="live-check.asp?city="+document.getElementById("live").selectedIndex;
	frm1.submit();
		
	}


function getParameterByName( name ){
    var regexS = "[\\?&]"+name+"=([^&#]*)", 
  regex = new RegExp( regexS ),
  results = regex.exec( window.location.search );
  if( results == null ){
    return "";
  } else{
    return decodeURIComponent(results[1].replace(/\+/g, " "));
  }
}


	</script>
  </head>
  <body>
   <!--Start Top Bar-->
    <div class="container top-container">
      <div class="row">
          <div class="col-lg-6 col-md-6 col-sm-6 col-xs-12 text-right" id="left-text">
            <p>ANTICOAGULATION CARE 3.0</p>
          </div> <!--End Left Col-->
          
          
            &nbsp;
          <img src="img/lg.png">
          <!--
          <div class="col-lg-6 col-md-6 col-sm-6 col-xs-12" id="right-text">
            <p>Prof. Georg Nickenig <br><span>University of Bonn, Germany</span></p>
          </div>--> <!--End Right Col-->
      </div> <!--End Row-->
    </div> <!--End Container--> 
   <!--End Top Bar-->  
   
   <!--Start Navigation Bar-->
    <div class="container nav-container">
          <nav class="navbar navbar-inverse">
      <div class="container-fluid">
        <div class="navbar-header">
          <button type="button" class="navbar-toggle" data-toggle="collapse" data-target="#myNavbar">
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>                        
          </button>
          <a class="navbar-brand" href="#" style="display:none;">WebSiteName</a>
        </div>
        <div class="collapse navbar-collapse" id="myNavbar">
          <ul class="nav navbar-nav">
            <li><a href="index.html" target="_self">About Prof. Dr. Georg Nickenig</a></li>
            <li class="active"><a href="survey.html" target="_self">Survey</a></li>
            <li><a href="contact-us.html" target="_self">Contact Us</a></li>
          </ul>
        </div>
      </div>
    </nav>
</div> <!--End Container-->
   <!--End Navigation Bar-->
   
   <!--Start Content-->
   <div class="container" id="gradient-background-color">
       <div class="row" id="margin-top">
         <div class="col-lg-5 col-md-5 col-sm-5 col-xs-12 col-lg-offset-1 col-md-offset-1 col-sm-offset-1">
              <h2>Survey</h2>
			  <h2><div style="color:green;font-weight:bold;" id="displaymsg"></div></h2>
             <form id="frm1" name="frm1" action="survey.asp?uid=<%=request("uid")%>" method="post">
                <div class="form-group">
                <p>1. Did this meeting cover all the topics of your interest?</p>
                  <label class="radio-inline">
                      <input type="radio" name="q1" value="y"><span class="survey">Yes</span>
                    </label>
                    <label class="radio-inline">
                      <input type="radio" name="q1" value="n"><span class="survey">No</span>
                    </label>
                </div>
                   <div class="form-group">
                      <p>2. Which other topics pertaining to SPAF management would you like us to include in similar forums in future?</p>
                      <textarea class="form-control" rows="2" id="q2" name="q2"></textarea>
                   </div>
                    
                <div class="form-group">
                <p>3. Were the faculty members well-read and informed on the subject?</p>
                 <label class="radio-inline">
                      <input type="radio" name="q3" value="y"><span class="survey">Yes</span>
                    </label>
                    <label class="radio-inline">
                      <input type="radio" name="q3" value="n"><span class="survey">No</span>
                    </label>
              </div>
                   
                  
                <div class="text-left text-xs-left">
                <button type="submit" class="btn btn-default btn-adjust">Submit</button>
                </div>
              </form>
            </div> <!--End Left Col-->  
            <div class="col-lg-6 col-md-6 col-sm-6 col-xs-12 prof-xs-margin-top">
              <div class="text-right">
                <img src="img/prof.png" class="image-responsive">
              </div>
            </div> <!--End Right Col-->
        </div> <!--End Row-->
</div>
   <!--End Content-->
    <!--Script-->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script src="js/bootstrap.min.js"></script>
	<script>
	//document.getElementById('displaymsg').innerHtml = getParameterByName('msg')
	$('#displaymsg').text(getParameterByName('msg'))
	var cityID = getParameterByName('city')
	if(cityID !== "")
	{
	$('#live')[0].selectedIndex = cityID;
	}
	</script>
  </body>
</html>