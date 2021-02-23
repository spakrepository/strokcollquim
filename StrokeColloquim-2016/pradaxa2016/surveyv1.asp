<html>
<body>

<form id="frm1" name="frm1" action="survey.asp?uid=<%=request("uid")%>" method="post">
Q1   <input type="radio" name="q1" value="y"> Yes
  <input type="radio" name="q1" value="n"> No<br><br>
Q2 <input type="text" id="q2" name="q2"> <br><br>
Q3 <input type="radio" name="q3" value="y"> Yes
  <input type="radio" name="q3" value="n"> No<br><br>

<br><br>
<input type="submit" text="Send">
<input type="reset" text="reset">


</form>
</body>
</html>