function formValidation()
{
var fname = document.registration.fname;
var mobile = document.registration.mobile;
var lname = document.registration.lname;
var add = document.registration.address;
var state = document.registration.state;
var city = document.registration.city;
var email = document.registration.email;
if(allLetter(fname))
{
if(allnumeric(mobile))
{
if(allLetter(lname))
{ 
if(stateselect(state))
{
if(allLetter(city))
{
if(ValidateEmail(email))
{
window.location="D:\BDD\\projectdetails.html";
}
}
} 
}
}
}
return false;

} 
function allLetter(uname)
{ 
var letters = /^[A-Za-z]+$/;
if(uname.value.match(letters))
{
return true;
}
else
{
alert('Name and City must have alphabet characters only');
uname.focus();
return false;
}
}
function alphanumeric(uadd)
{ 
var letters = /^[0-9a-zA-Z]+$/;
if(uadd.value.match(letters))
{
return true;
}
else
{
alert('User address must have alphanumeric characters only');
uadd.focus();
return false;
}
}
function stateselect(state)
{
if(state.value == "Default")
{
alert('Select your state from the list');
ucountry.focus();
return false;
}
else
{
return true;
}
}
function allnumeric(mobile)
{ 
var numbers = /^[0-9]+$/;
if(mobile.value.match(numbers) && mobile.value.length==10 )
{
return true;
}
else
{
alert('Mobile must have numeric characters only and should of length 10');
uzip.focus();
return false;
}
}
function ValidateEmail(email)
{
var mailformat = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;
if(email.value.match(mailformat))
{
return true;
}
else
{
alert("You have entered an invalid email address!");
email.focus();
return false;
}
} 

