function formValidation2()
{
var pname = document.project.pname;
var tsize = document.project.tsize;
var cname = document.project.cname;
if(IsEmpty(pname))
{
if(IsEmpty(cname))
{
if(allnumeric(tsize))
{
alert('Registered Successfully');
}
}
}
return false;

} 


function IsEmpty(pname){ 

                if(pname.value == "")
                {
                alert("project name and client name cannot be empty");
				pname.focus();
				return false
                }
				else
				{
                return true;
				}
            }

function allnumeric(tsize)
{ 
var numbers = /^[0-9]+$/;
if(tsize.value.match(numbers))
{
return true;
}
else
{
alert('Team Size must have numeric characters only ');
tsize.focus();
return false;
}
}


