#imported required assemblies
Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;
#Add-Type -AssemblyName Microsoft.VisualBasic;

#declare global variables
#alphabet array for looping
$arrlett = @("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z");
#$arrlett = @("A");

#declare functions
#function to get emails for ship
function GetShip {

    #loop through the letters
    <#foreach($let in $arrlett){
        #loop through the com objects until you get the internet explorer window
        #if not found through an error
        foreach($objwindow in $objapp.Windows()){
            Write-Output $objwindow.Name();
            write-output $ie.ReadyState;
        }
    }#>

    #declare file to write emails to
    $fyle = (Get-Location).Path + "\" + $args[0] + ".txt";

    #the site is divided into frames
    #the first frame is where the search selections go
    #the second frame is where the search results are stored
    $fram = $ie.Document.getElementsByTagName("frame")[0].contentDocument();

    #select reference
    #https://stackoverflow.com/questions/33558807/powershell-internet-explorer-com-object-select-class-drop-down-menu-item
    #to interact with the different elements of the html you have to loop through them because
    #the developers didn't add IDs to the tags, so you can't just reference them
    #to interact with the drop down menus you have to loop through the select elements
    foreach($inp in $fram.getElementsByTagName("select")){
        #Write-host $inp.Name
        #Write-host $inp.Value
        #only interact with the specific drop down menus we want
        #change their values to the correct index
        if ($inp.Name -eq "LNAME_CONSTRAINT"){
            $inp.Options.SelectedIndex = 2;
            #Write-host $inp.Name
            #Write-host $inp.Value
        }
        if ($inp.Name -eq "EMAIL_CONSTRAINT"){
            $inp.Options.SelectedIndex = 1;
            #Write-host $inp.Name
            #Write-host $inp.Value
        }
    }
    #the same goes for this loop, there are no IDs added so all the input elements have to be looped through
    foreach($inp in $fram.getElementsByTagName("input")){
        #only interact with the email input
        if ($inp.Name -eq "FORM_PARAM_EMAIL"){
            $inp.value = $args[0];
            #Write-host $inp.Name
            #Write-host $inp.Value
        }
    }
    #after the unchanging values are edited, loop through the alphabet
    #because a search is limited to 500 results
    foreach($let in $arrlett){
        #loop through the input elements and change the last name value to the current letter
        foreach($inp in $fram.getElementsByTagName("input")){
            if ($inp.Name -eq "FORM_PARAM_LNAME"){
                $inp.value = $let;
                #Write-host $inp.Name
                #Write-host $inp.Value
            }
            #because the search input element is after the last name element
            #it can be safely clicked after the last name is changed
            if ($inp.Name -eq "Search"){
                #click button reference
                #https://www.reddit.com/r/PowerShell/comments/4ehrxu/clicking_an_ie_button_with_powershell/
                $inp.click();
            }
        }
        #now grab the results in the second frame and write them to a file
        #wait for the page to be ready, just in case it's taking awhile to load
        while($ie.ReadyState -ne 4 -Or $ie.Busy -eq $true){
            start-sleep -m 100;
            #write-output $ie.ReadyState;
        }
        $frame = $ie.Document.getElementsByTagName("frame")[1].contentDocument();
        #set check variable
        $check = " ";
        #check if the results are empty
        foreach($p in $frame.getElementsByTagName("p")){
            if($p.innerText -eq "No match found."){
            #if($p.innerText -eq "No entries match the requested search term. Please try a different search."){
                #set check value to empty
                $check = "";
                #break out of while loop if empty
                #break reference
                #https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_break?view=powershell-7.2
                break;
            }
        }
        #if not empty grab the results and write it to a file
        if ($check){
            #string for results that include more than the email
            $allResults = "";
            $endStr = "";
            #loop for next click pages
            :nextClick while ($true){
                foreach($inp in $frame.getElementsByTagName("input")){
                    #check if second argument exists, to just get emails
                    if($args[1]){
                        #check for email value
                        if($inp.Name -eq "FORM_PARAM_EMAIL"){
                            #write email to file
                            Add-Content -Path $fyle -Value $inp.Value;
                        }
                        #check if the loop is at the next button
                        if($inp.value -eq "  Next >>  "){
                            #check if the button is disabled
                            if($inp.disabled){
                                #exit loop
                                break nextClick;
                            } else {
                                #click next
                                $inp.click();
                                #wait for the page to be ready, just in case it's taking awhile to load
                                while($ie.ReadyState -ne 4 -Or $ie.Busy -eq $true){
                                    start-sleep -m 100;
                                    #write-output $ie.ReadyState;
                                }
                                #regrab the second frame after it is loaded
                                $frame = $ie.Document.getElementsByTagName("frame")[1].contentDocument();
                            }
                        }
                    } else {
                        #get almost all the inputs
                        #FORM_PARAM_DN,FORM_PARAM_LNAME,FORM_PARAM_FNAME,FORM_PARAM_MIDDLE_INITIAL
                        #FORM_PARAM_EMAIL,FORM_PARAM_LDAP_MODIFYTIMESTAMP
                        #don't get following inputs
                        #FORM_PARAM_SUFFIX,REQ_ACTION
                        #reference for multiline if statement
                        #https://stackoverflow.com/questions/36689644/how-to-split-an-if-condition-over-multiline-lines-with-comments
                        <#if (
                            $inp.Name -eq "FORM_PARAM_DN" -Or $inp.Name -eq "FORM_PARAM_LNAME" -Or
                            $inp.Name -eq "FORM_PARAM_FNAME" -Or $inp.Name -eq "FORM_PARAM_MIDDLE_INITIAL" -Or 
                            $inp.Name -eq "FORM_PARAM_EMAIL" -Or $inp.Name -eq "FORM_PARAM_LDAP_MODIFYTIMESTAMP"
                            ){
                            write-host $inp.name;
                            Write-host $inp.value;
                        }#>
                        #reference for switch
                        #https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_switch?view=powershell-7.2
                        switch ($inp.Name) {
                            "FORM_PARAM_DN"{$endStr = $inp.value; Break}
                            "FORM_PARAM_LNAME"{$allResults = $inp.value; Break}
                            "FORM_PARAM_FNAME"{$allResults = $allResults + " " + $inp.value; Break}
                            "FORM_PARAM_MIDDLE_INITIAL"{$allResults = $allResults + " " + $inp.value; Break}
                            "FORM_PARAM_EMAIL"{$allResults = $inp.value + " " + $allResults; Break}
                            "FORM_PARAM_LDAP_MODIFYTIMESTAMP"{$allResults = $allResults + " " + $inp.value; Break}
                            "REQ_ACTION"{$line = $allResults + $endStr; Add-Content -Path $fyle -Value $line; $allResults = ""; $endStr = ""; Break}
                        }
                        #check if the loop is at the next button
                        if($inp.value -eq "  Next >>  "){
                            #check if the button is disabled
                            if($inp.disabled){
                                #exit loop
                                break nextClick;
                            } else {
                                #click next
                                $inp.click();
                                #wait for the page to be ready, just in case it's taking awhile to load
                                while($ie.ReadyState -ne 4 -Or $ie.Busy -eq $true){
                                    start-sleep -m 100;
                                    #write-output $ie.ReadyState;
                                }
                                #regrab the second frame after it is loaded
                                $frame = $ie.Document.getElementsByTagName("frame")[1].contentDocument();
                            }
                        }
                    }
                }
            }
        }
    }
}

#declare com object that can attach to internet explorer
#$objapp = new-object -comobject "Shell.Application";

#declare com object to prompt user to login
#$sh = new-object -ComObject "WScript.Shell";

$ie = new-object -com "InternetExplorer.Application";
$ie.navigate("https://dod411.gds.disa.mil");
#$ie.visible = $true;
#write-output $ie.ReadyState;
#$ie | get-member;
#$objapp.Windows() | get-member;
#$objapp.Windows().Name

#declare empty variable to store internet explorer object when found
#$objie;

#get the full path of the out file
#$outfile = (pwd).Path + "\" + $out;

#prompt user to continue
#1 = OK
#$sh.Popup("Press OK after logging in to the website with your CAC and accepting the warning", 0, "Please Log In", 0);
#enums for MsgBox
#https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.interaction.msgbox?view=net-6.0
#$loggedin = [Microsoft.VisualBasic.Interaction]::MsgBox("Press OK after logging in to the website with your CAC and accepting the warning", 1, "Please Log In");

#check if internet explorer is ready
#enums for readystates:
#https://docs.microsoft.com/en-us/previous-versions/bb268229(v=vs.85)
#write-output $ie.ReadyState;
while($ie.ReadyState -ne 4 -Or $ie.Busy -eq $true){
    start-sleep -m 1000;
    #write-output $ie.ReadyState;
}
#Write-Output ($ie.Document.IHTMLDocument2_scripts | Get-Member);
#$ie | get-member;
#prompt user for ship simple old way
$ship = " ";
#while($ship){
    #prompt user for ship hull number with simple form
    #$ship = [Microsoft.VisualBasic.Interaction]::InputBox("Ship Hull Number", "Enter ship's hull number (i.e., cvn75, lhd7): ");
    #if ($ship){
        #GetShip $ship "true"
    #}
#}
#prompt user for ship new way
$result = "OK";
while($result -eq "OK" -And $ship){
    #custom form to include checkbox
    #custom form reference
    #https://docs.microsoft.com/en-us/powershell/scripting/samples/creating-a-custom-input-box?view=powershell-7.2
    $form = New-Object System.Windows.Forms.Form;
    $form.Text = "Ship Hull Number(s)";
    $form.Size = New-Object System.Drawing.Size(300,200);
    $form.StartPosition = "CenterScreen";
    #add ok button
    $okBut = New-Object System.Windows.Forms.Button;
    $okBut.Location = New-Object System.Drawing.Point(75,120);
    $okBut.Size = New-Object System.Drawing.Size(75,23);
    $okBut.Text = "OK";
    $okBut.DialogResult = [System.Windows.Forms.DialogResult]::OK;
    $form.AcceptButton = $okBut;
    $form.Controls.Add($okBut);
    #add cancel button
    $cancelBut = New-Object System.Windows.Forms.Button;
    $cancelBut.Location = New-Object System.Drawing.Point(150,120);
    $cancelBut.Size = New-Object System.Drawing.Size(75,23);
    $cancelBut.Text = "Cancel";
    $cancelBut.DialogResult = [System.Windows.Forms.DialogResult]::Cancel;
    $form.CancelButton = $cancelBut;
    $form.Controls.Add($cancelBut);
    #add label
    $label = New-Object System.Windows.Forms.Label;
    $label.Location = New-Object System.Drawing.Point(10,20);
    $label.Size = New-Object System.Drawing.Size(280,20);
    $label.Text = "Enter ship's hull number (i.e., cvn75, lhd7): ";
    $form.Controls.Add($label);
    #add textbox
    $textbox = New-Object System.Windows.Forms.TextBox;
    $textbox.Location = New-Object System.Drawing.Point(10,40);
    $textbox.Size = New-Object System.Drawing.Size(260,20);
    $form.Controls.Add($textbox);
    #add checkbox
    #checkbox reference
    #http://serverfixes.com/powershell-checkboxes
    $checkbox = New-Object System.Windows.Forms.CheckBox;
    $checkbox.Location = New-Object System.Drawing.Size(10,60);
    $checkbox.Size = New-Object System.Drawing.Size(280,20);
    $checkbox.Text = "Get email addresses only";
    $checkbox.Checked = $true;
    $form.Controls.Add($checkbox);
    #make the form on top
    $form.Topmost = $true;
    #display form
    $form.Add_Shown({$textbox.Select()});
    $form.MaximizeBox = $false;
    $result = $form.ShowDialog();
    #Write-Host $result;
    #Write-Output $ship;
    #call function with ship name if ship is not empty
    #cancel button and nothing in the input box will make ship empty
    if($result -eq "OK"){
        #Write-output "clicked ok";
        #get textbox input and set to ship variable
        $ship = $textbox.Text;
        if ($ship){
            #check if multiple ships were provided with commas
            #else just send the one ship
            if ($ship -Match ","){
                #remove any whitespace and split the ships by comma
                $ships = ($ship -replace "\s","").Split(",");
                #loop through ships array and get the emails
                foreach($shp in $ships){
                    #get emails for ship
                    #check if emails only is checked
                    if($checkbox.Checked){
                        #Write-host "multiple ships with checkbox checked";
                        #emails only is checked so call GetShip with a second argument
                        GetShip $shp "true";
                    } else {
                        #Write-host "multiple ships with checkbox not checked";
                        #emails only is not checked so call GetShip without a second argument
                        GetShip $shp;
                    }
                }
            } else {
                #get emails for ship
                #check if emails only is checked
                if($checkbox.Checked){
                    #Write-host "single ship with checkbox checked";
                    #emails only is checked so call GetShip with a second argument
                    GetShip $ship "true";
                } else {
                    #Write-host "single ship with checkbox not checked";
                    #emails only is not checked so call GetShip without a second argument
                    GetShip $ship;
                }
            }
        }
        #GetShip $ship "true";
    }
    #GetShip $ship $true;
}

<#foreach($objwindow in $objapp.Windows()){
    Write-Output $objwindow.Name();
    write-output $ie.ReadyState;
}#>
#Write-Output $ie.Document;

#close internet explorer at the end
$ie.Quit();