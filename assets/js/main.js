document.addEventListener('DOMContentLoaded', enhanceRequestAbesence(), false);

// Prensentation:
/*
    Manager:  Are you in today?
    Me:  No... I submitted leave in WorkDay
    Manager:  Your calendar doesn'ts how you're out of the office.   I didn't realize you were out
    Me:  You approved my leave...
    Manager: Please update your outlook to reflect this
    Me:  Sorry, signal is  going out...
*/

var selectors = {
    submitButton: "[data-automation-id='toolbarButtonContainer'] [data-uxi-actionbutton-action='bpf-submit']",
    requestButton: "[data-uxi-actionbutton-action='bpf-submit']",
    absenceFormFooter: "footer div.WJKI",
    startDate: "[headers^=columnheader1] .gwt-Label",
    endDate: "[headers^=columnheader2] .gwt-Label",
    leaveType: "[headers^=columnheader3] ul li p",
}

function enhanceRequestAbesence(){

// Wait for the request abense button to appear
var intervalID = setInterval(function() {

    var button = $(selectors.requestButton);
    
    // If it isn't there yet, wait
    if (button != null && button.length > 0) {
        
        console.log(button);

        // Once the user clicks the button, begin waiting for the request leave window to appear
        document.getElementById($(button).attr("id")).addEventListener("click", activateOutlookFields);

        clearInterval(intervalID);
    }
}, 2000);
 
}

function activateOutlookFields() {

    console.log("Handling time off request button.");

    var intervalID = setInterval(function(){

        console.log("Looking fo the form footer");
        var absenceFormFooter = $(selectors.absenceFormFooter);

        // Once the form has appeared add a checkbox for the user to check if they want to sync to Outlook.
        if(absenceFormFooter != null && absenceFormFooter.length > 0){

            console.log(absenceFormFooter);
            console.log("Appending the checkbox");
            var myWidget = $("<br /><div class='WILI'><h3 class='WIVK WGVK'>Outlook Calendar Information</h3><ul><li><label>Start Time:<input id='startTime' type='time' /></label></li><li><label>End Time:<input id='endTime' type='time' /></label></li></ul></div>");
            $(absenceFormFooter).append(myWidget);

            // The new submit button
            document.getElementById($(selectors.submitButton).attr("id")).addEventListener("mouseover", function(){

                sendEventToOutlook(myWidget);
            });

            clearInterval(intervalID);
        }
    }, 1000);

}

function sendEventToOutlook(myWidget) {

    // Get the start and end time from added fields   
    var startTime = $(myWidget).find('#startTime')[0].value;
    var endTime = $(myWidget).find('#endTime')[0].value;

    // Get the date from WD fields
    var startDate = $(selectors.startDate)[0].innerText;
    var endDate = $(selectors.endDate)[0].innerText;

   var summary = $(selectors.leaveType)[0].innerText;
   //var summary = "Hello World";

    console.log("Sending to outlook: " + startDate + " " + startTime + " - " + endDate + " " + endTime);
    console.log("STart datetime: " + createIcalDate(startDate, startTime));

    var file = new Blob(["BEGIN:VCALENDAR\n" + 
    "VERSION:2.0\n" + 
    "BEGIN:VEVENT\n" + 
    "CATEGORIES:APPOINTMENT\n" + 
    "STATUS:BUSY\n" + 
    "DTSTART:" + createIcalDate(startDate, startTime) + "\n" + 
    "DTEND:" + createIcalDate(endDate, endTime) + "\n" + 
    "SUMMARY:" + summary + "\n" + 
    "DESCRIPTION:Copied from Workday Outlook extension.\n" + 
    "CLASS:PUBLIC\n" + 
    "X-MICROSOFT-CDO-BUSYSTATUS:OOF\n" +
    "END:VEVENT\n" + 
    "END:VCALENDAR"], {type: 'text/plain'});
    if (window.navigator.msSaveOrOpenBlob) // IE10+
        window.navigator.msSaveOrOpenBlob(file, 'mytest.ics');
    else { // Others
        var a = document.createElement("a"),
                url = URL.createObjectURL(file);
        a.href = url;
        a.download = 'mytest.ics';
        document.body.appendChild(a);
        a.click();
        setTimeout(function() {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);  
        }, 0); 
    }
}

function createIcalDate(date, time){
    var dateChunks = date.split("/");
    var date = dateChunks[2] + dateChunks[0] + dateChunks[1] + 'T';

    //Add 4 hours for UTC??
    var timeChunks = time.split(":");
    //var time = time.replace(":", "") + "00Z";
    var time = (parseInt(timeChunks[0]) + 4).toString() + timeChunks[1] + "00Z";

    return (date + time);
}

