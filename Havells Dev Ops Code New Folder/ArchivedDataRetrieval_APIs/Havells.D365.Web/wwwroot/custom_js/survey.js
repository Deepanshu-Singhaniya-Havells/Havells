
var Q1Html = '<div class="txt">';
Q1Html += 'We are sorry, May be Know what did you not like about Havells Product &amp; Services.<br>';
Q1Html += '<textarea name="answer" style="width:100%;" class="answer" id="Q1Answer" cols="5" rows="2" placeholder="Enter Your Answer"></textarea>';
Q1Html += '</div>';


var Q2Html = '<div class="txt">';
Q2Html += 'Thank you for the score, May be Know your Suggestion for any improvement in Havells Product &amp; Services.<br>';
Q2Html += '<textarea name="answer" style="width:100%; class="answer" id="Q2Answer" "cols="5" rows="2" placeholder="Enter Your Answer"></textarea>';
Q2Html += '</div>';

var Q3Html = '<div class="txt">';
Q3Html += 'Thank you for the high score, May be Know what did you not like about Havells Product & Services.<br>';
Q3Html += '<textarea name="answer" style="width:100%; class="answer" id="Q3Answer" "cols="5" rows="2" placeholder="Enter Your Answer"></textarea>';
Q3Html += '</div>';

$(document).ready(function () {

    if ($("#status").val() == "True") {
        $("#messageForm").show();
        $("#feedBackForm").hide();
        $("#resultMessage").text(" You Have Already Submitted The Survey");

    }
    else if ($("#status").val() == "False") {
        $("#QuestionSection").html(Q1Html);
        $("#feedBackForm").show();
        $("#messageForm").hide();
    }


});

$("input[name='rating']").change(function () {
    $("#QuestionSection").html('');
    $("#ratingMessage").hide();
    var scoreValue = $('input[name="rating"]:checked').val();
    if (scoreValue < 7) {
        $("#ratingMessage").hide();
        $("#QuestionSection").append(Q1Html);
    }
    else if (scoreValue > 6 && scoreValue < 9) {
        $("#ratingMessage").hide();
        $("#QuestionSection").append(Q2Html);
    }
    else if (scoreValue > 8) {
        $("#ratingMessage").hide();
        $("#QuestionSection").append(Q3Html);
    }

});

$('body').on('click', '#btnSave', function () {

    var res = validate();
    if (res == false) {
        return false;
    }

    var DetractorsResponse = "";
    var PassivesResponse = "";
    var PromotersResponse = "";
    if ($("#Q1Answer").val() != undefined)
        DetractorsResponse = $("#Q1Answer").val();
    if ($("#Q2Answer").val() != undefined)
        PassivesResponse = $("#Q2Answer").val();
    if ($("#Q3Answer").val() != undefined)
        PromotersResponse = $("#Q3Answer").val();


    var survey = {
        JobId: $('#jobId').val(),
        NPSValue: $('input[name="rating"]:checked').val(),
        DetractorsResponse: DetractorsResponse,
        PassivesResponse: PassivesResponse,
        PromotersResponse: PromotersResponse,
        Feedback: $("#feedBack").val(),
        ServiceEngineerRating: $('input[name="feedback"]:checked').val(),
        SubmitStatus: true
    };
    $.ajax({
        url: "/Survey/SubmitFeedback",
        data: JSON.stringify(survey),
        type: "POST",
        contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (result) {
            debugger;
            if (result.result == true) {
                $("#feedBackForm").hide();
                $("#messageForm").show();
                $("#resultMessage").text("Thanks For Sharing Your Feed Back");
            }
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

});

$('body').on('click', '#Q1Answer', function () {
    $("#answerMessage").hide();
});
$('body').on('click', '#Q2Answer', function () {
    $("#answerMessage").hide();
});
$('body').on('click', '#Q3Answer', function () {
    $("#answerMessage").hide();
});

$('body').on('click', '#feedBack', function () {
    $("#feedBackMessage").hide();
});

$("input[name='feedback']").change(function () {
    $("#feedsmileMessage").hide();
});

function validate() {
    var isValid = true;
    var scoreValue = $('input[name="rating"]:checked').val();
    var smileValue = $('input[name="feedback"]:checked').val();

    var answerValue = 0;
    if (scoreValue < 7)
        answerValue = $("#Q1Answer").val();
    else if (scoreValue > 6 && scoreValue < 9)
        answerValue = $("#Q2Answer").val();
    else if (scoreValue > 8)
        answerValue = $("#Q3Answer").val();

    if (scoreValue == undefined) {
        $("#ratingMessage").show();
        isValid = false;
    }
    else {
        $("#ratingMessage").hide();
    }

    if (answerValue == undefined || answerValue == null || answerValue < 0 || answerValue == "") {
        $("#answerMessage").show();
        isValid = false;
    }
    else {
        $("#answerMessage").hide();
    }

    if ($('#feedBack').val().trim() == "") {
        $('#feedBackMessage').show();
        isValid = false;
    }
    else {
        $('#feedBackMessage').hide();
    }
    if (smileValue == undefined) {
        $("#feedsmileMessage").show();
        isValid = false;
    }
    else {
        $("#feedsmileMessage").hide();
    }
    return isValid;
}