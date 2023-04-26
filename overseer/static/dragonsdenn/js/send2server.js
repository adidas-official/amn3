$(document).ready(function() {
    $("#run-enforcer").click(function(e) {
        e.preventDefault();

        var csrftoken = $('[name=csrfmiddlewaretoken]').val();
        enforce(csrftoken);
    });

    $("#run-inspector").click(function(e) {
        e.preventDefault();

        var csrftoken = $('[name=csrfmiddlewaretoken]').val();
        inspect(csrftoken);
    });
    
});

function displayResult(result) {
    var month_keys =['month', 'payout', 'cell'];
    for (var i = 0; i < result.length; i++) {
        console.log(result[i]);
        var name = result[i].name;
        var sheet = result[i].sheet;
        var is_new = result[i].new;
        var pension = result[i].pension;
        var months = result[i].months;
        
        var row = "<div class='row'>";
        row += "<span>" + name + "</span>";
        row += "<span>" + sheet + "</span>";
        row += "<span>" + is_new + "</span>";
        row += "<span>" + pension + "</span>";
        row += "<ul>";
        for (var j = 0; j < months.length; j++) {
            row += "<li>";
            row += months[j].month;
            row += "</li>";
        };
        row += "</ul>";
        row += "<ul>";
        for (var j = 0; j < months.length; j++) {
            row += "<li>";
            row += months[j].payout;
            row += "</li>";
        };
        row += "</ul>";
        row += "<ul>";
        for (var j = 0; j < months.length; j++) {
            row += "<li>";
            row += months[j].cell;
            row += "</li>";
        };
        row += "</ul>";
        row += "</div>";

        $("#result").append(row);
        // console.log(result[i].name);
    };
};

async function enforce(csrf_token) {

    var form = document.getElementById('run-form');
    var formData = new FormData(form);

    // Make the AJAX request
    try {
        const response = await fetch('/run-enforcer/', {
            method: 'POST',
            body: formData,
            headers: {
                'X-CSRFToken': csrf_token
            }
        });
        const data = await response.json();
        displayResult(data['result'][0]);
        // $("#result-text").text(data['result']);
    } catch (error) {
        console.error(error);
    }
}

async function inspect(csrf_token) {
    try {
        console.log('Requesting inspection');
        const response = await fetch('/run-inspector/', {
            method: 'POST',
            headers: {
                'X-CSRFToken': csrf_token
            }
        });
        console.log('Inspection requested!');
    } catch (error) {
        console.log('Inspection failed!');
    }
}
