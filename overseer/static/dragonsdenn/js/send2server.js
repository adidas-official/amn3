$(document).ready(function() {
    $("#run-enforcer").click(function(e) {
        e.preventDefault();

        var csrftoken = $('[name=csrfmiddlewaretoken]').val();
        sendData(csrftoken);
    });
    
});

async function sendData(csrf_token) {

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
        console.log(data);
    } catch (error) {
        console.error(error);
    }
}