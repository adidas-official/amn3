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
        console.log(data);
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