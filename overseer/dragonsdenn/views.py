from django.shortcuts import render
from django.http import JsonResponse
from . import enforcer, inspector

def index(request):
    return render(request, 'dragonsdenn/home.html', {'title': 'Home'})


def run_enforcer(request):

    if request.method == 'POST':
        # get the data from the request
        wages_data = request.FILES.get('wages').read().decode('cp1250')
        employees_data = request.FILES.get('employees').read().decode('cp1250')

        # Call run_enforcer function with dataframes as arguments
        result = enforcer.main(wages_data, employees_data)

        # Return result as JSON response
        response_data = {'result': result}
        return JsonResponse(response_data)
    else:
        response_data = {"status": "error", "message": "Invalid request method"}
        return JsonResponse(response_data, status=405)


def run_inspector(request):
    if request.method == 'POST':
        inspector.main()
        # return a JSON response to the front-end
        response_data = {'message': 'Enforcer has run successfully!'}
        return JsonResponse(response_data)