from django.shortcuts import render
from django.http import JsonResponse
from . import enforcer, inspector

def index(request):
    return render(request, 'dragonsdenn/home.html', {'title': 'Home'})

def run_enforcer(request):
    # get content from element id 'wages'
    if request.method == 'POST':
        wages = request.FILES.get('wages')
        employees = request.FILES.get('employees')
        print(wages)
        print(employees)
        # do something with the files
        enforcer.main(wages, employees)
        # return a JSON response to the front-end
        response_data = {'message': 'Enforcer has run successfully!'}
        return JsonResponse(response_data)

def run_inspector(request):
    if request.method == 'POST':
        wages = request.FILES.get('wages')
        employees = request.FILES.get('employees')
        # do something with the files
        inspector.main(wages, employees)
        # return a JSON response to the front-end
        response_data = {'message': 'Enforcer has run successfully!'}
        return JsonResponse(response_data)