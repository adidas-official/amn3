from django.shortcuts import render
from django.http import JsonResponse
from . import enforcer, inspector

def index(request):
    return render(request, 'dragonsdenn/home.html', {'title': 'Home'})

def run_enforcer(request):
    result = enforcer.main()
    return JsonResponse({'result': 'success'})

def run_inspector(request):
    result = inspector.main()
    return JsonResponse({'result': 'success'})