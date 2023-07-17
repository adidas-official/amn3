# connect urls to views in dragonsdenn/urls.py
from django.urls import path

from . import views

urlpatterns = [
    path('', views.index, name='home'),
    path("run-enforcer/", views.run_enforcer, name="run_enforcer"),
    path("run-inspector/", views.run_inspector, name="run_inspector"),
    path("run-vanguard/", views.run_vanguard, name="run_vanguard"),
]