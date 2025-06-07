from django.urls import path
from .views import generate_tickets

urlpatterns = [
    path('', generate_tickets, name='generate_tickets'),
path('', generate_tickets, name='home'),
]
