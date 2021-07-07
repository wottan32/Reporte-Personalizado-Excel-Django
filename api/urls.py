# api/urls.py
from django.urls import path
from .views import DatosAPIView

urlpatterns = [
    path('', DatosAPIView.as_view()),
]