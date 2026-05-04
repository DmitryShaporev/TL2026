from django.shortcuts import render

# Create your views here.
from django.shortcuts import render

def index(request):
    """Простая страница для теста """
    return render(request, 'lumber_track/index.html')