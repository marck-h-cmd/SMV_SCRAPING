from django.http import HttpResponse
from django.template import Template,Context
from django.template.loader import get_template
from django.shortcuts import render



def acceder(request):
    return render(request,"main/index.html")
