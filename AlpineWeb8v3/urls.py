"""
Definition of urls for AlpineWeb8v3.
"""

from datetime import datetime
from django.urls import path
from django.contrib import admin
from app import forms, views
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('', views.home, name='home'),
    path('mkquestions/', views.mkquestions, name='mkquestions'),
    path('mkquestions_matrix/', views.mkquestions_matrix, name='mkquestions_matrix'),
    path('subjectchoice/', views.subjectchoice, name='subjectchoice'),
    path('contact/', views.contact, name='contact'),
    path('about/', views.about, name='about'),
    path('unitsstart/', views.unitsstart, name='unitsstart'),
    path('unitspart/', views.unitspart, name='unitspart'),
    path('unitspartviews/', views.unitspartviews, name='unitspartviews'),
    path('unitspartrepair/', views.unitspartrepair, name='unitspartrepair'),
    path('apmajor/', views.apmajor, name='apmajor'),
    path('apmajorviews/', views.apmajorviews, name='apmajorviews'),
    path('apmajorrepair/', views.apmajorrepair, name='apmajorrepair'),
    path('settest/', views.settest, name='settest'),
    path('makeexam/', views.makeexam, name='makeexam'),
    path('createstandard/', views.createstandard, name='createstandard'),
    path('gethoanvi/', views.gethoanvi, name='gethoanvi'),
    path('getmtrde/', views.getmtrde, name='getmtrde'),
    path('hien/', views.hien, name='hien'),
    path('login1/', views.login1, name='login1'),
    path('admin/', admin.site.urls),
    path('admin_dashboard/', views.admin_dashboard, name='admin_dashboard'),
    path('admin_users/', views.admin_users, name='admin_users'),
    path('admin_student/', views.admin_student, name='admin_student'),
    path('admin_sys/', views.admin_sys, name='admin_sys'),
    path('teacher_dashboard/', views.teacher_dashboard, name='teacher_dashboard'),
    path('teacher_lophp/', views.teacher_lophp, name='teacher_lophp'),
    path('teacher_lichthi/', views.teacher_lichthi, name='teacher_lichthi'),
    path('tea_nghg/', views.tea_nghg, name='tea_nghg'),
    path('tea_kqt/', views.tea_kqt, name='tea_kqt'),
    path('student_dashboard/', views.student_dashboard, name='student_dashboard'),
    path('student_subj/', views.student_subj, name='student_subj'),
    path('student_exam/', views.student_exam, name='student_exam'),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

