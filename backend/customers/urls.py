from django.urls import path
from . import views

urlpatterns = [
    path('status/', views.customer_status, name='customer-status'),
    path('upload/', views.customer_upload, name='customer-upload'),
    path('save/', views.customer_save, name='customer-save'),
    path('cancel_upload/', views.customer_cancel_upload, name='customer-cancel-upload'),
    path('sheets/', views.customer_sheets, name='customer-sheets'),
    path('list/', views.customer_list, name='customer-list'),
    path('search/', views.customer_search, name='customer-search'),
    path('records/', views.customer_records_by_name, name='customer-records-by-name'),
    path('record/', views.customer_record, name='customer-record'),
    path('customer/', views.delete_customer_view, name='customer-delete'),
    path('delete_customer/', views.delete_customer_view, name='customer-delete-alias'),
    path('all/', views.delete_all_customers_view, name='customer-delete-all'),
    path('delete_all/', views.delete_all_customers_view, name='customer-delete-all-alias'),
    path('replace/', views.customer_replace, name='customer-replace'),
    path('photo/', views.customer_photo, name='customer-photo'),
    path('auth/login/', views.auth_login, name='auth-login'),
    path('auth/verify/', views.auth_verify, name='auth-verify'),
]
