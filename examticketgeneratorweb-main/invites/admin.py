from django.contrib import admin
from .models import Invite

@admin.register(Invite)
class InviteAdmin(admin.ModelAdmin):
    list_display = ['code', 'used', 'created_at']
    search_fields = ['code']
    list_filter = ['used']
