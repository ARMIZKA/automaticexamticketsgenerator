from django.db import models

class Invite(models.Model):
    code = models.CharField(max_length=32, unique=True)
    used = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.code} ({'использован' if self.used else 'не использован'})"
