from django.db import models

class Person(models.Model):
    full_name = models.CharField(max_length=200, verbose_name="Full Name")
    email = models.EmailField(max_length=100, verbose_name="Email")
    educational_program = models.CharField(max_length=100, verbose_name="Educational Program")
    created = models.DateTimeField(auto_now_add=True, verbose_name="Creation Date")
    updated = models.DateTimeField(auto_now=True, verbose_name="Update Date")

    class Meta:
        verbose_name = "person"
        verbose_name_plural = "people"
        ordering = ["-created"]

    def __str__(self):
        return self.full_name
