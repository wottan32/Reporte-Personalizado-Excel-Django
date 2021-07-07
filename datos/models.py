from django.db import models
from django.utils import timezone


class Datos(models.Model):
    id = models.IntegerField(primary_key=True)
    empresa_id = models.IntegerField(max_length=255, blank=False)
    nombre = models.CharField(max_length=255, blank=False, default='')
    razon_social = models.CharField(max_length=255, blank=False)
    rut = models.IntegerField(max_length=255, blank=False)
    plazo_pago = models.IntegerField(max_length=255, blank=False)
    oc = models.BooleanField(blank=False)
    giro = models.CharField(max_length=255, blank=False)
    contacto_factura = models.CharField(max_length=255, blank=False)
    direccion_legal = models.CharField(max_length=255, blank=False)
    comuna_legal = models.CharField(max_length=255, blank=False)
    contactos = models.CharField(max_length=255, blank=False)

    created = models.DateTimeField(default=timezone.now)
    modified = models.DateTimeField(default=timezone.now)

    class Meta:
        ordering = ('-id',)
        verbose_name = 'Datos'
        verbose_name_plural = 'Datos'
        # abstract = True

    def __str__(self):
        return self.id
