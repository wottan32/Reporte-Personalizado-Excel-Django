# api/serializers.py
from rest_framework import serializers
from datos.models import Datos


class DatosSerializer(serializers.ModelSerializer):
    class Meta:
        model = Datos
        fields = ('id', 'empresa_id', 'nombre', 'razon_social', 'rut', 'plazo_pago', 'oc',
                  'giro', 'contacto_factura', 'direccion_legal', 'comuna_legal', 'contactos', 'created',
                  'modified')
