# -*- coding: utf-8 -*-
import datetime
from south.db import db
from south.v2 import SchemaMigration
from django.db import models


class Migration(SchemaMigration):

    def forwards(self, orm):
        # Adding model 'cat001continentes'
        db.create_table(u'tracking_cat001continentes', (
            ('continente_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre_continente', self.gf('django.db.models.fields.CharField')(max_length=20)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001continentes'])

        # Adding model 'cat001tipo_contenedores'
        db.create_table(u'tracking_cat001tipo_contenedores', (
            ('tipo_contenedor_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('descripcion', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('categoria', self.gf('django.db.models.fields.CharField')(max_length=20, null=True, blank=True)),
            ('siglas', self.gf('django.db.models.fields.CharField')(max_length=4, null=True, blank=True)),
            ('dimensiones', self.gf('django.db.models.fields.CharField')(max_length=4, null=True, blank=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001tipo_contenedores'])

        # Adding model 'cat001monedas'
        db.create_table(u'tracking_cat001monedas', (
            ('moneda_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre_moneda', self.gf('django.db.models.fields.CharField')(max_length=20)),
            ('clave', self.gf('django.db.models.fields.CharField')(max_length=10, db_index=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001monedas'])

        # Adding model 'cat001paises'
        db.create_table(u'tracking_cat001paises', (
            ('pais_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre_pais', self.gf('django.db.models.fields.CharField')(max_length=100)),
            ('nombre_pais_ingles', self.gf('django.db.models.fields.CharField')(max_length=100, null=True)),
            ('codigo_numerico', self.gf('django.db.models.fields.CharField')(max_length=5, null=True, db_index=True)),
            ('codigo_2_letras', self.gf('django.db.models.fields.CharField')(max_length=2, null=True, db_index=True)),
            ('codigo_3_letras', self.gf('django.db.models.fields.CharField')(max_length=3, db_index=True)),
            ('continente', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001continentes'], null=True)),
            ('moneda', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001monedas'], null=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001paises'])

        # Adding model 'cat001puertos'
        db.create_table(u'tracking_cat001puertos', (
            ('puerto_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre_puerto', self.gf('django.db.models.fields.CharField')(max_length=150, db_index=True)),
            ('pais', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001paises'])),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001puertos'])

        # Adding model 'cat001aduanas'
        db.create_table(u'tracking_cat001aduanas', (
            ('aduana_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('clave', self.gf('django.db.models.fields.IntegerField')(unique=True)),
            ('descripcion', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001aduanas'])

        # Adding model 'cat001direcciones'
        db.create_table(u'tracking_cat001direcciones', (
            ('direccion_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('calle', self.gf('django.db.models.fields.TextField')()),
            ('numero_exterior', self.gf('django.db.models.fields.CharField')(max_length=15)),
            ('numero_interior', self.gf('django.db.models.fields.CharField')(max_length=15)),
            ('colonia', self.gf('django.db.models.fields.CharField')(max_length=50)),
            ('municipio', self.gf('django.db.models.fields.CharField')(max_length=50)),
            ('estado', self.gf('django.db.models.fields.CharField')(max_length=50)),
            ('codigo_postal', self.gf('django.db.models.fields.CharField')(max_length=20)),
            ('pais', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001paises'])),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001direcciones'])

        # Adding model 'cat001patentes'
        db.create_table(u'tracking_cat001patentes', (
            ('patente_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('clave', self.gf('django.db.models.fields.CharField')(unique=True, max_length=4)),
            ('agente_aduanal', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001patentes'])

        # Adding model 'cat001oficinas'
        db.create_table(u'tracking_cat001oficinas', (
            ('oficina_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('razon_social', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('clave_oficina', self.gf('django.db.models.fields.CharField')(max_length=10)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001oficinas'])

        # Adding model 'cat001razones_sociales'
        db.create_table(u'tracking_cat001razones_sociales', (
            ('razon_social_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('razon_social', self.gf('django.db.models.fields.CharField')(max_length=255, db_index=True)),
            ('descripcion', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('rfc', self.gf('django.db.models.fields.CharField')(max_length=20, db_index=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001razones_sociales'])

        # Adding model 'cat001navieras'
        db.create_table(u'tracking_cat001navieras', (
            ('naviera_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre_naviera', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('sitio_web', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001navieras'])

        # Adding model 'cat001buques'
        db.create_table(u'tracking_cat001buques', (
            ('buque', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre_buque', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('naviera', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001navieras'])),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001buques'])

        # Adding model 'cat001linea_aerea'
        db.create_table(u'tracking_cat001linea_aerea', (
            ('linea_aerea', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('sitio_web', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001linea_aerea'])

        # Adding model 'cat001impuestos'
        db.create_table(u'tracking_cat001impuestos', (
            ('impuesto_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('descripcion', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('clave', self.gf('django.db.models.fields.IntegerField')()),
            ('abreviacion', self.gf('django.db.models.fields.CharField')(max_length=10)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001impuestos'])

        # Adding model 'cat001proveedores'
        db.create_table(u'tracking_cat001proveedores', (
            ('proveedor_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('identificador_fiscal', self.gf('django.db.models.fields.CharField')(db_index=True, max_length=50, null=True, blank=True)),
            ('vinculado', self.gf('django.db.models.fields.SmallIntegerField')(null=True, blank=True)),
            ('codigo_samsung', self.gf('django.db.models.fields.CharField')(max_length=50, null=True, blank=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001proveedores'])

        # Adding model 'cat001unidad_medida'
        db.create_table(u'tracking_cat001unidad_medida', (
            ('unidad_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('descripcion', self.gf('django.db.models.fields.CharField')(max_length=30)),
            ('clave', self.gf('django.db.models.fields.CharField')(max_length=5, null=True, blank=True)),
            ('abreviacion', self.gf('django.db.models.fields.CharField')(max_length=10, null=True, blank=True)),
            ('breviacion_ingles', self.gf('django.db.models.fields.CharField')(max_length=10, null=True, blank=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001unidad_medida'])

        # Adding model 'cat001articulos'
        db.create_table(u'tracking_cat001articulos', (
            ('articulo_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('descripcion', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('fraccion', self.gf('django.db.models.fields.CharField')(max_length=8, null=True, blank=True)),
            ('fraccion_8', self.gf('django.db.models.fields.CharField')(max_length=8, null=True, blank=True)),
            ('observaciones', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('tipo_mercancia', self.gf('django.db.models.fields.CharField')(max_length=4, null=True, blank=True)),
            ('unidad', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001unidad_medida'], null=True, blank=True)),
            ('codigo_producto', self.gf('django.db.models.fields.CharField')(max_length=150, db_index=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001articulos'])

        # Adding model 'rel001claves_cliente'
        db.create_table(u'tracking_rel001claves_cliente', (
            ('clave_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('numero_clave', self.gf('django.db.models.fields.CharField')(max_length=15, db_index=True)),
            ('descripcion', self.gf('django.db.models.fields.CharField')(max_length=255)),
            ('razon_social', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001razones_sociales'])),
            ('oficina', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001oficinas'])),
            ('activo', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['rel001claves_cliente'])

        # Adding model 'rel001direcciones_clientes'
        db.create_table(u'tracking_rel001direcciones_clientes', (
            ('direccion_cliente_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('direccion', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001direcciones'])),
            ('cliente', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.rel001claves_cliente'])),
            ('activo', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['rel001direcciones_clientes'])

        # Adding model 'rel001direcciones_proveedor'
        db.create_table(u'tracking_rel001direcciones_proveedor', (
            ('direccion_proveedor_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('direccion', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001direcciones'])),
            ('proveedor', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001proveedores'])),
            ('activo', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['rel001direcciones_proveedor'])

        # Adding model 'reg001operacion'
        db.create_table(u'tracking_reg001operacion', (
            ('operacion_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('oficina', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001oficinas'])),
            ('referencia', self.gf('django.db.models.fields.CharField')(max_length=20, db_index=True)),
            ('tipo', self.gf('django.db.models.fields.CharField')(max_length=4)),
            ('clave_pedimento', self.gf('django.db.models.fields.CharField')(max_length=2, db_index=True)),
            ('aduana', self.gf('django.db.models.fields.CharField')(max_length=2)),
            ('patente', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001patentes'])),
            ('seccion', self.gf('django.db.models.fields.CharField')(max_length=1)),
            ('pedimento', self.gf('django.db.models.fields.CharField')(max_length=10)),
            ('firmado', self.gf('django.db.models.fields.BooleanField')(default=False)),
            ('firma', self.gf('django.db.models.fields.CharField')(max_length=50, null=True, blank=True)),
            ('puerto_embarque', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001puertos'])),
            ('clave_embarque', self.gf('django.db.models.fields.IntegerField')(null=True, blank=True)),
            ('pais_origen', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001paises'])),
            ('pais_procedencia', self.gf('django.db.models.fields.related.ForeignKey')(related_name='+', to=orm['tracking.cat001paises'])),
            ('tipo_cambio', self.gf('django.db.models.fields.DecimalField')(max_digits=12, decimal_places=6)),
            ('buque', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001buques'])),
            ('linea_aerea', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001linea_aerea'])),
            ('fecha_pago', self.gf('django.db.models.fields.DateField')(null=True, blank=True)),
            ('fecha_entrada', self.gf('django.db.models.fields.DateField')(null=True, blank=True)),
            ('fecha_revalidacion', self.gf('django.db.models.fields.DateField')(null=True, blank=True)),
            ('fecha_despacho', self.gf('django.db.models.fields.DateField')(null=True, blank=True)),
            ('fecha_bl', self.gf('django.db.models.fields.DateField')(null=True, blank=True)),
            ('fecha_arribo', self.gf('django.db.models.fields.DateField')(null=True, blank=True)),
            ('peso_bruto', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=4, blank=True)),
            ('total_bultos', self.gf('django.db.models.fields.CharField')(max_length=150, null=True, blank=True)),
            ('contenedores_pedimento', self.gf('django.db.models.fields.IntegerField')(max_length=4)),
            ('contenedores_embarque', self.gf('django.db.models.fields.IntegerField')(max_length=4)),
            ('fletes', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=4, blank=True)),
            ('seguros', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=4, blank=True)),
            ('otros_incrementables', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=4, blank=True)),
            ('valor_dls_factura', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=4, blank=True)),
            ('valor_monex_factura', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=4, blank=True)),
            ('valor_aduana', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=4, blank=True)),
            ('clave_cliente', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.rel001claves_cliente'])),
            ('direccion_cliente', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001direcciones'])),
            ('pais_cliente', self.gf('django.db.models.fields.related.ForeignKey')(related_name='++', to=orm['tracking.cat001paises'])),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
            ('estatus', self.gf('django.db.models.fields.CharField')(default='A', max_length=1, db_index=True, blank=True)),
        ))
        db.send_create_signal(u'tracking', ['reg001operacion'])

        # Adding model 'det001impuestos_pedimento'
        db.create_table(u'tracking_det001impuestos_pedimento', (
            ('impuestos_pedimento_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('operacion', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.reg001operacion'])),
            ('impuesto', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001impuestos'])),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['det001impuestos_pedimento'])

        # Adding model 'det001contenedores'
        db.create_table(u'tracking_det001contenedores', (
            ('contenedor_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('numero_cotenedor', self.gf('django.db.models.fields.CharField')(max_length=50, null=True)),
            ('tipo_contenedor', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001tipo_contenedores'])),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['det001contenedores'])

        # Adding model 'det001candados'
        db.create_table(u'tracking_det001candados', (
            ('candado_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('numero', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('contenedor', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.det001contenedores'])),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['det001candados'])

        # Adding model 'det001guias'
        db.create_table(u'tracking_det001guias', (
            ('guia_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('operacion', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.reg001operacion'])),
            ('numero_guia', self.gf('django.db.models.fields.CharField')(max_length=50)),
            ('tipo', self.gf('django.db.models.fields.CharField')(default='ma', max_length=6)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['det001guias'])

        # Adding model 'det001identificadores_pedimento'
        db.create_table(u'tracking_det001identificadores_pedimento', (
            ('identificador_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('operacion', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.reg001operacion'])),
            ('clave_identificador', self.gf('django.db.models.fields.CharField')(max_length=5)),
            ('descripcion', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('permiso', self.gf('django.db.models.fields.CharField')(max_length=50, null=True, blank=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['det001identificadores_pedimento'])

        # Adding model 'reg001facturas'
        db.create_table(u'tracking_reg001facturas', (
            ('factura_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('operacion', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.reg001operacion'], null=True, blank=True)),
            ('proveedor', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001proveedores'])),
            ('pais_factura', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001paises'])),
            ('moneda_factura', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001monedas'])),
            ('fecha_factura', self.gf('django.db.models.fields.DateField')()),
            ('incoterm', self.gf('django.db.models.fields.CharField')(max_length=3, null=True, blank=True)),
            ('factor_moneda', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=12, decimal_places=6, blank=True)),
            ('valor_dls', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=6, blank=True)),
            ('valor_monex', self.gf('django.db.models.fields.DecimalField')(null=True, max_digits=32, decimal_places=6, blank=True)),
            ('direccion', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001direcciones'])),
            ('edocument', self.gf('django.db.models.fields.CharField')(max_length=25, null=True, blank=True)),
            ('numero_operacion', self.gf('django.db.models.fields.CharField')(max_length=25, null=True, blank=True)),
            ('folio', self.gf('django.db.models.fields.CharField')(max_length=30, null=True, blank=True)),
            ('serie', self.gf('django.db.models.fields.CharField')(max_length=30, null=True, blank=True)),
            ('borrado', self.gf('django.db.models.fields.BooleanField')(default=False, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['reg001facturas'])

        # Adding model 'det001facturas'
        db.create_table(u'tracking_det001facturas', (
            ('detalle_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('factura', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.reg001facturas'])),
        ))
        db.send_create_signal(u'tracking', ['det001facturas'])

        # Adding model 'det001partidas'
        db.create_table(u'tracking_det001partidas', (
            ('partida_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('operacion', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.reg001operacion'])),
            ('numero_partida', self.gf('django.db.models.fields.SmallIntegerField')()),
            ('detalle_mercancia', self.gf('django.db.models.fields.TextField')()),
            ('fraccion', self.gf('django.db.models.fields.CharField')(max_length=8)),
            ('cantidad_comercial', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('um_comercial', self.gf('django.db.models.fields.SmallIntegerField')()),
            ('cantidad_tarifa', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('um_tarifa', self.gf('django.db.models.fields.SmallIntegerField')()),
            ('valor_aduana', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('valor_mercancia', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('valor_dls', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('numeros_serie', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('observaciones', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('metodo_valoracion', self.gf('django.db.models.fields.SmallIntegerField')()),
            ('vinculacion', self.gf('django.db.models.fields.SmallIntegerField')()),
            ('cuota_operacion', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('tasa_igie', self.gf('django.db.models.fields.DecimalField')(max_digits=6, decimal_places=3)),
            ('igie', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('tasa_iva', self.gf('django.db.models.fields.DecimalField')(max_digits=6, decimal_places=3)),
            ('iva', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('tasa_isan', self.gf('django.db.models.fields.DecimalField')(max_digits=6, decimal_places=3)),
            ('isan', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('tasa_ieps', self.gf('django.db.models.fields.DecimalField')(max_digits=6, decimal_places=3)),
            ('ieps', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('tasa_max', self.gf('django.db.models.fields.DecimalField')(max_digits=6, decimal_places=3)),
            ('tasa_cc', self.gf('django.db.models.fields.DecimalField')(max_digits=6, decimal_places=3)),
            ('cc', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('recargos', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('factor_actualizacion', self.gf('django.db.models.fields.DecimalField')(max_digits=6, decimal_places=4)),
            ('precio_unitario', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
            ('dta_partida', self.gf('django.db.models.fields.DecimalField')(max_digits=32, decimal_places=6)),
        ))
        db.send_create_signal(u'tracking', ['det001partidas'])

        # Adding model 'det001identificadores_partida'
        db.create_table(u'tracking_det001identificadores_partida', (
            ('identificador_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('partida', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.det001partidas'])),
            ('identificador', self.gf('django.db.models.fields.CharField')(max_length=10, db_index=True)),
            ('descripcion', self.gf('django.db.models.fields.CharField')(max_length=100)),
            ('complemento_1', self.gf('django.db.models.fields.CharField')(max_length=50, null=True, blank=True)),
            ('complemento_2', self.gf('django.db.models.fields.CharField')(max_length=50, null=True, blank=True)),
            ('complemento_3', self.gf('django.db.models.fields.CharField')(max_length=50, null=True, blank=True)),
        ))
        db.send_create_signal(u'tracking', ['det001identificadores_partida'])

        # Adding model 'bit001archivo_mercancias'
        db.create_table(u'tracking_bit001archivo_mercancias', (
            ('archivo_mercancia_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('archivo', self.gf('django.db.models.fields.files.FileField')(max_length=100)),
            ('nombre_original', self.gf('django.db.models.fields.CharField')(max_length=50, null=True, blank=True)),
            ('fecha_carga', self.gf('django.db.models.fields.DateTimeField')(auto_now=True, auto_now_add=True, db_index=True, blank=True)),
            ('estatus', self.gf('django.db.models.fields.CharField')(default='D', max_length=2, db_index=True)),
            ('usuario', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['auth.User'])),
            ('bl', self.gf('django.db.models.fields.CharField')(max_length=30, db_index=True)),
            ('numero_factura', self.gf('django.db.models.fields.CharField')(max_length=254, db_index=True)),
            ('moneda', self.gf('django.db.models.fields.CharField')(max_length=15)),
            ('proveedor', self.gf('django.db.models.fields.CharField')(max_length=255, null=True, blank=True)),
        ))
        db.send_create_signal(u'tracking', ['bit001archivo_mercancias'])

        # Adding model 'bit001facturas'
        db.create_table(u'tracking_bit001facturas', (
            (u'id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('patente', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001patentes'])),
            ('proveedor', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001proveedores'])),
            ('cliente', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.rel001claves_cliente'])),
            ('archivo', self.gf('django.db.models.fields.files.FileField')(max_length=100)),
            ('fecha_carga', self.gf('django.db.models.fields.DateTimeField')(auto_now=True, blank=True)),
            ('ultima_actualizacion', self.gf('django.db.models.fields.DateTimeField')(auto_now=True, blank=True)),
            ('usuario', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['auth.User'])),
            ('numero_factura', self.gf('django.db.models.fields.CharField')(max_length=255, db_index=True)),
            ('nombre_proveedor', self.gf('django.db.models.fields.CharField')(max_length=255, null=True, blank=True)),
            ('fecha_factura', self.gf('django.db.models.fields.DateField')(null=True, blank=True)),
        ))
        db.send_create_signal(u'tracking', ['bit001facturas'])

        # Adding model 'cat001CamposXls'
        db.create_table(u'tracking_cat001camposxls', (
            ('campo_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre', self.gf('django.db.models.fields.CharField')(max_length=50, db_index=True)),
            ('descripcion', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('tipo', self.gf('django.db.models.fields.CharField')(default='C', max_length=1, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001CamposXls'])

        # Adding model 'cat001FormatosXls'
        db.create_table(u'tracking_cat001formatosxls', (
            ('formato_id', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('nombre', self.gf('django.db.models.fields.CharField')(max_length=25, db_index=True)),
            ('descripcion', self.gf('django.db.models.fields.TextField')(null=True, blank=True)),
            ('fecha_alta', self.gf('django.db.models.fields.DateTimeField')(auto_now=True, blank=True)),
            ('activo', self.gf('django.db.models.fields.BooleanField')(default=True, db_index=True)),
        ))
        db.send_create_signal(u'tracking', ['cat001FormatosXls'])

        # Adding model 'det001FormatoCampos'
        db.create_table(u'tracking_det001formatocampos', (
            ('campos', self.gf('django.db.models.fields.AutoField')(primary_key=True)),
            ('formato', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001FormatosXls'])),
            ('nombre_columna', self.gf('django.db.models.fields.CharField')(max_length=150)),
            ('campo', self.gf('django.db.models.fields.related.ForeignKey')(to=orm['tracking.cat001CamposXls'])),
        ))
        db.send_create_signal(u'tracking', ['det001FormatoCampos'])

        # Adding unique constraint on 'det001FormatoCampos', fields ['formato', 'campos']
        db.create_unique(u'tracking_det001formatocampos', ['formato_id', 'campos'])


    def backwards(self, orm):
        # Removing unique constraint on 'det001FormatoCampos', fields ['formato', 'campos']
        db.delete_unique(u'tracking_det001formatocampos', ['formato_id', 'campos'])

        # Deleting model 'cat001continentes'
        db.delete_table(u'tracking_cat001continentes')

        # Deleting model 'cat001tipo_contenedores'
        db.delete_table(u'tracking_cat001tipo_contenedores')

        # Deleting model 'cat001monedas'
        db.delete_table(u'tracking_cat001monedas')

        # Deleting model 'cat001paises'
        db.delete_table(u'tracking_cat001paises')

        # Deleting model 'cat001puertos'
        db.delete_table(u'tracking_cat001puertos')

        # Deleting model 'cat001aduanas'
        db.delete_table(u'tracking_cat001aduanas')

        # Deleting model 'cat001direcciones'
        db.delete_table(u'tracking_cat001direcciones')

        # Deleting model 'cat001patentes'
        db.delete_table(u'tracking_cat001patentes')

        # Deleting model 'cat001oficinas'
        db.delete_table(u'tracking_cat001oficinas')

        # Deleting model 'cat001razones_sociales'
        db.delete_table(u'tracking_cat001razones_sociales')

        # Deleting model 'cat001navieras'
        db.delete_table(u'tracking_cat001navieras')

        # Deleting model 'cat001buques'
        db.delete_table(u'tracking_cat001buques')

        # Deleting model 'cat001linea_aerea'
        db.delete_table(u'tracking_cat001linea_aerea')

        # Deleting model 'cat001impuestos'
        db.delete_table(u'tracking_cat001impuestos')

        # Deleting model 'cat001proveedores'
        db.delete_table(u'tracking_cat001proveedores')

        # Deleting model 'cat001unidad_medida'
        db.delete_table(u'tracking_cat001unidad_medida')

        # Deleting model 'cat001articulos'
        db.delete_table(u'tracking_cat001articulos')

        # Deleting model 'rel001claves_cliente'
        db.delete_table(u'tracking_rel001claves_cliente')

        # Deleting model 'rel001direcciones_clientes'
        db.delete_table(u'tracking_rel001direcciones_clientes')

        # Deleting model 'rel001direcciones_proveedor'
        db.delete_table(u'tracking_rel001direcciones_proveedor')

        # Deleting model 'reg001operacion'
        db.delete_table(u'tracking_reg001operacion')

        # Deleting model 'det001impuestos_pedimento'
        db.delete_table(u'tracking_det001impuestos_pedimento')

        # Deleting model 'det001contenedores'
        db.delete_table(u'tracking_det001contenedores')

        # Deleting model 'det001candados'
        db.delete_table(u'tracking_det001candados')

        # Deleting model 'det001guias'
        db.delete_table(u'tracking_det001guias')

        # Deleting model 'det001identificadores_pedimento'
        db.delete_table(u'tracking_det001identificadores_pedimento')

        # Deleting model 'reg001facturas'
        db.delete_table(u'tracking_reg001facturas')

        # Deleting model 'det001facturas'
        db.delete_table(u'tracking_det001facturas')

        # Deleting model 'det001partidas'
        db.delete_table(u'tracking_det001partidas')

        # Deleting model 'det001identificadores_partida'
        db.delete_table(u'tracking_det001identificadores_partida')

        # Deleting model 'bit001archivo_mercancias'
        db.delete_table(u'tracking_bit001archivo_mercancias')

        # Deleting model 'bit001facturas'
        db.delete_table(u'tracking_bit001facturas')

        # Deleting model 'cat001CamposXls'
        db.delete_table(u'tracking_cat001camposxls')

        # Deleting model 'cat001FormatosXls'
        db.delete_table(u'tracking_cat001formatosxls')

        # Deleting model 'det001FormatoCampos'
        db.delete_table(u'tracking_det001formatocampos')


    models = {
        u'auth.group': {
            'Meta': {'object_name': 'Group'},
            u'id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'name': ('django.db.models.fields.CharField', [], {'unique': 'True', 'max_length': '80'}),
            'permissions': ('django.db.models.fields.related.ManyToManyField', [], {'to': u"orm['auth.Permission']", 'symmetrical': 'False', 'blank': 'True'})
        },
        u'auth.permission': {
            'Meta': {'ordering': "(u'content_type__app_label', u'content_type__model', u'codename')", 'unique_together': "((u'content_type', u'codename'),)", 'object_name': 'Permission'},
            'codename': ('django.db.models.fields.CharField', [], {'max_length': '100'}),
            'content_type': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['contenttypes.ContentType']"}),
            u'id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'name': ('django.db.models.fields.CharField', [], {'max_length': '50'})
        },
        u'auth.user': {
            'Meta': {'object_name': 'User'},
            'date_joined': ('django.db.models.fields.DateTimeField', [], {'default': 'datetime.datetime.now'}),
            'email': ('django.db.models.fields.EmailField', [], {'max_length': '75', 'blank': 'True'}),
            'first_name': ('django.db.models.fields.CharField', [], {'max_length': '30', 'blank': 'True'}),
            'groups': ('django.db.models.fields.related.ManyToManyField', [], {'symmetrical': 'False', 'related_name': "u'user_set'", 'blank': 'True', 'to': u"orm['auth.Group']"}),
            u'id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'is_active': ('django.db.models.fields.BooleanField', [], {'default': 'True'}),
            'is_staff': ('django.db.models.fields.BooleanField', [], {'default': 'False'}),
            'is_superuser': ('django.db.models.fields.BooleanField', [], {'default': 'False'}),
            'last_login': ('django.db.models.fields.DateTimeField', [], {'default': 'datetime.datetime.now'}),
            'last_name': ('django.db.models.fields.CharField', [], {'max_length': '30', 'blank': 'True'}),
            'password': ('django.db.models.fields.CharField', [], {'max_length': '128'}),
            'user_permissions': ('django.db.models.fields.related.ManyToManyField', [], {'symmetrical': 'False', 'related_name': "u'user_set'", 'blank': 'True', 'to': u"orm['auth.Permission']"}),
            'username': ('django.db.models.fields.CharField', [], {'unique': 'True', 'max_length': '30'})
        },
        u'contenttypes.contenttype': {
            'Meta': {'ordering': "('name',)", 'unique_together': "(('app_label', 'model'),)", 'object_name': 'ContentType', 'db_table': "'django_content_type'"},
            'app_label': ('django.db.models.fields.CharField', [], {'max_length': '100'}),
            u'id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'model': ('django.db.models.fields.CharField', [], {'max_length': '100'}),
            'name': ('django.db.models.fields.CharField', [], {'max_length': '100'})
        },
        u'tracking.bit001archivo_mercancias': {
            'Meta': {'object_name': 'bit001archivo_mercancias'},
            'archivo': ('django.db.models.fields.files.FileField', [], {'max_length': '100'}),
            'archivo_mercancia_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'bl': ('django.db.models.fields.CharField', [], {'max_length': '30', 'db_index': 'True'}),
            'estatus': ('django.db.models.fields.CharField', [], {'default': "'D'", 'max_length': '2', 'db_index': 'True'}),
            'fecha_carga': ('django.db.models.fields.DateTimeField', [], {'auto_now': 'True', 'auto_now_add': 'True', 'db_index': 'True', 'blank': 'True'}),
            'moneda': ('django.db.models.fields.CharField', [], {'max_length': '15'}),
            'nombre_original': ('django.db.models.fields.CharField', [], {'max_length': '50', 'null': 'True', 'blank': 'True'}),
            'numero_factura': ('django.db.models.fields.CharField', [], {'max_length': '254', 'db_index': 'True'}),
            'proveedor': ('django.db.models.fields.CharField', [], {'max_length': '255', 'null': 'True', 'blank': 'True'}),
            'usuario': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['auth.User']"})
        },
        u'tracking.bit001facturas': {
            'Meta': {'object_name': 'bit001facturas'},
            'archivo': ('django.db.models.fields.files.FileField', [], {'max_length': '100'}),
            'cliente': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.rel001claves_cliente']"}),
            'fecha_carga': ('django.db.models.fields.DateTimeField', [], {'auto_now': 'True', 'blank': 'True'}),
            'fecha_factura': ('django.db.models.fields.DateField', [], {'null': 'True', 'blank': 'True'}),
            u'id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'nombre_proveedor': ('django.db.models.fields.CharField', [], {'max_length': '255', 'null': 'True', 'blank': 'True'}),
            'numero_factura': ('django.db.models.fields.CharField', [], {'max_length': '255', 'db_index': 'True'}),
            'patente': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001patentes']"}),
            'proveedor': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001proveedores']"}),
            'ultima_actualizacion': ('django.db.models.fields.DateTimeField', [], {'auto_now': 'True', 'blank': 'True'}),
            'usuario': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['auth.User']"})
        },
        u'tracking.cat001aduanas': {
            'Meta': {'object_name': 'cat001aduanas'},
            'aduana_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'clave': ('django.db.models.fields.IntegerField', [], {'unique': 'True'}),
            'descripcion': ('django.db.models.fields.CharField', [], {'max_length': '150'})
        },
        u'tracking.cat001articulos': {
            'Meta': {'object_name': 'cat001articulos'},
            'articulo_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'codigo_producto': ('django.db.models.fields.CharField', [], {'max_length': '150', 'db_index': 'True'}),
            'descripcion': ('django.db.models.fields.CharField', [], {'max_length': '150'}),
            'fraccion': ('django.db.models.fields.CharField', [], {'max_length': '8', 'null': 'True', 'blank': 'True'}),
            'fraccion_8': ('django.db.models.fields.CharField', [], {'max_length': '8', 'null': 'True', 'blank': 'True'}),
            'observaciones': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'}),
            'tipo_mercancia': ('django.db.models.fields.CharField', [], {'max_length': '4', 'null': 'True', 'blank': 'True'}),
            'unidad': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001unidad_medida']", 'null': 'True', 'blank': 'True'})
        },
        u'tracking.cat001buques': {
            'Meta': {'object_name': 'cat001buques'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'buque': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'naviera': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001navieras']"}),
            'nombre_buque': ('django.db.models.fields.CharField', [], {'max_length': '150'})
        },
        u'tracking.cat001camposxls': {
            'Meta': {'object_name': 'cat001CamposXls'},
            'campo_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'descripcion': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'}),
            'nombre': ('django.db.models.fields.CharField', [], {'max_length': '50', 'db_index': 'True'}),
            'tipo': ('django.db.models.fields.CharField', [], {'default': "'C'", 'max_length': '1', 'db_index': 'True'})
        },
        u'tracking.cat001continentes': {
            'Meta': {'object_name': 'cat001continentes'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'continente_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'nombre_continente': ('django.db.models.fields.CharField', [], {'max_length': '20'})
        },
        u'tracking.cat001direcciones': {
            'Meta': {'object_name': 'cat001direcciones'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'calle': ('django.db.models.fields.TextField', [], {}),
            'codigo_postal': ('django.db.models.fields.CharField', [], {'max_length': '20'}),
            'colonia': ('django.db.models.fields.CharField', [], {'max_length': '50'}),
            'direccion_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'estado': ('django.db.models.fields.CharField', [], {'max_length': '50'}),
            'municipio': ('django.db.models.fields.CharField', [], {'max_length': '50'}),
            'numero_exterior': ('django.db.models.fields.CharField', [], {'max_length': '15'}),
            'numero_interior': ('django.db.models.fields.CharField', [], {'max_length': '15'}),
            'pais': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001paises']"})
        },
        u'tracking.cat001formatosxls': {
            'Meta': {'object_name': 'cat001FormatosXls'},
            'activo': ('django.db.models.fields.BooleanField', [], {'default': 'True', 'db_index': 'True'}),
            'descripcion': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'}),
            'fecha_alta': ('django.db.models.fields.DateTimeField', [], {'auto_now': 'True', 'blank': 'True'}),
            'formato_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'nombre': ('django.db.models.fields.CharField', [], {'max_length': '25', 'db_index': 'True'})
        },
        u'tracking.cat001impuestos': {
            'Meta': {'object_name': 'cat001impuestos'},
            'abreviacion': ('django.db.models.fields.CharField', [], {'max_length': '10'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'clave': ('django.db.models.fields.IntegerField', [], {}),
            'descripcion': ('django.db.models.fields.CharField', [], {'max_length': '150'}),
            'impuesto_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'})
        },
        u'tracking.cat001linea_aerea': {
            'Meta': {'object_name': 'cat001linea_aerea'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'linea_aerea': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'nombre': ('django.db.models.fields.CharField', [], {'max_length': '150'}),
            'sitio_web': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'})
        },
        u'tracking.cat001monedas': {
            'Meta': {'object_name': 'cat001monedas'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'clave': ('django.db.models.fields.CharField', [], {'max_length': '10', 'db_index': 'True'}),
            'moneda_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'nombre_moneda': ('django.db.models.fields.CharField', [], {'max_length': '20'})
        },
        u'tracking.cat001navieras': {
            'Meta': {'object_name': 'cat001navieras'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'naviera_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'nombre_naviera': ('django.db.models.fields.CharField', [], {'max_length': '150'}),
            'sitio_web': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'})
        },
        u'tracking.cat001oficinas': {
            'Meta': {'object_name': 'cat001oficinas'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'clave_oficina': ('django.db.models.fields.CharField', [], {'max_length': '10'}),
            'nombre': ('django.db.models.fields.CharField', [], {'max_length': '150'}),
            'oficina_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'razon_social': ('django.db.models.fields.CharField', [], {'max_length': '150'})
        },
        u'tracking.cat001paises': {
            'Meta': {'object_name': 'cat001paises'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'codigo_2_letras': ('django.db.models.fields.CharField', [], {'max_length': '2', 'null': 'True', 'db_index': 'True'}),
            'codigo_3_letras': ('django.db.models.fields.CharField', [], {'max_length': '3', 'db_index': 'True'}),
            'codigo_numerico': ('django.db.models.fields.CharField', [], {'max_length': '5', 'null': 'True', 'db_index': 'True'}),
            'continente': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001continentes']", 'null': 'True'}),
            'moneda': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001monedas']", 'null': 'True'}),
            'nombre_pais': ('django.db.models.fields.CharField', [], {'max_length': '100'}),
            'nombre_pais_ingles': ('django.db.models.fields.CharField', [], {'max_length': '100', 'null': 'True'}),
            'pais_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'})
        },
        u'tracking.cat001patentes': {
            'Meta': {'object_name': 'cat001patentes'},
            'agente_aduanal': ('django.db.models.fields.CharField', [], {'max_length': '150'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'clave': ('django.db.models.fields.CharField', [], {'unique': 'True', 'max_length': '4'}),
            'patente_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'})
        },
        u'tracking.cat001proveedores': {
            'Meta': {'object_name': 'cat001proveedores'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'codigo_samsung': ('django.db.models.fields.CharField', [], {'max_length': '50', 'null': 'True', 'blank': 'True'}),
            'identificador_fiscal': ('django.db.models.fields.CharField', [], {'db_index': 'True', 'max_length': '50', 'null': 'True', 'blank': 'True'}),
            'nombre': ('django.db.models.fields.CharField', [], {'max_length': '150'}),
            'proveedor_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'vinculado': ('django.db.models.fields.SmallIntegerField', [], {'null': 'True', 'blank': 'True'})
        },
        u'tracking.cat001puertos': {
            'Meta': {'object_name': 'cat001puertos'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'nombre_puerto': ('django.db.models.fields.CharField', [], {'max_length': '150', 'db_index': 'True'}),
            'pais': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001paises']"}),
            'puerto_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'})
        },
        u'tracking.cat001razones_sociales': {
            'Meta': {'object_name': 'cat001razones_sociales'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'descripcion': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'}),
            'razon_social': ('django.db.models.fields.CharField', [], {'max_length': '255', 'db_index': 'True'}),
            'razon_social_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'rfc': ('django.db.models.fields.CharField', [], {'max_length': '20', 'db_index': 'True'})
        },
        u'tracking.cat001tipo_contenedores': {
            'Meta': {'object_name': 'cat001tipo_contenedores'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'categoria': ('django.db.models.fields.CharField', [], {'max_length': '20', 'null': 'True', 'blank': 'True'}),
            'descripcion': ('django.db.models.fields.CharField', [], {'max_length': '150'}),
            'dimensiones': ('django.db.models.fields.CharField', [], {'max_length': '4', 'null': 'True', 'blank': 'True'}),
            'siglas': ('django.db.models.fields.CharField', [], {'max_length': '4', 'null': 'True', 'blank': 'True'}),
            'tipo_contenedor_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'})
        },
        u'tracking.cat001unidad_medida': {
            'Meta': {'object_name': 'cat001unidad_medida'},
            'abreviacion': ('django.db.models.fields.CharField', [], {'max_length': '10', 'null': 'True', 'blank': 'True'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'breviacion_ingles': ('django.db.models.fields.CharField', [], {'max_length': '10', 'null': 'True', 'blank': 'True'}),
            'clave': ('django.db.models.fields.CharField', [], {'max_length': '5', 'null': 'True', 'blank': 'True'}),
            'descripcion': ('django.db.models.fields.CharField', [], {'max_length': '30'}),
            'unidad_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'})
        },
        u'tracking.det001candados': {
            'Meta': {'object_name': 'det001candados'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'candado_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'contenedor': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.det001contenedores']"}),
            'numero': ('django.db.models.fields.CharField', [], {'max_length': '150'})
        },
        u'tracking.det001contenedores': {
            'Meta': {'object_name': 'det001contenedores'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'contenedor_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'numero_cotenedor': ('django.db.models.fields.CharField', [], {'max_length': '50', 'null': 'True'}),
            'tipo_contenedor': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001tipo_contenedores']"})
        },
        u'tracking.det001facturas': {
            'Meta': {'object_name': 'det001facturas'},
            'detalle_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'factura': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.reg001facturas']"})
        },
        u'tracking.det001formatocampos': {
            'Meta': {'unique_together': "(('formato', 'campos'),)", 'object_name': 'det001FormatoCampos'},
            'campo': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001CamposXls']"}),
            'campos': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'formato': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001FormatosXls']"}),
            'nombre_columna': ('django.db.models.fields.CharField', [], {'max_length': '150'})
        },
        u'tracking.det001guias': {
            'Meta': {'object_name': 'det001guias'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'guia_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'numero_guia': ('django.db.models.fields.CharField', [], {'max_length': '50'}),
            'operacion': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.reg001operacion']"}),
            'tipo': ('django.db.models.fields.CharField', [], {'default': "'ma'", 'max_length': '6'})
        },
        u'tracking.det001identificadores_partida': {
            'Meta': {'object_name': 'det001identificadores_partida'},
            'complemento_1': ('django.db.models.fields.CharField', [], {'max_length': '50', 'null': 'True', 'blank': 'True'}),
            'complemento_2': ('django.db.models.fields.CharField', [], {'max_length': '50', 'null': 'True', 'blank': 'True'}),
            'complemento_3': ('django.db.models.fields.CharField', [], {'max_length': '50', 'null': 'True', 'blank': 'True'}),
            'descripcion': ('django.db.models.fields.CharField', [], {'max_length': '100'}),
            'identificador': ('django.db.models.fields.CharField', [], {'max_length': '10', 'db_index': 'True'}),
            'identificador_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'partida': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.det001partidas']"})
        },
        u'tracking.det001identificadores_pedimento': {
            'Meta': {'object_name': 'det001identificadores_pedimento'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'clave_identificador': ('django.db.models.fields.CharField', [], {'max_length': '5'}),
            'descripcion': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'}),
            'identificador_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'operacion': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.reg001operacion']"}),
            'permiso': ('django.db.models.fields.CharField', [], {'max_length': '50', 'null': 'True', 'blank': 'True'})
        },
        u'tracking.det001impuestos_pedimento': {
            'Meta': {'object_name': 'det001impuestos_pedimento'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'impuesto': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001impuestos']"}),
            'impuestos_pedimento_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'operacion': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.reg001operacion']"})
        },
        u'tracking.det001partidas': {
            'Meta': {'object_name': 'det001partidas'},
            'cantidad_comercial': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'cantidad_tarifa': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'cc': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'cuota_operacion': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'detalle_mercancia': ('django.db.models.fields.TextField', [], {}),
            'dta_partida': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'factor_actualizacion': ('django.db.models.fields.DecimalField', [], {'max_digits': '6', 'decimal_places': '4'}),
            'fraccion': ('django.db.models.fields.CharField', [], {'max_length': '8'}),
            'ieps': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'igie': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'isan': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'iva': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'metodo_valoracion': ('django.db.models.fields.SmallIntegerField', [], {}),
            'numero_partida': ('django.db.models.fields.SmallIntegerField', [], {}),
            'numeros_serie': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'}),
            'observaciones': ('django.db.models.fields.TextField', [], {'null': 'True', 'blank': 'True'}),
            'operacion': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.reg001operacion']"}),
            'partida_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'precio_unitario': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'recargos': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'tasa_cc': ('django.db.models.fields.DecimalField', [], {'max_digits': '6', 'decimal_places': '3'}),
            'tasa_ieps': ('django.db.models.fields.DecimalField', [], {'max_digits': '6', 'decimal_places': '3'}),
            'tasa_igie': ('django.db.models.fields.DecimalField', [], {'max_digits': '6', 'decimal_places': '3'}),
            'tasa_isan': ('django.db.models.fields.DecimalField', [], {'max_digits': '6', 'decimal_places': '3'}),
            'tasa_iva': ('django.db.models.fields.DecimalField', [], {'max_digits': '6', 'decimal_places': '3'}),
            'tasa_max': ('django.db.models.fields.DecimalField', [], {'max_digits': '6', 'decimal_places': '3'}),
            'um_comercial': ('django.db.models.fields.SmallIntegerField', [], {}),
            'um_tarifa': ('django.db.models.fields.SmallIntegerField', [], {}),
            'valor_aduana': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'valor_dls': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'valor_mercancia': ('django.db.models.fields.DecimalField', [], {'max_digits': '32', 'decimal_places': '6'}),
            'vinculacion': ('django.db.models.fields.SmallIntegerField', [], {})
        },
        u'tracking.reg001facturas': {
            'Meta': {'object_name': 'reg001facturas'},
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'direccion': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001direcciones']"}),
            'edocument': ('django.db.models.fields.CharField', [], {'max_length': '25', 'null': 'True', 'blank': 'True'}),
            'factor_moneda': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '12', 'decimal_places': '6', 'blank': 'True'}),
            'factura_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'fecha_factura': ('django.db.models.fields.DateField', [], {}),
            'folio': ('django.db.models.fields.CharField', [], {'max_length': '30', 'null': 'True', 'blank': 'True'}),
            'incoterm': ('django.db.models.fields.CharField', [], {'max_length': '3', 'null': 'True', 'blank': 'True'}),
            'moneda_factura': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001monedas']"}),
            'numero_operacion': ('django.db.models.fields.CharField', [], {'max_length': '25', 'null': 'True', 'blank': 'True'}),
            'operacion': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.reg001operacion']", 'null': 'True', 'blank': 'True'}),
            'pais_factura': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001paises']"}),
            'proveedor': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001proveedores']"}),
            'serie': ('django.db.models.fields.CharField', [], {'max_length': '30', 'null': 'True', 'blank': 'True'}),
            'valor_dls': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '6', 'blank': 'True'}),
            'valor_monex': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '6', 'blank': 'True'})
        },
        u'tracking.reg001operacion': {
            'Meta': {'object_name': 'reg001operacion'},
            'aduana': ('django.db.models.fields.CharField', [], {'max_length': '2'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'buque': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001buques']"}),
            'clave_cliente': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.rel001claves_cliente']"}),
            'clave_embarque': ('django.db.models.fields.IntegerField', [], {'null': 'True', 'blank': 'True'}),
            'clave_pedimento': ('django.db.models.fields.CharField', [], {'max_length': '2', 'db_index': 'True'}),
            'contenedores_embarque': ('django.db.models.fields.IntegerField', [], {'max_length': '4'}),
            'contenedores_pedimento': ('django.db.models.fields.IntegerField', [], {'max_length': '4'}),
            'direccion_cliente': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001direcciones']"}),
            'estatus': ('django.db.models.fields.CharField', [], {'default': "'A'", 'max_length': '1', 'db_index': 'True', 'blank': 'True'}),
            'fecha_arribo': ('django.db.models.fields.DateField', [], {'null': 'True', 'blank': 'True'}),
            'fecha_bl': ('django.db.models.fields.DateField', [], {'null': 'True', 'blank': 'True'}),
            'fecha_despacho': ('django.db.models.fields.DateField', [], {'null': 'True', 'blank': 'True'}),
            'fecha_entrada': ('django.db.models.fields.DateField', [], {'null': 'True', 'blank': 'True'}),
            'fecha_pago': ('django.db.models.fields.DateField', [], {'null': 'True', 'blank': 'True'}),
            'fecha_revalidacion': ('django.db.models.fields.DateField', [], {'null': 'True', 'blank': 'True'}),
            'firma': ('django.db.models.fields.CharField', [], {'max_length': '50', 'null': 'True', 'blank': 'True'}),
            'firmado': ('django.db.models.fields.BooleanField', [], {'default': 'False'}),
            'fletes': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '4', 'blank': 'True'}),
            'linea_aerea': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001linea_aerea']"}),
            'oficina': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001oficinas']"}),
            'operacion_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'otros_incrementables': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '4', 'blank': 'True'}),
            'pais_cliente': ('django.db.models.fields.related.ForeignKey', [], {'related_name': "'++'", 'to': u"orm['tracking.cat001paises']"}),
            'pais_origen': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001paises']"}),
            'pais_procedencia': ('django.db.models.fields.related.ForeignKey', [], {'related_name': "'+'", 'to': u"orm['tracking.cat001paises']"}),
            'patente': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001patentes']"}),
            'pedimento': ('django.db.models.fields.CharField', [], {'max_length': '10'}),
            'peso_bruto': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '4', 'blank': 'True'}),
            'puerto_embarque': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001puertos']"}),
            'referencia': ('django.db.models.fields.CharField', [], {'max_length': '20', 'db_index': 'True'}),
            'seccion': ('django.db.models.fields.CharField', [], {'max_length': '1'}),
            'seguros': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '4', 'blank': 'True'}),
            'tipo': ('django.db.models.fields.CharField', [], {'max_length': '4'}),
            'tipo_cambio': ('django.db.models.fields.DecimalField', [], {'max_digits': '12', 'decimal_places': '6'}),
            'total_bultos': ('django.db.models.fields.CharField', [], {'max_length': '150', 'null': 'True', 'blank': 'True'}),
            'valor_aduana': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '4', 'blank': 'True'}),
            'valor_dls_factura': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '4', 'blank': 'True'}),
            'valor_monex_factura': ('django.db.models.fields.DecimalField', [], {'null': 'True', 'max_digits': '32', 'decimal_places': '4', 'blank': 'True'})
        },
        u'tracking.rel001claves_cliente': {
            'Meta': {'object_name': 'rel001claves_cliente'},
            'activo': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'clave_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'descripcion': ('django.db.models.fields.CharField', [], {'max_length': '255'}),
            'numero_clave': ('django.db.models.fields.CharField', [], {'max_length': '15', 'db_index': 'True'}),
            'oficina': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001oficinas']"}),
            'razon_social': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001razones_sociales']"})
        },
        u'tracking.rel001direcciones_clientes': {
            'Meta': {'object_name': 'rel001direcciones_clientes'},
            'activo': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'cliente': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.rel001claves_cliente']"}),
            'direccion': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001direcciones']"}),
            'direccion_cliente_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'})
        },
        u'tracking.rel001direcciones_proveedor': {
            'Meta': {'object_name': 'rel001direcciones_proveedor'},
            'activo': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'borrado': ('django.db.models.fields.BooleanField', [], {'default': 'False', 'db_index': 'True'}),
            'direccion': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001direcciones']"}),
            'direccion_proveedor_id': ('django.db.models.fields.AutoField', [], {'primary_key': 'True'}),
            'proveedor': ('django.db.models.fields.related.ForeignKey', [], {'to': u"orm['tracking.cat001proveedores']"})
        }
    }

    complete_apps = ['tracking']