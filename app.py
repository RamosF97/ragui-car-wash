# RAGUI CAR WASH - AUTOLAVADO WEB
# Usuario: admin | Clave: tula123 (cámbiala en la línea de abajo)

import json, os
from datetime import datetime, timedelta
from collections import Counter
from functools import wraps
from flask import Flask, render_template_string, request, redirect, send_file, Response
import io

USUARIO = "admin"
CLAVE = "Gauss@12A"

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    TIENE_EXCEL = True
except:
    TIENE_EXCEL = False

app = Flask(__name__)
DATA_DIR = os.environ.get("DATA_DIR", ".")
os.makedirs(DATA_DIR, exist_ok=True)
ARCHIVO_VENTAS = os.path.join(DATA_DIR, "ventas.json")
ARCHIVO_CLIENTES = os.path.join(DATA_DIR, "clientes.json")
ARCHIVO_PRECIOS = os.path.join(DATA_DIR, "precios.json")
ARCHIVO_EMPLEADOS = os.path.join(DATA_DIR, "empleados.json")

PRECIOS_DEFECTO = {
    "Lavado Básico - Auto": 120,
    "Lavado Básico - Camioneta / SUV": 150,
    "Lavado Completo - Auto": 220,
    "Lavado Completo - Camioneta / SUV": 280,
    "Lavado Premium + Encerado": 350,
    "Lavado de Motor": 180,
    "Aspirado Interior": 80,
    "Pulido de Faros": 750,
}
EMPLEADOS_DEFECTO = {
    "Juan Pérez": {"telefono": "773-123-4567", "puesto": "Lavador", "comision_tipo": "Fija", "comision_valor": 20, "activo": True},
    "Carlos López": {"telefono": "773-987-6543", "puesto": "Lavador", "comision_tipo": "Fija", "comision_valor": 20, "activo": True},
    "María García": {"telefono": "773-555-1212", "puesto": "Cajero", "comision_tipo": "Porcentaje", "comision_valor": 5, "activo": True},
}

def cargar(archivo, defecto):
    if not os.path.exists(archivo):
        guardar(archivo, defecto); return defecto
    try:
        with open(archivo, "r", encoding="utf-8") as f: return json.load(f)
    except: return defecto
def guardar(archivo, datos):
    with open(archivo, "w", encoding="utf-8") as f: json.dump(datos, f, indent=4, ensure_ascii=False)
def calcular_comision(info, precio):
    if not info: return 0
    if info.get("comision_tipo")=="Porcentaje": return round(precio*info.get("comision_valor",0)/100,2)
    return float(info.get("comision_valor",0))

def requiere_login(f):
    @wraps(f)
    def decorada(*args, **kwargs):
        auth = request.authorization
        if not auth or auth.username != USUARIO or auth.password != CLAVE:
            return Response("Necesitas iniciar sesión para entrar a RAGUI Car Wash.", 401, {"WWW-Authenticate": 'Basic realm="RAGUI Car Wash"'})
        return f(*args, **kwargs)
    return decorada

BASE = """
<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>RAGUI Car Wash - Autolavado</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
body{background:#f0f2f5; padding-bottom:80px;}
.header{background:#0b2a4a; color:white; padding:14px; text-align:center; position:sticky; top:0; z-index:10;}
.header h1{font-size:22px; margin:0; font-weight:900; letter-spacing:1px;}
.header small{color:#7dd3fc;}
.card{border-radius:14px; border:none; box-shadow:0 2px 8px rgba(0,0,0,.08);}
.btn-lg{padding:14px; font-size:17px; font-weight:bold; border-radius:12px;}
input,select{font-size:16px !important; padding:12px !important; border-radius:10px !important;}
.nav-bottom{position:fixed; bottom:0; left:0; right:0; background:white; display:flex; border-top:1px solid #ddd; z-index:20;}
.nav-bottom a{flex:1; text-align:center; padding:10px 2px; text-decoration:none; color:#333; font-size:12px; font-weight:bold;}
.nav-bottom a.active{color:#0b2a4a; border-top:3px solid #0b2a4a;}
@media(min-width:768px){ .container{max-width:900px;} }
</style></head><body>
<div class="header"><h1>🚗 RAGUI CAR WASH</h1><div style="font-size:12px; letter-spacing:3px; font-weight:bold;">AUTOLAVADO</div><small>{{ total_hoy }}</small></div>
<div class="container mt-3">{{ contenido|safe }}</div>
<div class="nav-bottom">
<a href="/" class="{{ 'active' if tab=='venta' else '' }}">💰<br>Venta</a>
<a href="/clientes" class="{{ 'active' if tab=='clientes' else '' }}">⭐<br>Clientes</a>
<a href="/empleados" class="{{ 'active' if tab=='empleados' else '' }}">👷<br>Empleados</a>
<a href="/historial" class="{{ 'active' if tab=='historial' else '' }}">📋<br>Historial</a>
<a href="/corte" class="{{ 'active' if tab=='corte' else '' }}">🧾<br>Corte/Excel</a>
</div>
<script>
function buscarCliente(){
  let placa=document.getElementById('placa').value.toUpperCase();
  document.getElementById('placa').value=placa;
  if(placa.length<3) return;
  fetch('/api/cliente/'+placa).then(r=>r.json()).then(d=>{
    if(d.existe){
      document.getElementById('cliente').value=d.nombre;
      document.getElementById('telefono').value=d.telefono||'';
      document.getElementById('vehiculo').value=d.vehiculo||'';
      document.getElementById('info-puntos').innerHTML='⭐ '+d.puntos+' puntos | '+d.visitas+' visitas | Faltan '+(100-d.puntos>0?100-d.puntos:0)+' para premio';
      document.getElementById('info-puntos').style.display='block';
      document.getElementById('canje-box').style.display = d.puntos>=100 ? 'block' : 'none';
    } else {
      document.getElementById('info-puntos').innerHTML='Cliente nuevo - 0 puntos';
      document.getElementById('canje-box').style.display='none';
    }
  });
}
function actualizarPrecio(){
  let sel=document.getElementById('servicio'); let precio=parseFloat(sel.options[sel.selectedIndex].dataset.precio);
  let canje=document.getElementById('canje')?.checked;
  document.getElementById('total-label').innerText = canje ? 'TOTAL: $0.00 (GRATIS con puntos)' : 'TOTAL: $'+precio.toFixed(2);
  document.getElementById('puntos-gana').innerText = canje ? 'Se descontarán 100 puntos' : 'Ganará +'+Math.floor(precio/10)+' puntos';
}
</script></body></html>
"""

def total_hoy_txt():
    ventas=cargar(ARCHIVO_VENTAS,[])
    hoy=datetime.now().strftime("%Y-%m-%d")
    del_dia=[v for v in ventas if v['fecha']==hoy]
    total=sum(v['precio'] for v in del_dia); com=sum(v.get("comision",0) for v in del_dia)
    return f"Hoy: {len(del_dia)} autos | Ventas ${total:.2f} | Neto ${total-com:.2f}"
def render(contenido, tab):
    return render_template_string(BASE, contenido=contenido, tab=tab, total_hoy=total_hoy_txt())

@app.route("/", methods=["GET","POST"])
@requiere_login
def venta():
    ventas=cargar(ARCHIVO_VENTAS,[]); clientes=cargar(ARCHIVO_CLIENTES,{}); precios=cargar(ARCHIVO_PRECIOS,PRECIOS_DEFECTO); empleados=cargar(ARCHIVO_EMPLEADOS,EMPLEADOS_DEFECTO)
    mensaje=""
    if request.method=="POST":
        placa=request.form.get("placa","").upper().strip()
        nombre=request.form.get("cliente","").strip() or "Cliente General"
        telefono=request.form.get("telefono","").strip(); vehiculo=request.form.get("vehiculo","").strip() or "No especificado"
        servicio=request.form.get("servicio"); pago=request.form.get("pago"); empleado=request.form.get("empleado")
        es_canje = "canje" in request.form
        precio_orig=precios.get(servicio,0)
        if es_canje and clientes.get(placa,{}).get("puntos",0)<100:
            mensaje='<div class="alert alert-danger">❌ No tiene 100 puntos.</div>'
        elif placa:
            total=0 if es_canje else precio_orig; pts_ganados=0 if es_canje else int(total//10); pts_usados=100 if es_canje else 0
            comision=10.0 if es_canje else calcular_comision(empleados.get(empleado), total)
            if placa not in clientes: clientes[placa]={"nombre":nombre,"telefono":telefono,"vehiculo":vehiculo,"puntos":0,"visitas":0,"total_gastado":0}
            c=clientes[placa]; c["nombre"]=nombre; c["telefono"]=telefono; c["vehiculo"]=vehiculo
            c["puntos"]=c.get("puntos",0)-pts_usados+pts_ganados; c["visitas"]=c.get("visitas",0)+1; c["total_gastado"]=c.get("total_gastado",0)+total; c["ultima_visita"]=datetime.now().strftime("%Y-%m-%d %H:%M")
            ahora=datetime.now()
            ventas.append({"fecha":ahora.strftime("%Y-%m-%d"),"hora":ahora.strftime("%H:%M:%S"),"placa":placa,"cliente":nombre,"vehiculo":vehiculo,"servicio":servicio,"precio":total,"precio_original":precio_orig,"forma_pago":pago,"puntos_ganados":pts_ganados,"canje":es_canje,"empleado":empleado,"comision":comision})
            guardar(ARCHIVO_VENTAS,ventas); guardar(ARCHIVO_CLIENTES,clientes)
            mensaje=f'<div class="alert alert-success">✅ <b>Venta lista en RAGUI Car Wash!</b> {nombre} ({placa}) - {servicio} - ${total:.2f} - {empleado} - Comisión ${comision:.2f} - Puntos: {c["puntos"]}</div>'
    opciones_serv="".join([f'<option value="{k}" data-precio="{v}">{k} - ${v:.2f}</option>' for k,v in precios.items()])
    opciones_emp="".join([f'<option value="{n}">{n} - {d.get("puesto")}</option>' for n,d in empleados.items() if d.get("activo",True)])
    html=f"""
    {mensaje}
    <div class="card p-3"><h4>💰 Nueva Venta - RAGUI</h4><form method="POST">
    <label class="fw-bold">PLACA *</label><input id="placa" name="placa" class="form-control mb-2" placeholder="Ej: HJK-123-A" oninput="buscarCliente()" required>
    <div id="info-puntos" class="alert alert-warning" style="display:none"></div>
    <label class="fw-bold">Nombre</label><input id="cliente" name="cliente" class="form-control mb-2" placeholder="Nombre del cliente">
    <label class="fw-bold">Teléfono</label><input id="telefono" name="telefono" class="form-control mb-2" placeholder="WhatsApp">
    <label class="fw-bold">Vehículo</label><input id="vehiculo" name="vehiculo" class="form-control mb-2" placeholder="Ej: Tsuru blanco">
    <label class="fw-bold">Servicio</label><select id="servicio" name="servicio" class="form-select mb-2" onchange="actualizarPrecio()">{opciones_serv}</select>
    <label class="fw-bold">Forma de Pago</label><select name="pago" class="form-select mb-2"><option>Efectivo</option><option>Tarjeta</option><option>Transferencia</option></select>
    <label class="fw-bold">👷 ¿Quién lavó?</label><select name="empleado" class="form-select mb-3">{opciones_emp}</select>
    <div id="canje-box" class="alert alert-warning" style="display:none"><label><input type="checkbox" id="canje" name="canje" onchange="actualizarPrecio()"> ⭐ Canjear 100 puntos por lavado GRATIS</label></div>
    <div class="card bg-light p-3 text-center mb-3"><h3 id="total-label">TOTAL: ${list(precios.values())[0]:.2f}</h3><span id="puntos-gana" class="text-success fw-bold"></span></div>
    <button class="btn btn-success btn-lg w-100">✅ COBRAR VENTA</button></form></div>
    <script>actualizarPrecio()</script>"""
    return render(html, "venta")

@app.route("/api/cliente/<placa>")
@requiere_login
def api_cliente(placa):
    clientes=cargar(ARCHIVO_CLIENTES,{})
    c=clientes.get(placa.upper())
    if c: return {"existe":True, **c}
    return {"existe":False}

@app.route("/clientes")
@requiere_login
def clientes():
    clientes=cargar(ARCHIVO_CLIENTES,{}); q=request.args.get("q","").upper()
    filas=""
    for placa,c in sorted(clientes.items(), key=lambda x: x[1].get("puntos",0), reverse=True):
        if q and q not in placa and q not in c.get("nombre","").upper(): continue
        color="success" if c.get("puntos",0)>=100 else "secondary"
        filas+=f"<tr><td><b>{placa}</b></td><td>{c.get('nombre','')}</td><td>{c.get('telefono','')}</td><td><span class='badge bg-{color}'>{c.get('puntos',0)} pts</span></td><td>{c.get('visitas',0)}</td><td>${c.get('total_gastado',0):.2f}</td></tr>"
    html=f"""<div class="card p-3"><h4>⭐ Clientes Frecuentes - RAGUI</h4>
    <form><div class="input-group mb-3"><input name="q" value="{q}" class="form-control" placeholder="Buscar por placa o nombre"><button class="btn btn-primary">Buscar</button></div></form>
    <p class="text-muted small">1 punto = $10 | 100 puntos = Lavado GRATIS</p>
    <div class="table-responsive"><table class="table table-striped"><tr><th>Placa</th><th>Nombre</th><th>Tel</th><th>Puntos</th><th>Visitas</th><th>Total</th></tr>{filas}</table></div></div>"""
    return render(html, "clientes")

@app.route("/empleados", methods=["GET","POST"])
@requiere_login
def empleados():
    empleados=cargar(ARCHIVO_EMPLEADOS,EMPLEADOS_DEFECTO); ventas=cargar(ARCHIVO_VENTAS,[])
    if request.method=="POST":
        nombre=request.form.get("nombre","").strip()
        if nombre and nombre not in empleados:
            empleados[nombre]={"telefono":request.form.get("telefono",""),"puesto":request.form.get("puesto"),"comision_tipo":request.form.get("tipo"),"comision_valor":float(request.form.get("valor",20)),"activo":True}
            guardar(ARCHIVO_EMPLEADOS,empleados); return redirect("/empleados")
    hoy=datetime.now().strftime("%Y-%m-%d"); filas=""
    for n,d in empleados.items():
        ve=[v for v in ventas if v.get("empleado")==n and v['fecha']==hoy]
        com_txt=f"${d.get('comision_valor')} fija" if d.get("comision_tipo")=="Fija" else f"{d.get('comision_valor')}%"
        filas+=f"<tr><td><b>{n}</b><br><small>{d.get('puesto')}</small></td><td>{com_txt}</td><td>{len(ve)}</td><td>${sum(v.get('comision',0) for v in ve):.2f}</td><td>${sum(v['precio'] for v in ve):.2f}</td><td>{'✅' if d.get('activo',True) else '❌'}</td></tr>"
    html=f"""<div class="card p-3 mb-3"><h4>👷 Empleados - RAGUI</h4>
    <form method="POST" class="row g-2"><div class="col-12"><input name="nombre" class="form-control" placeholder="Nombre del empleado" required></div>
    <div class="col-6"><input name="telefono" class="form-control" placeholder="Teléfono"></div>
    <div class="col-6"><select name="puesto" class="form-select"><option>Lavador</option><option>Cajero</option><option>Encargado</option></select></div>
    <div class="col-6"><select name="tipo" class="form-select"><option>Fija</option><option>Porcentaje</option></select></div>
    <div class="col-6"><input name="valor" type="number" value="20" class="form-control"></div>
    <div class="col-12"><button class="btn btn-success w-100">➕ Agregar Empleado</button></div></form></div>
    <div class="card p-3"><h5>Rendimiento de Hoy</h5><div class="table-responsive"><table class="table"><tr><th>Empleado</th><th>Comisión</th><th>Autos</th><th>Comisión Hoy</th><th>Vendido</th><th>Estado</th></tr>{filas}</table></div></div>"""
    return render(html, "empleados")

@app.route("/historial")
@requiere_login
def historial():
    ventas=cargar(ARCHIVO_VENTAS,[]); fecha=request.args.get("fecha", datetime.now().strftime("%Y-%m-%d"))
    filas=""
    for v in reversed([x for x in ventas if x['fecha']==fecha]):
        filas+=f"<tr><td>{v['hora']}</td><td><b>{v['placa']}</b></td><td>{v['cliente']}</td><td>{v['servicio']}</td><td>${v['precio']:.2f}</td><td>{v['forma_pago']}</td><td>{v.get('empleado','-')}</td><td>${v.get('comision',0):.2f}</td></tr>"
    html=f"""<div class="card p-3"><h4>📋 Historial - RAGUI</h4>
    <form class="row g-2 mb-3"><div class="col-8"><input type="date" name="fecha" value="{fecha}" class="form-control"></div><div class="col-4"><button class="btn btn-primary w-100">Filtrar</button></div></form>
    <div class="table-responsive"><table class="table table-striped"><tr><th>Hora</th><th>Placa</th><th>Cliente</th><th>Servicio</th><th>Precio</th><th>Pago</th><th>Empleado</th><th>Comisión</th></tr>{filas}</table></div></div>"""
    return render(html, "historial")

@app.route("/corte")
@requiere_login
def corte():
    ventas=cargar(ARCHIVO_VENTAS,[]); empleados=cargar(ARCHIVO_EMPLEADOS,EMPLEADOS_DEFECTO)
    hoy=datetime.now()
    inicio=request.args.get("inicio", (hoy - timedelta(days=hoy.weekday())).strftime("%Y-%m-%d"))
    fin=request.args.get("fin", hoy.strftime("%Y-%m-%d"))
    del_rango=[v for v in ventas if inicio <= v['fecha'] <= fin]
    efectivo=sum(v['precio'] for v in del_rango if v['forma_pago']=="Efectivo"); tarjeta=sum(v['precio'] for v in del_rango if v['forma_pago']=="Tarjeta"); transf=sum(v['precio'] for v in del_rango if v['forma_pago']=="Transferencia")
    total=efectivo+tarjeta+transf; com=sum(v.get("comision",0) for v in del_rango)
    filas_emp=""
    for n in empleados:
        ve=[v for v in del_rango if v.get("empleado")==n]
        if ve: filas_emp+=f"<tr><td>{n}</td><td>{len(ve)}</td><td>${sum(v['precio'] for v in ve):.2f}</td><td>${sum(v.get('comision',0) for v in ve):.2f}</td></tr>"
    mas_vendido=Counter(v['servicio'] for v in del_rango).most_common(1)
    html=f"""<div class="card p-3"><h4>🧾 Corte / Reporte - RAGUI</h4>
    <form class="row g-2 mb-3"><div class="col-6"><label>Desde</label><input type="date" name="inicio" value="{inicio}" class="form-control"></div>
    <div class="col-6"><label>Hasta</label><input type="date" name="fin" value="{fin}" class="form-control"></div>
    <div class="col-12"><button class="btn btn-primary w-100">🔍 Ver Corte Filtrado</button></div></form>
    <div class="row g-2 mb-3">
      <div class="col-6"><a href="/corte?inicio={hoy.strftime('%Y-%m-%d')}&fin={hoy.strftime('%Y-%m-%d')}" class="btn btn-success w-100">Hoy</a></div>
      <div class="col-6"><a href="/corte?inicio={(hoy - timedelta(days=hoy.weekday())).strftime('%Y-%m-%d')}&fin={(hoy + timedelta(days=6-hoy.weekday())).strftime('%Y-%m-%d')}" class="btn btn-outline-primary w-100">Esta Semana</a></div>
      <div class="col-6"><a href="/corte?inicio={(hoy - timedelta(days=hoy.weekday()+7)).strftime('%Y-%m-%d')}&fin={(hoy - timedelta(days=hoy.weekday()+1)).strftime('%Y-%m-%d')}" class="btn btn-outline-secondary w-100">Semana Pasada</a></div>
      <div class="col-6"><a href="/corte?inicio={hoy.strftime('%Y-%m-01')}&fin={hoy.strftime('%Y-%m-%d')}" class="btn btn-outline-warning w-100">Este Mes</a></div>
    </div>
    <div class="alert alert-info"><b>Del {inicio} al {fin}</b><br>Servicios: {len(del_rango)} | Efectivo: ${efectivo:.2f} | Tarjeta: ${tarjeta:.2f} | Transferencia: ${transf:.2f}<br><b>Total: ${total:.2f} | Comisiones: ${com:.2f} | Neto: ${total-com:.2f}</b><br>Servicio top: {mas_vendido[0][0] if mas_vendido else '-'} ({mas_vendido[0][1] if mas_vendido else 0})</div>
    <table class="table"><tr><th>Empleado</th><th>Autos</th><th>Vendido</th><th>Comisión</th></tr>{filas_emp}</table>
    <a href="/excel?inicio={inicio}&fin={fin}" class="btn btn-dark btn-lg w-100">📊 DESCARGAR EXCEL DE ESTE RANGO</a></div>"""
    return render(html, "corte")

@app.route("/excel")
@requiere_login
def excel():
    if not TIENE_EXCEL: return "Instala openpyxl"
    ventas=cargar(ARCHIVO_VENTAS,[]); empleados=cargar(ARCHIVO_EMPLEADOS,EMPLEADOS_DEFECTO)
    inicio=request.args.get("inicio", datetime.now().strftime("%Y-%m-%d")); fin=request.args.get("fin", inicio)
    wb=Workbook(); ws=wb.active; ws.title="Resumen"
    ws.merge_cells("A1:F1"); ws["A1"]=f"RAGUI CAR WASH - Del {inicio} al {fin}"; ws["A1"].font=Font(bold=True,size=14,color="0B2A4A")
    ws.append([]); ws.append(["Fecha","Día","Autos","Ventas","Comisiones","Neto"])
    estilo=PatternFill(start_color="0B2A4A",end_color="0B2A4A",fill_type="solid")
    for c in ws[3]: c.fill=estilo; c.font=Font(bold=True,color="FFFFFF")
    dias=["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
    d=datetime.strptime(inicio,"%Y-%m-%d"); fin_d=datetime.strptime(fin,"%Y-%m-%d"); tot_a=tot_v=tot_c=0
    while d<=fin_d:
        f=d.strftime("%Y-%m-%d"); dd=[v for v in ventas if v['fecha']==f]; vt=sum(v['precio'] for v in dd); cm=sum(v.get("comision",0) for v in dd)
        ws.append([f,dias[d.weekday()],len(dd),vt,cm,vt-cm]); tot_a+=len(dd); tot_v+=vt; tot_c+=cm; d+=timedelta(days=1)
    ws.append(["TOTAL","",tot_a,tot_v,tot_c,tot_v-tot_c])
    ws2=wb.create_sheet("Ventas"); ws2.append(["Fecha","Hora","Placa","Cliente","Servicio","Precio","Pago","Empleado","Comisión","Puntos"])
    for c in ws2[1]: c.fill=estilo; c.font=Font(bold=True,color="FFFFFF")
    for v in sorted([x for x in ventas if inicio<=x['fecha']<=fin], key=lambda x:(x['fecha'],x['hora'])):
        ws2.append([v['fecha'],v['hora'],v['placa'],v['cliente'],v['servicio'],v['precio'],v['forma_pago'],v.get("empleado","-"),v.get("comision",0),v.get("puntos_ganados",0)])
    ws3=wb.create_sheet("Por Empleado"); ws3.append(["Empleado","Autos","Vendido","Comisión"])
    for c in ws3[1]: c.fill=estilo; c.font=Font(bold=True,color="FFFFFF")
    for n in empleados:
        ve=[v for v in ventas if inicio<=v['fecha']<=fin and v.get("empleado")==n]
        if ve: ws3.append([n,len(ve),sum(v['precio'] for v in ve),sum(v.get("comision",0) for v in ve)])
    bio=io.BytesIO(); wb.save(bio); bio.seek(0)
    return send_file(bio, as_attachment=True, download_name=f"RAGUI_Reporte_{inicio}_a_{fin}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/health")
def health(): return "OK - RAGUI Car Wash"

if __name__=="__main__":
    port=int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
