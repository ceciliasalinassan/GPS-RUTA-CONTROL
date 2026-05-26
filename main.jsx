
import React,{useEffect,useMemo,useState}from"react";
import{createRoot}from"react-dom/client";
import{AlertTriangle,Bell,CalendarDays,CheckCircle,Clock,CreditCard,Edit,Eye,FileText,Lock,LogOut,Mail,Paperclip,Plus,Search,ShieldCheck,Trash2,TrendingDown,TrendingUp,UploadCloud,User,Users,Wallet,Save,Building2,Download,Upload,HardDrive}from"lucide-react";
import{BarChart,Bar,XAxis,YAxis,Tooltip,ResponsiveContainer,CartesianGrid,PieChart,Pie,Cell}from"recharts";
import * as XLSX from "xlsx";
import"./style.css";

const KEY="gpsruta_financiero_pro_v2";
const SESSION="gpsruta_login";
const PASS="1234";
const incomeCats=["Venta de GPS","Instalación de GPS","Servicio mensual","Pago de factura","Abono cliente","Otro ingreso"];
const expenseCats=["Pago instalador","Pago IVA","Pago contadora","Pago arriendo","Pago combustible","Pago internet","Pago plataforma","Compra de equipos","Otro egreso"];
const seed={clients:[
{id:1,nombre:"Transportes del Sur SpA",rut:"76.543.210-9",giro:"Transporte de carga",telefono:"56912345678",email:"contacto@delsur.cl",direccion:"Santiago",contacto:"Juan Pérez"},
{id:2,nombre:"Constructora Andes Ltda.",rut:"77.555.333-1",giro:"Construcción",telefono:"56998765432",email:"pagos@andes.cl",direccion:"San Carlos",contacto:"María Torres"}],
invoices:[
{id:1,clienteId:1,factura:"FAC-2026-001",emision:"2026-05-01",vencimiento:"2026-05-18",monto:1250000,estado:"Pendiente",detalle:"Servicio GPS mensual"},
{id:2,clienteId:2,factura:"FAC-2026-002",emision:"2026-04-20",vencimiento:"2026-05-10",monto:2850000,estado:"Vencida",detalle:"Instalación y monitoreo"},
{id:3,clienteId:1,factura:"FAC-2026-003",emision:"2026-05-04",vencimiento:"2026-05-26",monto:950000,estado:"Pagada",detalle:"Mantención plataforma"}],
incomes:[{id:1,fecha:"2026-05-10",categoria:"Pago de factura",descripcion:"Pago FAC-2026-003",monto:950000,facturaId:3}],
expenses:[{id:1,fecha:"2026-05-11",categoria:"Pago plataforma",descripcion:"Servidor",monto:180000},{id:2,fecha:"2026-05-13",categoria:"Pago contadora",descripcion:"Honorarios contables",monto:75000}],
debts:[{id:1,fecha:"2026-05-15",proveedor:"Proveedor GPS",categoria:"Compra de equipos",descripcion:"Equipos GPS por pagar",monto:1200000,vencimiento:"2026-05-30",estado:"Pendiente"}],
attachments:{}};

const money=v=>new Intl.NumberFormat("es-CL",{style:"currency",currency:"CLP",maximumFractionDigits:0}).format(+v||0);
const today=()=>new Date().toISOString().slice(0,10);
const mk=d=>(d||today()).slice(0,7);
const ml=k=>{let[a,b]=k.split("-");return `${b}/${a}`};
function load(){try{let s=JSON.parse(localStorage.getItem(KEY));if(!s)return seed;return{...seed,...s,clients:(s.clients||seed.clients).map(c=>({giro:"",...c})),incomes:s.incomes||seed.incomes,expenses:s.expenses||seed.expenses,debts:s.debts||seed.debts}}catch{return seed}}
function days(d){let n=new Date();n.setHours(0,0,0,0);return Math.ceil((new Date(d+"T00:00:00")-n)/86400000)}
function ist(i){if(i.estado==="Pagada")return{l:"Pagada",c:"ok",I:CheckCircle};let d=days(i.vencimiento);if(i.estado==="Vencida"||d<0)return{l:"Vencida",c:"bad",I:AlertTriangle};if(d<=3&&d>=0)return{l:"Por vencer",c:"warn",I:Clock};return{l:"Pendiente",c:"soft",I:FileText}}
function wa(i,c){let t=`BUENOS DIAS ESTIMADOS,\n\nSE ADJUNTA FACTURA DEL MES.\n\nFactura: ${i.factura}\nMonto: ${money(i.monto)}\nVencimiento: ${i.vencimiento}\n\nSALUDOS CORDIALES.\nGPSRUTA.`;return`https://wa.me/${c?.telefono||""}?text=${encodeURIComponent(t)}`}
function em(i,c){let s=`Factura del mes - ${i.factura}`,b=`BUENOS DIAS ESTIMADOS,\n\nSE ADJUNTA FACTURA DEL MES.\n\nFactura: ${i.factura}\nMonto: ${money(i.monto)}\nVencimiento: ${i.vencimiento}\n\nSALUDOS CORDIALES.\nGPSRUTA.\n\nNota: adjuntar factura antes de enviar.`;return`mailto:${c?.email||""}?subject=${encodeURIComponent(s)}&body=${encodeURIComponent(b)}`}
function W(){return <svg viewBox="0 0 32 32" width="18" height="18"><path fill="currentColor" d="M16.04 3C8.86 3 3 8.84 3 16.02c0 2.3.6 4.54 1.75 6.51L3 29l6.64-1.7a12.95 12.95 0 0 0 6.4 1.64h.01C23.22 28.94 29 23.1 29 15.92 29 8.8 23.18 3 16.04 3Zm7.56 18.45c-.32.9-1.86 1.7-2.6 1.8-.67.1-1.52.14-2.45-.15-.56-.18-1.28-.42-2.2-.82-3.87-1.68-6.4-5.6-6.6-5.86-.2-.26-1.58-2.1-1.58-4s1-2.84 1.35-3.23c.36-.4.78-.5 1.04-.5h.75c.24.01.57-.09.9.69.32.78 1.1 2.69 1.2 2.89.1.2.16.43.03.69-.13.26-.2.42-.4.64-.2.23-.42.5-.6.67-.2.2-.4.42-.17.82.23.4 1.02 1.68 2.2 2.72 1.51 1.35 2.78 1.77 3.18 1.97.4.2.63.17.86-.1.23-.26 1-1.16 1.26-1.56.26-.4.53-.33.9-.2.36.13 2.3 1.08 2.7 1.28.4.2.66.3.76.46.1.16.1.95-.22 1.85Z"/></svg>}
function Logo(){return <div className="logo"><div className="pin">⌖</div><div><h1><span>GPS</span><b>ruta</b><small>.cl</small></h1><p>SEGUIMIENTO Y SEGURIDAD</p></div></div>}
function Login({onLogin}){const[p,setP]=useState(""),[e,setE]=useState("");return <div className="loginPage"><form className="loginCard" onSubmit={x=>{x.preventDefault();if(p===PASS){sessionStorage.setItem(SESSION,"1");onLogin()}else setE("Clave incorrecta. Clave demo: 1234")}}><Logo/><h2>Ingreso Seguro</h2><p>Sistema financiero y cobranza</p><div className="loginInput"><Lock size={18}/><input type="password" value={p} onChange={x=>setP(x.target.value)} placeholder="Clave de acceso"/></div>{e&&<div className="error">{e}</div>}<button className="primary full"><ShieldCheck size={18}/>Ingresar</button></form></div>}
function K({t,v,s,icon:Icon,tone="green"}){return <div className="card kpi"><div className={`kpiIcon ${tone}`}><Icon size={32}/></div><div><small>{t}</small><h3>{v}</h3><p>{s}</p></div></div>}
function Fields({obj,set,fields}){return <div className="formGrid">{fields.map(f=><input key={f} value={obj[f]||""} onChange={e=>set({...obj,[f]:e.target.value})} placeholder={f.toUpperCase()} type={["fecha","emision","vencimiento"].includes(f)?"date":f==="monto"?"number":"text"}/>)}</div>}
function InvTable({items,client,edit,del,data={},attachFile=()=>{},canSendReminder=()=>true}){return <div className="tableWrap"><table><thead><tr><th>Factura</th><th>Cliente</th><th>Vence</th><th>Mes/Año</th><th>Monto</th><th>Estado</th><th>Acciones</th></tr></thead><tbody>{items.map(i=>{let c=client(i.clienteId),s=ist(i),Icon=s.I;return <tr key={i.id}><td><b>{i.factura}</b></td><td>{c?.nombre}</td><td>{i.vencimiento}</td><td>{ml(mk(i.vencimiento))}</td><td>{money(i.monto)}</td><td><span className={`status ${s.c}`}><Icon size={14}/>{s.l}</span></td><td><div className="actions"><label className="icon attachMini" title="Adjuntar factura"><Paperclip size={17}/><input type="file" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.xls,.xlsx" onChange={e=>attachFile(i.id,e.target.files?.[0])}/></label><a className="icon whatsapp" href={data.attachments?.[i.id]?wa(i,c):"#"} onClick={e=>{if(!canSendReminder(i.id))e.preventDefault()}} target="_blank"><W/></a><a className="icon mail" href={data.attachments?.[i.id]?em(i,c):"#"} onClick={e=>{if(!canSendReminder(i.id))e.preventDefault()}}><Mail size={17}/></a><button className="icon edit" onClick={()=>edit(i)}><Edit size={17}/></button><button className="icon trash" onClick={()=>del(i.id)}><Trash2 size={17}/></button>{data.attachments?.[i.id]&&<span className="attachedOk">Adjunta</span>}</div></td></tr>})}</tbody></table></div>}
function App(){const[logged,setLogged]=useState(()=>sessionStorage.getItem(SESSION)==="1"),[data,setData]=useState(load),[tab,setTab]=useState("dashboard"),[clock,setClock]=useState(new Date()),[search,setSearch]=useState(""),[alertSearch,setAlertSearch]=useState(""),[saved,setSaved]=useState("Sin cambios"),[selectedInvoiceId,setSelectedInvoiceId]=useState(null),[selectedMonth,setSelectedMonth]=useState(mk(today()));
const[clientForm,setClientForm]=useState({nombre:"",rut:"",giro:"",telefono:"569",email:"",direccion:"",contacto:""}),[invoiceForm,setInvoiceForm]=useState({clienteId:"",factura:"",emision:today(),vencimiento:today(),monto:"",estado:"Pendiente",detalle:""}),[incomeForm,setIncomeForm]=useState({fecha:today(),categoria:"Pago de factura",descripcion:"",monto:"",facturaId:""}),[expenseForm,setExpenseForm]=useState({fecha:today(),categoria:"Pago instalador",descripcion:"",monto:"",debtId:""}),[debtForm,setDebtForm]=useState({fecha:today(),proveedor:"",categoria:"Compra de equipos",descripcion:"",monto:"",vencimiento:today(),estado:"Pendiente"}),[editingClient,setEditingClient]=useState(null),[editingInvoice,setEditingInvoice]=useState(null);
useEffect(()=>{localStorage.setItem(KEY,JSON.stringify(data));setSaved(new Date().toLocaleTimeString("es-CL"))},[data]);
useEffect(()=>{let t=setInterval(()=>setClock(new Date()),1000);return()=>clearInterval(t)},[]);
const client=id=>data.clients.find(c=>+c.id===+id);
const months=useMemo(()=>{let s=new Set([selectedMonth,mk(today())]);data.invoices.forEach(i=>s.add(mk(i.vencimiento)));data.incomes.forEach(i=>s.add(mk(i.fecha)));data.expenses.forEach(e=>s.add(mk(e.fecha)));data.debts.forEach(d=>s.add(mk(d.vencimiento)));return [...s].sort().reverse()},[data,selectedMonth]);
const fm=useMemo(()=>({invoices:data.invoices.filter(i=>mk(i.vencimiento)===selectedMonth),incomes:data.incomes.filter(i=>mk(i.fecha)===selectedMonth),expenses:data.expenses.filter(e=>mk(e.fecha)===selectedMonth),debts:data.debts.filter(d=>mk(d.vencimiento)===selectedMonth)}),[data,selectedMonth]);
const stats=useMemo(()=>{let ingresos=fm.incomes.reduce((s,x)=>s+ +x.monto,0),egresos=fm.expenses.reduce((s,x)=>s+ +x.monto,0),deudas=fm.debts.filter(d=>d.estado!=="Pagada").reduce((s,x)=>s+ +x.monto,0),pend=fm.invoices.filter(i=>ist(i).l!=="Pagada").reduce((s,x)=>s+ +x.monto,0),venc=fm.invoices.filter(i=>ist(i).l==="Vencida");return{ingresos,egresos,deudas,pend,venc,saldo:ingresos-egresos}},[fm]);
const filteredClients=data.clients.filter(c=>`${c.nombre} ${c.rut} ${c.email} ${c.giro}`.toLowerCase().includes(search.toLowerCase()));
const filteredInvoices=data.invoices.filter(i=>`${i.factura} ${client(i.clienteId)?.nombre||""}`.toLowerCase().includes(search.toLowerCase()));
const alertInvoices=data.invoices.filter(i=>["Vencida","Por vencer"].includes(ist(i).l)||days(i.vencimiento)===3).filter(i=>`${i.factura} ${client(i.clienteId)?.nombre||""}`.toLowerCase().includes(alertSearch.toLowerCase()));
const selectedInvoice=selectedInvoiceId?data.invoices.find(i=>i.id===selectedInvoiceId):alertInvoices[0], selectedClient=selectedInvoice?client(selectedInvoice.clienteId):null;
const monthly=months.slice().reverse().map(m=>({mes:ml(m),ingresos:data.incomes.filter(i=>mk(i.fecha)===m).reduce((s,x)=>s+ +x.monto,0),egresos:data.expenses.filter(e=>mk(e.fecha)===m).reduce((s,x)=>s+ +x.monto,0),deudas:data.debts.filter(d=>mk(d.vencimiento)===m&&d.estado!=="Pagada").reduce((s,x)=>s+ +x.monto,0),vencidas:data.invoices.filter(i=>mk(i.vencimiento)===m&&ist(i).l==="Vencida").reduce((s,x)=>s+ +x.monto,0)}));
const pie=[{name:"Pagadas",value:fm.invoices.filter(i=>ist(i).l==="Pagada").length},{name:"Pendientes",value:fm.invoices.filter(i=>ist(i).l==="Pendiente").length},{name:"Por vencer",value:fm.invoices.filter(i=>ist(i).l==="Por vencer").length},{name:"Vencidas",value:fm.invoices.filter(i=>ist(i).l==="Vencida").length}];
function saveClient(){if(!clientForm.nombre||!clientForm.rut)return;if(editingClient){setData({...data,clients:data.clients.map(c=>c.id===editingClient?{...clientForm,id:editingClient}:c)});setEditingClient(null)}else setData({...data,clients:[{...clientForm,id:Date.now()},...data.clients]});setClientForm({nombre:"",rut:"",giro:"",telefono:"569",email:"",direccion:"",contacto:""})}
function editClient(c){setEditingClient(c.id);setClientForm(c);setTab("clientes")}
function saveInvoice(){if(!invoiceForm.clienteId||!invoiceForm.factura||!invoiceForm.monto)return;let p={...invoiceForm,id:editingInvoice||Date.now(),clienteId:+invoiceForm.clienteId,monto:+invoiceForm.monto};setData({...data,invoices:editingInvoice?data.invoices.map(i=>i.id===editingInvoice?p:i):[p,...data.invoices]});setEditingInvoice(null);setInvoiceForm({clienteId:"",factura:"",emision:today(),vencimiento:today(),monto:"",estado:"Pendiente",detalle:""})}
function editInvoice(i){setEditingInvoice(i.id);setInvoiceForm({...i,clienteId:String(i.clienteId)});setTab("facturas")}
function saveIncome(){if(!incomeForm.categoria||!incomeForm.monto)return;let inv=data.invoices,desc=incomeForm.descripcion;if(incomeForm.facturaId){let f=data.invoices.find(i=>+i.id===+incomeForm.facturaId);inv=data.invoices.map(i=>+i.id===+incomeForm.facturaId?{...i,estado:"Pagada"}:i);if(!desc&&f)desc=`Pago ${f.factura}`}setData({...data,invoices:inv,incomes:[{...incomeForm,id:Date.now(),monto:+incomeForm.monto,descripcion:desc},...data.incomes]});setIncomeForm({fecha:today(),categoria:"Pago de factura",descripcion:"",monto:"",facturaId:""})}
function saveExpense(){
if(!expenseForm.categoria||!expenseForm.monto)return;
let debts=data.debts;
let descripcion=expenseForm.descripcion;
if(expenseForm.debtId){
 const deuda=data.debts.find(d=>+d.id===+expenseForm.debtId);
 debts=data.debts.map(d=>+d.id===+expenseForm.debtId?{...d,estado:"Pagada"}:d);
 if(!descripcion&&deuda)descripcion=`Pago factura/deuda ${deuda.proveedor} - ${deuda.descripcion}`;
}
setData({...data,debts,expenses:[{...expenseForm,id:Date.now(),monto:+expenseForm.monto,descripcion},...data.expenses]});
setExpenseForm({fecha:today(),categoria:"Pago instalador",descripcion:"",monto:"",debtId:""});
}
function saveDebt(){if(!debtForm.proveedor||!debtForm.monto)return;setData({...data,debts:[{...debtForm,id:Date.now(),monto:+debtForm.monto},...data.debts]});setDebtForm({fecha:today(),proveedor:"",categoria:"Compra de equipos",descripcion:"",monto:"",vencimiento:today(),estado:"Pendiente"})}
function attachFile(id,file){if(!file)return;let r=new FileReader();r.onload=()=>setData({...data,attachments:{...(data.attachments||{}),[id]:{name:file.name,size:file.size,type:file.type,dataUrl:r.result,attachedAt:new Date().toLocaleString("es-CL")}}});r.readAsDataURL(file)}
function canSendReminder(id){
  if(!data.attachments?.[id]){
    alert("Debe adjuntar la factura antes de enviar el recordatorio.");
    return false;
  }
  return true;
}


function manualSave(){
  localStorage.setItem(KEY, JSON.stringify(data));
  setSaved(new Date().toLocaleTimeString("es-CL"));
  alert("Información guardada correctamente.");
}
function exportBackup(){
  const blob = new Blob([JSON.stringify(data,null,2)], {type:"application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "respaldo_gpsruta_"+new Date().toISOString().slice(0,10)+".json";
  a.click();
  URL.revokeObjectURL(url);
}
function importBackup(file){
  if(!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try{
      const imported = JSON.parse(reader.result);
      setData({...seed, ...imported});
      alert("Respaldo importado correctamente.");
    }catch(e){
      alert("El archivo no es válido.");
    }
  };
  reader.readAsText(file);
}
function deleteMonthHistory(kind){
  if(!confirm("¿Eliminar todos los registros de "+ml(selectedMonth)+"? Esta acción no se puede deshacer.")) return;
  if(kind==="debts") setData({...data, debts:data.debts.filter(d=>mk(d.vencimiento)!==selectedMonth)});
  if(kind==="incomes") setData({...data, incomes:data.incomes.filter(i=>mk(i.fecha)!==selectedMonth)});
  if(kind==="expenses") setData({...data, expenses:data.expenses.filter(e=>mk(e.fecha)!==selectedMonth)});
  if(kind==="invoices") setData({...data, invoices:data.invoices.filter(i=>mk(i.vencimiento)!==selectedMonth)});
}
function deleteAllHistory(kind){
  if(!confirm("¿Eliminar TODO el historial de esta sección? Esta acción no se puede deshacer.")) return;
  if(kind==="debts") setData({...data, debts:[]});
  if(kind==="incomes") setData({...data, incomes:[]});
  if(kind==="expenses") setData({...data, expenses:[]});
  if(kind==="invoices") setData({...data, invoices:[]});
}

function importClientsExcel(file){
  if(!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try{
      const workbook = XLSX.read(new Uint8Array(e.target.result), {type:"array"});
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {defval:""});
      if(!rows.length){
        alert("El Excel no tiene datos.");
        return;
      }
      const normalize = (row, keys) => {
        for(const k of keys){
          const found = Object.keys(row).find(x => x.toLowerCase().trim() === k.toLowerCase());
          if(found) return String(row[found] || "").trim();
        }
        return "";
      };
      const nuevos = rows.map((r,idx)=>({
        id: Date.now()+idx,
        nombre: normalize(r,["nombre","cliente","razon social","razón social","RAZON SOCIAL","RAZÓN SOCIAL","empresa"]),
        rut: normalize(r,["rut","r.u.t","rut cliente"]),
        giro: normalize(r,["giro","actividad","actividad economica","actividad económica"]),
        telefono: normalize(r,["telefono","teléfono","celular","whatsapp"]),
        email: normalize(r,["email","correo","correo electronico","correo electrónico"]),
        direccion: normalize(r,["direccion","dirección","domicilio"]),
        contacto: normalize(r,["contacto","persona contacto","encargado"])
      })).filter(c=>c.nombre || c.rut);
      if(!nuevos.length){
        alert("No se encontraron clientes válidos. Revisa que el Excel tenga columnas: nombre, rut, giro, telefono, email, direccion, contacto.");
        return;
      }
      setData({...data, clients:[...nuevos, ...data.clients]});
      alert("Clientes cargados correctamente: "+nuevos.length);
    }catch(err){
      alert("No se pudo leer el Excel. Revisa el formato del archivo.");
    }
  };
  reader.readAsArrayBuffer(file);
}

function exportFinancialReport(){
  const clienteNombre = (id) => client(id)?.nombre || "";
  const clienteRut = (id) => client(id)?.rut || "";
  const workbook = XLSX.utils.book_new();

  const resumen = [{
    "Mes/Año": ml(selectedMonth),
    "Ingresos": stats.ingresos,
    "Egresos": stats.egresos,
    "Saldo Neto": stats.saldo,
    "Deudas por pagar": stats.deudas,
    "Facturas pendientes por cobrar": stats.pend,
    "Facturas vencidas": stats.venc.length,
    "Fecha descarga": new Date().toLocaleString("es-CL")
  }];

  const pagos = data.incomes
    .filter(i => i.categoria === "Pago de factura" || i.facturaId)
    .map(i => {
      const inv = data.invoices.find(f => Number(f.id) === Number(i.facturaId));
      return {
        "Fecha Pago": i.fecha,
        "Mes/Año": ml(mk(i.fecha)),
        "Factura": inv?.factura || i.descripcion || "",
        "Cliente": inv ? clienteNombre(inv.clienteId) : "",
        "RUT": inv ? clienteRut(inv.clienteId) : "",
        "Categoría": i.categoria,
        "Descripción": i.descripcion,
        "Monto": Number(i.monto || 0)
      };
    });

  const ingresos = data.incomes.map(i => ({
    "Fecha": i.fecha,
    "Mes/Año": ml(mk(i.fecha)),
    "Categoría": i.categoria,
    "Descripción": i.descripcion,
    "Factura Asociada": data.invoices.find(f => Number(f.id) === Number(i.facturaId))?.factura || "",
    "Monto": Number(i.monto || 0)
  }));

  const egresos = data.expenses.map(e => ({
    "Fecha": e.fecha,
    "Mes/Año": ml(mk(e.fecha)),
    "Categoría": e.categoria,
    "Descripción": e.descripcion,
    "Monto": Number(e.monto || 0)
  }));

  const deudas = data.debts.map(d => ({
    "Fecha Registro": d.fecha,
    "Proveedor": d.proveedor,
    "Categoría": d.categoria,
    "Descripción": d.descripcion,
    "Vencimiento": d.vencimiento,
    "Mes/Año": ml(mk(d.vencimiento)),
    "Estado": d.estado,
    "Monto": Number(d.monto || 0)
  }));

  const facturasPorPagar = data.debts.filter(d => d.estado !== "Pagada").map(d => ({
    "Proveedor": d.proveedor,
    "Categoría": d.categoria,
    "Descripción": d.descripcion,
    "Vencimiento": d.vencimiento,
    "Mes/Año": ml(mk(d.vencimiento)),
    "Estado": d.estado,
    "Monto": Number(d.monto || 0)
  }));

  const facturasPorCobrar = data.invoices.map(f => ({
    "Factura": f.factura,
    "Cliente": clienteNombre(f.clienteId),
    "RUT Cliente": clienteRut(f.clienteId),
    "Emisión": f.emision,
    "Vencimiento": f.vencimiento,
    "Mes/Año": ml(mk(f.vencimiento)),
    "Estado": ist(f).l,
    "Detalle": f.detalle,
    "Monto": Number(f.monto || 0)
  }));

  const clientes = data.clients.map(c => ({
    "Nombre / Razón Social": c.nombre,
    "RUT": c.rut,
    "Giro": c.giro,
    "Teléfono": c.telefono,
    "Email": c.email,
    "Dirección": c.direccion,
    "Contacto": c.contacto
  }));

  const sheets = [
    ["Resumen", resumen],
    ["Pagos", pagos],
    ["Ingresos", ingresos],
    ["Egresos", egresos],
    ["Deudas", deudas],
    ["Facturas por pagar", facturasPorPagar],
    ["Facturas por cobrar", facturasPorCobrar],
    ["Clientes", clientes]
  ];

  sheets.forEach(([name, rows]) => {
    const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{}]);
    XLSX.utils.book_append_sheet(workbook, ws, name);
  });

  XLSX.writeFile(workbook, "informe_gpsruta_"+selectedMonth+".xlsx");
}

function importInvoicesExcel(file){
 if(!file)return;
 const reader=new FileReader();
 reader.onload=(e)=>{
  try{
   const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
   const sh=wb.Sheets[wb.SheetNames[0]];
   const rows=XLSX.utils.sheet_to_json(sh,{defval:""});
   const norm=(row,keys)=>{for(const k of keys){const f=Object.keys(row).find(x=>x.toLowerCase().trim()===k.toLowerCase());if(f)return String(row[f]||"").trim()}return ""};
   const findClientId=(rut,nombre)=>{
    const rc=String(rut||"").replace(/\./g,"").replace(/-/g,"").toLowerCase();
    let c=data.clients.find(cli=>String(cli.rut||"").replace(/\./g,"").replace(/-/g,"").toLowerCase()===rc);
    if(c)return c.id;
    c=data.clients.find(cli=>String(cli.nombre||"").toLowerCase().trim()===String(nombre||"").toLowerCase().trim());
    return c?.id||"";
   };
   const nuevas=rows.map((r,idx)=>{
    const rut=norm(r,["rut cliente","rut","RUT","RUT CLIENTE"]);
    const nombre=norm(r,["cliente","nombre cliente","razon social","razón social","RAZON SOCIAL"]);
    const monto=Number(String(norm(r,["monto","valor","total"])).replace(/\$/g,"").replace(/\./g,"").replace(/,/g,"."))||0;
    return {id:Date.now()+idx,clienteId:findClientId(rut,nombre),factura:norm(r,["factura","n factura","n° factura","numero factura","número factura","folio"]),emision:norm(r,["emision","emisión","fecha emision","fecha emisión"])||today(),vencimiento:norm(r,["vencimiento","fecha vencimiento","vence"])||today(),monto,estado:norm(r,["estado"])||"Pendiente",detalle:norm(r,["detalle","descripcion","descripción","glosa"])||"Factura cargada masivamente"};
   }).filter(f=>f.factura&&f.clienteId&&f.monto);
   if(!nuevas.length){alert("No se encontraron facturas válidas. Revisa columnas: factura, rut cliente o cliente, vencimiento, monto.");return}
   setData({...data,invoices:[...nuevas,...data.invoices]});
   alert("Facturas cargadas correctamente: "+nuevas.length);
  }catch(err){alert("No se pudo leer el Excel. Revisa el formato del archivo.")}
 };
 reader.readAsArrayBuffer(file);
}
if(!logged)return <Login onLogin={()=>setLogged(true)}/>;
return <div className="app"><aside><Logo/><div className="admin"><User size={24}/><div><b>Administrador</b><p>admin@gpsruta.cl</p></div></div><nav>{[["dashboard","Dashboard",Eye],["clientes","Clientes",Users],["facturas","Facturas por cobrar",FileText],["deudas","Deudas / Facturas por pagar",CreditCard],["ingresos","Ingresos",TrendingUp],["egresos","Egresos",TrendingDown],["alertas","Cobros / Recordatorios",Bell]].map(([v,l,I])=><button key={v} onClick={()=>setTab(v)} className={tab===v?"active":""}><I size={20}/>{l}</button>)}</nav><div className="autosave"><CheckCircle size={20}/><div><b>Guardado automático activo</b><p>Último guardado: {saved}</p></div></div><button className="logout" onClick={()=>{sessionStorage.removeItem(SESSION);setLogged(false)}}><LogOut size={19}/>Cerrar sesión</button></aside><main><header><div className="search"><Search size={17}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar cliente, factura o giro..."/></div><div className="chips"><select value={selectedMonth} onChange={e=>setSelectedMonth(e.target.value)} className="monthSelect">{months.map(m=><option key={m} value={m}>{ml(m)}</option>)}</select><span><CalendarDays size={17}/>{clock.toLocaleDateString("es-CL")}</span><span><Clock size={17}/>{clock.toLocaleTimeString("es-CL")}</span><span className="green"><Save size={17}/>Guardado automático</span></div></header>
<section className="backupToolbar">
  <button className="backupBtn saveManual" onClick={manualSave}><HardDrive size={17}/>Guardar ahora</button>
  <button className="backupBtn" onClick={exportBackup}><Download size={17}/>Respaldar</button>
  <label className="backupBtn importBtn"><Upload size={17}/>Importar respaldo<input type="file" accept=".json" onChange={e=>importBackup(e.target.files?.[0])}/></label><button className="backupBtn reportBtn" onClick={exportFinancialReport}>📊 Descargar informe Excel</button>
</section>
<section className="kpis"><K t="Ingresos del mes" v={money(stats.ingresos)} s={ml(selectedMonth)} icon={TrendingUp}/><K t="Egresos del mes" v={money(stats.egresos)} s={ml(selectedMonth)} icon={TrendingDown} tone="red"/><K t="Deudas por pagar" v={money(stats.deudas)} s="Según vencimiento" icon={CreditCard} tone="gold"/><K t="Facturas pendientes" v={money(stats.pend)} s="Por cobrar" icon={FileText} tone="blue"/><K t="Facturas vencidas" v={stats.venc.length} s="Cantidad mensual" icon={AlertTriangle} tone="red"/></section>
{tab==="dashboard"&&<section className="gridDash"><div className="card chartPro wide"><h2>Resumen mensual financiero</h2><div className="chart compact"><ResponsiveContainer><BarChart data={monthly}><CartesianGrid stroke="rgba(255,255,255,.08)"/><XAxis dataKey="mes" stroke="#ccc"/><YAxis stroke="#ccc" tickFormatter={v=>`${Math.round(v/1000000)}M`}/><Tooltip formatter={v=>money(v)} contentStyle={{background:"#050505",border:"1px solid #00a3ff"}}/><Bar dataKey="ingresos" fill="#7CFC00" radius={[6,6,0,0]} barSize={18}/>
<Bar dataKey="egresos" fill="#ff3131" radius={[6,6,0,0]} barSize={18}/>
<Bar dataKey="deudas" fill="#FFD43B" radius={[6,6,0,0]} barSize={18}/>
<Bar dataKey="vencidas" fill="#ff9500" radius={[6,6,0,0]} barSize={18}/></BarChart></ResponsiveContainer></div></div><div className="card chartPro"><h2>Estado facturas del mes</h2><div className="chart donut"><ResponsiveContainer><PieChart><Pie data={pie} dataKey="value" nameKey="name" innerRadius={42} outerRadius={72}>{pie.map((_,i)=><Cell key={i} fill={["#7CFC00","#FFD43B","#ff9500","#ff3131"][i]}/>)}</Pie><Tooltip contentStyle={{background:"#050505",border:"1px solid #00a3ff"}}/></PieChart></ResponsiveContainer></div></div><div className="card wide"><h2>Resumen del mes seleccionado</h2><div className="summaryGrid">{[[money(stats.ingresos),"Ingresos"],[money(stats.egresos),"Egresos"],[money(stats.saldo),"Saldo neto"],[money(stats.deudas),"Deudas por pagar"],[money(stats.pend),"Facturas por cobrar"],[stats.venc.length,"Facturas vencidas"]].map(([a,b])=><div key={b}><b>{a}</b><span>{b}</span></div>)}</div></div></section>}
{tab==="clientes"&&<section className="two"><div className="card"><h2>{editingClient?"Editar cliente":"Nuevo cliente"}</h2>
<div className="excelImportBox">
  <label className="excelBtn">📥 Cargar clientes desde Excel
    <input type="file" accept=".xlsx,.xls,.csv" onChange={e=>importClientsExcel(e.target.files?.[0])}/>
  </label>
  <small>Columnas aceptadas: nombre, RAZON SOCIAL, rut, giro, telefono, email, direccion, contacto.</small>
</div>
<Fields obj={clientForm} set={setClientForm} fields={["nombre","rut","giro","telefono","email","direccion","contacto"]}/><button className="primary" onClick={saveClient}><Plus size={17}/>Guardar cliente</button></div><div className="cards">{filteredClients.map(c=><div className="card client" key={c.id}><div className="clientIcon"><Building2 size={34}/></div><div><h3>{c.nombre}</h3><p>RUT: {c.rut}</p><p>Giro: {c.giro||"Sin giro registrado"}</p><p>{c.email}</p><div className="actions"><button className="icon edit" onClick={()=>editClient(c)}><Edit size={17}/></button><button className="icon trash" onClick={()=>setData({...data,clients:data.clients.filter(x=>x.id!==c.id),invoices:data.invoices.filter(i=>+i.clienteId!==+c.id)})}><Trash2 size={17}/></button></div></div></div>)}</div></section>}
{tab==="facturas"&&<section className="two"><div className="card"><h2>{editingInvoice?"Editar factura":"Nueva factura por cobrar"}</h2>
<div className="excelImportBox"><label className="excelBtn">📥 Cargar facturas desde Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={e=>importInvoicesExcel(e.target.files?.[0])}/></label><small>Columnas: factura, rut cliente o cliente, emision, vencimiento, monto, estado, detalle.</small></div><select value={invoiceForm.clienteId} onChange={e=>setInvoiceForm({...invoiceForm,clienteId:e.target.value})}><option value="">Seleccionar cliente</option>{data.clients.map(c=><option key={c.id} value={c.id}>{c.nombre}</option>)}</select><Fields obj={invoiceForm} set={setInvoiceForm} fields={["factura","emision","vencimiento","monto","detalle"]}/><select value={invoiceForm.estado} onChange={e=>setInvoiceForm({...invoiceForm,estado:e.target.value})}><option>Pendiente</option><option>Vencida</option><option>Pagada</option></select><button className="primary" onClick={saveInvoice}><Plus size={17}/>Guardar factura</button></div><div className="card">
<div className="historyHead"><h2>Listado de facturas por cobrar</h2><small className="requiredAttach">Adjunte factura antes de enviar recordatorio.</small></div>
<InvTable items={filteredInvoices} client={client} edit={editInvoice} del={id=>setData({...data,invoices:data.invoices.filter(i=>i.id!==id)})} data={data} attachFile={attachFile} canSendReminder={canSendReminder}/>
</div></section>}
{tab==="deudas"&&<section className="two"><div className="card"><h2>Deudas / Facturas por pagar</h2><Fields obj={debtForm} set={setDebtForm} fields={["fecha","proveedor","descripcion","monto","vencimiento"]}/><select value={debtForm.categoria} onChange={e=>setDebtForm({...debtForm,categoria:e.target.value})}>{expenseCats.map(c=><option key={c}>{c}</option>)}</select><select value={debtForm.estado} onChange={e=>setDebtForm({...debtForm,estado:e.target.value})}><option>Pendiente</option><option>Pagada</option></select><button className="primary gold" onClick={saveDebt}><Plus size={17}/>Guardar deuda</button></div><div className="card"><div className="historyHead"><h2>Historial de deudas por mes/año</h2><div><button className="smallDanger" onClick={()=>deleteMonthHistory("debts")}>Eliminar mes/año</button><button className="smallDanger ghost" onClick={()=>deleteAllHistory("debts")}>Eliminar todo</button></div></div><Table rows={data.debts} type="debts" setData={setData} data={data}/></div></section>}
{tab==="ingresos"&&<section className="two"><div className="card"><h2>Ingresar ingreso</h2><Fields obj={incomeForm} set={setIncomeForm} fields={["fecha","descripcion","monto"]}/><select value={incomeForm.categoria} onChange={e=>setIncomeForm({...incomeForm,categoria:e.target.value})}>{incomeCats.map(c=><option key={c}>{c}</option>)}</select><select value={incomeForm.facturaId} onChange={e=>{let inv=data.invoices.find(i=>+i.id===+e.target.value);setIncomeForm({...incomeForm,facturaId:e.target.value,monto:inv?inv.monto:incomeForm.monto,descripcion:inv?`Pago ${inv.factura}`:incomeForm.descripcion})}}><option value="">Sin asociar factura</option>{data.invoices.filter(i=>ist(i).l!=="Pagada").map(i=><option key={i.id} value={i.id}>{i.factura} · {client(i.clienteId)?.nombre} · {money(i.monto)}</option>)}</select><button className="primary" onClick={saveIncome}><Plus size={17}/>Guardar ingreso</button></div><div className="card"><div className="historyHead"><h2>Historial de ingresos por fecha</h2><div><button className="smallDanger" onClick={()=>deleteMonthHistory("incomes")}>Eliminar mes/año</button><button className="smallDanger ghost" onClick={()=>deleteAllHistory("incomes")}>Eliminar todo</button></div></div><Table rows={data.incomes} type="incomes" setData={setData} data={data}/></div></section>}
{tab==="egresos"&&<section className="two"><div className="card"><h2>Ingresar egreso</h2><Fields obj={expenseForm} set={setExpenseForm} fields={["fecha","descripcion","monto"]}/><select value={expenseForm.categoria} onChange={e=>setExpenseForm({...expenseForm,categoria:e.target.value})}>{expenseCats.map(c=><option key={c}>{c}</option>)}</select>
<select value={expenseForm.debtId} onChange={e=>{let d=data.debts.find(x=>+x.id===+e.target.value);setExpenseForm({...expenseForm,debtId:e.target.value,monto:d?d.monto:expenseForm.monto,descripcion:d?`Pago factura/deuda ${d.proveedor} - ${d.descripcion}`:expenseForm.descripcion,categoria:d?d.categoria:expenseForm.categoria})}}>
<option value="">Sin vincular factura por pagar</option>
{data.debts.filter(d=>d.estado!=="Pagada").map(d=><option key={d.id} value={d.id}>{d.proveedor} · {d.descripcion} · {money(d.monto)} · vence {d.vencimiento}</option>)}
</select>
<small className="hint">Al vincular una factura/deuda por pagar y guardar el egreso, quedará marcada como Pagada.</small>
<button className="primary gold" onClick={saveExpense}><Plus size={17}/>Guardar egreso</button></div><div className="card"><div className="historyHead"><h2>Historial de egresos por fecha</h2><div><button className="smallDanger" onClick={()=>deleteMonthHistory("expenses")}>Eliminar mes/año</button><button className="smallDanger ghost" onClick={()=>deleteAllHistory("expenses")}>Eliminar todo</button></div></div><Table rows={data.expenses} type="expenses" setData={setData} data={data}/></div></section>}
{tab==="alertas"&&<section className="alerts"><div className="card reminders"><h2><Bell size={20}/>Cobros / Recordatorios</h2><div className="search inner"><Search size={17}/><input value={alertSearch} onChange={e=>setAlertSearch(e.target.value)} placeholder="Buscar factura..."/></div><div className="reminderList">{alertInvoices.map(inv=>{let c=client(inv.clienteId),s=ist(inv),Icon=s.I;return <button key={inv.id} className={`reminder ${selectedInvoice?.id===inv.id?"selected":""}`} onClick={()=>setSelectedInvoiceId(inv.id)}><Icon className={s.c} size={24}/><div><b>{inv.factura}</b><p>{c?.nombre}</p><small>{inv.vencimiento} · {money(inv.monto)}</small></div></button>})}</div></div><div className="card attach"><h2><Paperclip size={20}/>Adjuntar factura</h2>{selectedInvoice?<><div className="selected"><b>{selectedInvoice.factura}</b><p>{selectedClient?.nombre} · {money(selectedInvoice.monto)}</p><small className="requiredAttach">Para enviar recordatorio, primero debe adjuntar la factura.</small></div><label className="drop"><UploadCloud size={32}/><b>Buscar factura en mi PC</b><small>PDF, JPG, PNG, DOCX, XLSX</small><input type="file" onChange={e=>attachFile(selectedInvoice.id,e.target.files?.[0])}/></label>{data.attachments?.[selectedInvoice.id]&&<div className="fileBox"><FileText size={22}/><div><b>{data.attachments[selectedInvoice.id].name}</b><p>{(data.attachments[selectedInvoice.id].size/1024/1024).toFixed(2)} MB</p></div></div>}<div className="actions big"><a className={`send whatsapp ${!data.attachments?.[selectedInvoice.id]?"disabled":""}`} href={data.attachments?.[selectedInvoice.id]?wa(selectedInvoice,selectedClient):"#"} onClick={(e)=>{if(!canSendReminder(selectedInvoice.id))e.preventDefault()}} target="_blank"><W/>WhatsApp</a><a className={`send mail ${!data.attachments?.[selectedInvoice.id]?"disabled":""}`} href={data.attachments?.[selectedInvoice.id]?em(selectedInvoice,selectedClient):"#"} onClick={(e)=>{if(!canSendReminder(selectedInvoice.id))e.preventDefault()}}><Mail size={18}/>Correo</a></div></>:<p>No hay facturas por cobrar.</p>}</div></section>}
</main></div>}
function Table({rows,type,setData,data}){let sorted=[...rows].sort((a,b)=>(b.fecha||b.vencimiento).localeCompare(a.fecha||a.vencimiento));return <div className="tableWrap"><table><thead><tr><th>Fecha</th><th>Mes/Año</th><th>Categoría</th><th>Descripción</th><th>Monto</th><th></th></tr></thead><tbody>{sorted.map(r=><tr key={r.id}><td>{r.fecha||r.vencimiento}</td><td>{ml(mk(r.fecha||r.vencimiento))}</td><td>{r.categoria}{r.proveedor&&<small>{r.proveedor}</small>}</td><td>{r.descripcion}{r.estado&&<small>Estado: {r.estado}</small>}</td><td>{money(r.monto)}</td><td><button className="icon trash" onClick={()=>setData({...data,[type]:(data[type]||[]).filter(x=>x.id!==r.id)})}><Trash2 size={17}/></button></td></tr>)}</tbody></table></div>}
createRoot(document.getElementById("root")).render(<App/>);
