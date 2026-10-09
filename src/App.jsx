import React, { useEffect, useState } from 'react'
import { supabase } from './supabaseClient'
import * as XLSX from 'xlsx'
import { 
  Package, 
  CheckCircle2, 
  Plus, 
  Loader2, 
  Edit2, 
  Trash2, 
  Check, 
  X, 
  Layers, 
  ShoppingBag,
  Barcode,
  Search,
  AlertTriangle,
  XCircle,
  PlusCircle,
  MinusCircle,
  Calendar,
  ChevronDown,
  ChevronUp,
  Clock,
  ShoppingCart,
  TrendingUp,
  DollarSign,
  LayoutDashboard,
  RotateCcw,
  Receipt,
  Scale,
  Users,
  Printer,
  UserPlus,
  Vault,
  Lock,
  Unlock,
  Coins,
  History,
  FileSpreadsheet,
  Download
} from 'lucide-react'
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid
} from 'recharts'

// Mock de base de datos DGI (Uruguay)
const MOCK_EMPRESAS_DGI = {
  "211234560018": { nombre: "DISTRIBUIDORA URUGUAY S.A.", direccion: "Av. Italia 2450", telefono: "24801234" },
  "210000000019": { nombre: "CONAPROLE - COOP. NACIONAL DE LECHEROS", direccion: "Magallanes 1871", telefono: "29240000" },
  "210345670012": { nombre: "SUPERMERCADOS TA-TA S.A.", direccion: "Av. 18 de Julio 1234", telefono: "29001122" },
  "219998880015": { nombre: "PANADERÍA Y FIAMBRERÍA DEL ESTE SRL", direccion: "Giannattasio Km 21", telefono: "26829999" }
}

function App() {
  const [pestana, setPestana] = useState('pos')
  const [sidebarColapsado, setSidebarColapsado] = useState(false)
  
  // Control de Caja / Sesiones
  const [cajaActual, setCajaActual] = useState(null)
  const [cargandoCaja, setCargandoCaja] = useState(true)
  const [inputMontoInicial, setInputMontoInicial] = useState('')
  const [procesandoAperturaCaja, setProcesandoAperturaCaja] = useState(false)
  const [inputMontoArqueo, setInputMontoArqueo] = useState('')
  const [procesandoCierreCaja, setProcesandoCierreCaja] = useState(false)
  const [historialCajas, setHistorialCajas] = useState([])

  // Modales
  const [modalProductoAbierto, setModalProductoAbierto] = useState(false)
  const [modalCategoriaAbierto, setModalCategoriaAbierto] = useState(false)
  const [modalClienteAbierto, setModalClienteAbierto] = useState(false)
  const [modalTicketAbierto, setModalTicketAbierto] = useState(false)

  // Modal de Venta por Peso
  const [modalPesoAbierto, setModalPesoAbierto] = useState(false)
  const [productoParaPeso, setProductoParaPeso] = useState(null)
  const [inputGramos, setInputGramos] = useState('')
  const [inputMontoDinero, setInputMontoDinero] = useState('')

  // Categorías
  const [categorias, setCategorias] = useState([])
  const [cargandoCategorias, setCargandoCategorias] = useState(true)
  const [nuevoNombreCat, setNuevoNombreCat] = useState('')
  const [guardandoCat, setGuardandoCat] = useState(false)
  const [editandoCatId, setEditandoCatId] = useState(null)
  const [nombreCatEditado, setNombreCatEditado] = useState('')

  // Clientes
  const [clientes, setClientes] = useState([])
  const [cargandoClientes, setCargandoClientes] = useState(true)
  const [guardandoCliente, setGuardandoCliente] = useState(false)
  const [formCliente, setFormCliente] = useState({ nombre: '', rut_ci: '', telefono: '', direccion: '' })
  const [clienteSeleccionadoPOS, setClienteSeleccionadoPOS] = useState(null)
  const [inputRUTPOS, setInputRUTPOS] = useState('')

  // Productos
  const [productos, setProductos] = useState([])
  const [lotes, setLotes] = useState([])
  const [busqueda, setBusqueda] = useState('')
  const [filtroStock, setFiltroStock] = useState('todos')
  const [cargandoProductos, setCargandoProductos] = useState(true)
  const [guardandoProd, setGuardandoProd] = useState(false)
  const [formProd, setFormProd] = useState({
    nombre: '',
    codigo_barra: '',
    precio: '',
    stock: '',
    categoria_id: '',
    tipo_venta: 'unidad'
  })
  const [editandoProdId, setEditandoProdId] = useState(null)
  const [formProdEdit, setFormProdEdit] = useState({
    nombre: '',
    codigo_barra: '',
    precio: '',
    stock: '',
    categoria_id: '',
    tipo_venta: 'unidad'
  })

  // Lotes
  const [prodSeleccionadoLote, setProdSeleccionadoLote] = useState(null)
  const [nuevoLote, setNuevoLote] = useState({ cantidad: '', fecha_vencimiento: '' })
  const [guardandoLote, setGuardandoLote] = useState(false)

  // POS
  const [carrito, setCarrito] = useState([])
  const [busquedaPos, setBusquedaPos] = useState('')
  const [metodoPago, setMetodoPago] = useState('Efectivo')
  const [procesandoVenta, setProcesandoVenta] = useState(false)
  const [ultimaVentaRealizada, setUltimaVentaRealizada] = useState(null)

  // Ventas
  const [ventas, setVentas] = useState([])
  const [procesandoAnulacion, setProcesandoAnulacion] = useState(null)

  useEffect(() => {
    obtenerEstadoCaja()
    obtenerCategorias()
    obtenerProductos()
    obtenerLotes()
    obtenerVentas()
    obtenerClientes()
    obtenerHistorialCajas()
  }, [])

  // --- REIMPRESIÓN DESDE HISTORIAL ---
  function reimprimirTicketVenta(venta) {
    const itemsFormateados = (venta.detalle_ventas || []).map((d) => ({
      nombre: d.productos?.nombre || 'Producto',
      cantidadEnCarrito: d.cantidad,
      precio: d.precio_unitario,
      subtotal: d.subtotal,
      tipo_venta: d.productos?.tipo_venta || 'unidad',
      gramosOriginales: d.productos?.tipo_venta === 'peso' ? Math.round(d.cantidad * 1000) : null
    }))

    setUltimaVentaRealizada({
      id: venta.id,
      created_at: venta.created_at,
      total: venta.total,
      metodo_pago: venta.metodo_pago,
      clientes: venta.clientes,
      items: itemsFormateados
    })

    setModalTicketAbierto(true)
  }

  // --- EXPORTACIÓN A EXCEL ---
  function exportarVentasExcel() {
    if (ventas.length === 0) {
      alert('No hay ventas registradas para exportar.')
      return
    }

    const datosFormateados = ventas.map((v) => ({
      'ID Venta': v.id,
      'Fecha y Hora': new Date(v.created_at).toLocaleString('es-UY'),
      'Cliente / Razón Social': v.clientes?.nombre || 'Consumidor Final',
      'RUT / C.I.': v.clientes?.rut_ci || 'N/A',
      'Método de Pago': v.metodo_pago,
      'Estado': v.estado.toUpperCase(),
      'Total ($)': Number(v.total)
    }))

    const hoja = XLSX.utils.json_to_sheet(datosFormateados)
    const libro = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(libro, hoja, 'Reporte de Ventas')

    hoja['!cols'] = [
      { wch: 10 },
      { wch: 20 },
      { wch: 30 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 12 }
    ]

    const fechaHoy = new Date().toISOString().split('T')[0]
    XLSX.writeFile(libro, `Reporte_Ventas_StockPulse_${fechaHoy}.xlsx`)
  }

  function exportarInventarioExcel() {
    if (productos.length === 0) {
      alert('No hay productos en el inventario para exportar.')
      return
    }

    const datosFormateados = productos.map((p) => ({
      'ID': p.id,
      'Nombre': p.nombre,
      'Código de Barras': p.codigo_barra || 'Sin código',
      'Categoría': p.categorias?.nombre || 'Sin Categoría',
      'Tipo Venta': p.tipo_venta === 'peso' ? 'Por Peso (Kg)' : 'Por Unidad',
      'Precio ($)': Number(p.precio),
      'Stock Actual': p.stock,
      'Valor Invertido ($)': p.precio * p.stock
    }))

    const hoja = XLSX.utils.json_to_sheet(datosFormateados)
    const libro = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(libro, hoja, 'Inventario')

    hoja['!cols'] = [
      { wch: 8 },
      { wch: 25 },
      { wch: 18 },
      { wch: 20 },
      { wch: 18 },
      { wch: 12 },
      { wch: 12 },
      { wch: 18 }
    ]

    const fechaHoy = new Date().toISOString().split('T')[0]
    XLSX.writeFile(libro, `Inventario_StockPulse_${fechaHoy}.xlsx`)
  }

  // --- CONTROL DE CAJA ---
  async function obtenerEstadoCaja() {
    setCargandoCaja(true)
    const { data, error } = await supabase
      .from('sesiones_caja')
      .select('*')
      .eq('estado', 'abierta')
      .order('fecha_apertura', { ascending: false })
      .limit(1)

    if (error) console.error('Error al verificar caja:', error)
    else if (data && data.length > 0) setCajaActual(data[0])
    else setCajaActual(null)
    setCargandoCaja(false)
  }

  async function obtenerHistorialCajas() {
    const { data } = await supabase
      .from('sesiones_caja')
      .select('*')
      .order('fecha_apertura', { ascending: false })
      .limit(10)

    if (data) setHistorialCajas(data)
  }

  async function abrirCaja(e) {
    e.preventDefault()
    const inicial = parseFloat(inputMontoInicial) || 0
    setProcesandoAperturaCaja(true)

    const { data, error } = await supabase
      .from('sesiones_caja')
      .insert([{
        monto_inicial: inicial,
        estado: 'abierta',
        total_ventas_efectivo: 0,
        total_ventas_tarjeta: 0
      }])
      .select()

    if (error) {
      alert('Error al abrir la caja.')
    } else if (data) {
      setCajaActual(data[0])
      setInputMontoInicial('')
      obtenerHistorialCajas()
    }
    setProcesandoAperturaCaja(false)
  }

  async function cerrarCaja(e) {
    e.preventDefault()
    if (!cajaActual) return

    setProcesandoCierreCaja(true)
    const contadoFisico = parseFloat(inputMontoArqueo) || 0
    const esperadoEfectivo = Number(cajaActual.monto_inicial) + Number(cajaActual.total_ventas_efectivo || 0)
    const diferencia = contadoFisico - esperadoEfectivo

    const { error } = await supabase
      .from('sesiones_caja')
      .update({
        estado: 'cerrada',
        monto_final_efectivo: contadoFisico,
        monto_esperado_efectivo: esperadoEfectivo,
        diferencia: diferencia,
        fecha_cierre: new Date().toISOString()
      })
      .eq('id', cajaActual.id)

    if (error) {
      alert('Error al cerrar la caja.')
    } else {
      alert(`Caja cerrada. ${diferencia === 0 ? '¡Arqueo perfecto!' : diferencia > 0 ? `Sobrante: $${diferencia.toFixed(2)}` : `Faltante: $${Math.abs(diferencia).toFixed(2)}`}`)
      setCajaActual(null)
      setInputMontoArqueo('')
      obtenerHistorialCajas()
    }
    setProcesandoCierreCaja(false)
  }

  // --- CLIENTES & RUT ---
  async function obtenerClientes() {
    const { data, error } = await supabase
      .from('clientes')
      .select('*')
      .order('nombre', { ascending: true })

    if (error) console.error('Error al obtener clientes:', error)
    else setClientes(data || [])
    setCargandoClientes(false)
  }

  async function buscarYAutocompletarRUT(rutIngresado) {
    const rutLimpio = rutIngresado.trim()
    setInputRUTPOS(rutLimpio)

    if (rutLimpio.length < 12) return

    const { data } = await supabase
      .from('clientes')
      .select('*')
      .eq('rut_ci', rutLimpio)
      .maybeSingle()

    if (data) {
      setClienteSeleccionadoPOS(data)
      return
    }

    if (MOCK_EMPRESAS_DGI[rutLimpio]) {
      const empresaDGI = MOCK_EMPRESAS_DGI[rutLimpio]
      
      const { data: nuevoCliente } = await supabase
        .from('clientes')
        .insert([{
          nombre: empresaDGI.nombre,
          rut_ci: rutLimpio,
          direccion: empresaDGI.direccion,
          telefono: empresaDGI.telefono
        }])
        .select()

      if (nuevoCliente) {
        setClienteSeleccionadoPOS(nuevoCliente[0])
        obtenerClientes()
      }
    } else {
      setClienteSeleccionadoPOS({
        nombre: `EMPRESA RUT ${rutLimpio}`,
        rut_ci: rutLimpio
      })
    }
  }

  async function agregarCliente(e) {
    e.preventDefault()
    if (!formCliente.nombre.trim()) return

    setGuardandoCliente(true)
    const { data, error } = await supabase
      .from('clientes')
      .insert([formCliente])
      .select()

    if (error) alert('Error al agregar cliente.')
    else if (data) {
      setClientes((prev) => [...prev, ...data])
      setFormCliente({ nombre: '', rut_ci: '', telefono: '', direccion: '' })
      setModalClienteAbierto(false)
    }
    setGuardandoCliente(false)
  }

  // --- CATEGORÍAS ---
  async function obtenerCategorias() {
    const { data, error } = await supabase
      .from('categorias')
      .select('id, nombre')
      .order('nombre', { ascending: true })

    if (error) console.error('Error al obtener categorías:', error)
    else setCategorias(data || [])
    setCargandoCategorias(false)
  }

  async function agregarCategoria(e) {
    e.preventDefault()
    if (!nuevoNombreCat.trim()) return

    setGuardandoCat(true)
    const { data, error } = await supabase
      .from('categorias')
      .insert([{ nombre: nuevoNombreCat.trim() }])
      .select()

    if (error) alert('No se pudo guardar la categoría.')
    else if (data) {
      setCategorias((prev) => [...prev, ...data].sort((a, b) => a.nombre.localeCompare(b.nombre)))
      setNuevoNombreCat('')
      setModalCategoriaAbierto(false)
    }
    setGuardandoCat(false)
  }

  async function guardarEdicionCat(id) {
    if (!nombreCatEditado.trim()) return
    const { error } = await supabase
      .from('categorias')
      .update({ nombre: nombreCatEditado.trim() })
      .eq('id', id)

    if (error) alert('No se pudo actualizar la categoría.')
    else {
      setCategorias((prev) =>
        prev
          .map((c) => (c.id === id ? { ...c, nombre: nombreCatEditado.trim() } : c))
          .sort((a, b) => a.nombre.localeCompare(b.nombre))
      )
      setEditandoCatId(null)
    }
  }

  async function eliminarCategoria(id, nombre) {
    if (!confirm(`¿Eliminar categoría "${nombre}"? Los productos asociados perderán su categoría.`)) return
    const { error } = await supabase.from('categorias').delete().eq('id', id)
    if (error) alert('No se pudo eliminar la categoría.')
    else {
      setCategorias((prev) => prev.filter((c) => c.id !== id))
      obtenerProductos()
    }
  }

  // --- PRODUCTOS ---
  async function obtenerProductos() {
    const { data, error } = await supabase
      .from('productos')
      .select('id, nombre, codigo_barra, precio, stock, categoria_id, tipo_venta, categorias(nombre)')
      .order('id', { ascending: false })

    if (error) console.error('Error al obtener productos:', error)
    else setProductos(data || [])
    setCargandoProductos(false)
  }

  async function obtenerLotes() {
    const { data, error } = await supabase
      .from('lotes_vencimiento')
      .select('id, producto_id, cantidad, fecha_vencimiento')
      .order('fecha_vencimiento', { ascending: true })

    if (error) console.error('Error al obtener lotes:', error)
    else setLotes(data || [])
  }

  async function obtenerVentas() {
    const { data, error } = await supabase
      .from('ventas')
      .select('id, total, metodo_pago, estado, created_at, clientes(nombre, rut_ci), detalle_ventas(cantidad, precio_unitario, subtotal, productos(nombre, tipo_venta))')
      .order('created_at', { ascending: false })

    if (error) console.error('Error al obtener ventas:', error)
    else setVentas(data || [])
  }

  async function agregarProducto(e) {
    e.preventDefault()
    if (!formProd.nombre.trim()) return

    setGuardandoProd(true)
    const nuevo = {
      nombre: formProd.nombre.trim(),
      precio: formProd.precio ? parseFloat(formProd.precio) : 0,
      stock: formProd.stock ? parseFloat(formProd.stock) : 0,
      tipo_venta: formProd.tipo_venta || 'unidad'
    }

    if (formProd.codigo_barra.trim()) nuevo.codigo_barra = formProd.codigo_barra.trim()
    if (formProd.categoria_id) nuevo.categoria_id = parseInt(formProd.categoria_id, 10)

    const { error } = await supabase.from('productos').insert([nuevo])

    if (error) alert(`Error al guardar: ${error.message}`)
    else {
      setFormProd({ nombre: '', codigo_barra: '', precio: '', stock: '', categoria_id: '', tipo_venta: 'unidad' })
      setModalProductoAbierto(false)
      obtenerProductos()
    }
    setGuardandoProd(false)
  }

  async function agregarLote(e) {
    e.preventDefault()
    if (!prodSeleccionadoLote || !nuevoLote.fecha_vencimiento) return

    setGuardandoLote(true)
    const cantidad = parseFloat(nuevoLote.cantidad) || 1

    const { error } = await supabase
      .from('lotes_vencimiento')
      .insert([{
        producto_id: prodSeleccionadoLote.id,
        cantidad: cantidad,
        fecha_vencimiento: nuevoLote.fecha_vencimiento
      }])

    if (error) alert(`Error al agregar lote: ${error.message}`)
    else {
      setNuevoLote({ cantidad: '', fecha_vencimiento: '' })
      obtenerLotes()
    }
    setGuardandoLote(false)
  }

  async function eliminarLote(id) {
    const { error } = await supabase.from('lotes_vencimiento').delete().eq('id', id)
    if (error) alert('No se pudo eliminar el lote.')
    else setLotes((prev) => prev.filter((l) => l.id !== id))
  }

  async function modificarStockRapido(id, stockActual, cambio) {
    const nuevoStock = Math.max(0, stockActual + cambio)
    const { error } = await supabase
      .from('productos')
      .update({ stock: nuevoStock })
      .eq('id', id)

    if (error) console.error('Error al actualizar stock:', error)
    else {
      setProductos((prev) =>
        prev.map((p) => (p.id === id ? { ...p, stock: nuevoStock } : p))
      )
    }
  }

  function iniciarEdicionProd(p) {
    setEditandoProdId(p.id)
    setFormProdEdit({
      nombre: p.nombre,
      codigo_barra: p.codigo_barra || '',
      precio: p.precio,
      stock: p.stock,
      categoria_id: p.categoria_id || '',
      tipo_venta: p.tipo_venta || 'unidad'
    })
  }

  async function guardarEdicionProd(id) {
    const actualizado = {
      nombre: formProdEdit.nombre.trim(),
      codigo_barra: formProdEdit.codigo_barra.trim() || null,
      precio: parseFloat(formProdEdit.precio) || 0,
      stock: parseFloat(formProdEdit.stock) || 0,
      categoria_id: formProdEdit.categoria_id ? parseInt(formProdEdit.categoria_id, 10) : null,
      tipo_venta: formProdEdit.tipo_venta
    }

    const { error } = await supabase
      .from('productos')
      .update(actualizado)
      .eq('id', id)

    if (error) alert('Error al actualizar el producto.')
    else {
      setEditandoProdId(null)
      obtenerProductos()
    }
  }

  async function eliminarProducto(id, nombre) {
    if (!confirm(`¿Eliminar el producto "${nombre}"?`)) return
    const { error } = await supabase.from('productos').delete().eq('id', id)
    if (error) alert('No se pudo eliminar el producto.')
    else setProductos((prev) => prev.filter((p) => p.id !== id))
  }

  // --- CARRITO & VENTA ---
  function solicitarAgregarAlCarrito(producto) {
    if (!cajaActual) {
      alert('¡Debes abrir la caja antes de registrar ventas!')
      setPestana('caja')
      return
    }

    if (producto.stock <= 0) {
      alert('Producto agotado.')
      return
    }

    if (producto.tipo_venta === 'peso') {
      setProductoParaPeso(producto)
      setInputGramos('')
      setInputMontoDinero('')
      setModalPesoAbierto(true)
    } else {
      agregarUnidadAlCarrito(producto)
    }
  }

  function agregarUnidadAlCarrito(producto) {
    const existe = carrito.find((item) => item.id === producto.id)
    if (existe) {
      if (existe.cantidadEnCarrito >= producto.stock) {
        alert('No hay suficiente stock disponible.')
        return
      }
      setCarrito(
        carrito.map((item) =>
          item.id === producto.id
            ? { ...item, cantidadEnCarrito: item.cantidadEnCarrito + 1 }
            : item
        )
      )
    } else {
      setCarrito([...carrito, { ...producto, cantidadEnCarrito: 1, subtotal: producto.precio }])
    }
  }

  function agregarPesoAlCarrito(e) {
    e.preventDefault()
    if (!productoParaPeso) return

    const gramos = parseFloat(inputGramos)
    if (!gramos || gramos <= 0) {
      alert('Ingresa una cantidad válida en gramos.')
      return
    }

    const kg = gramos / 1000
    const subtotal = kg * productoParaPeso.precio

    const itemEnCarrito = {
      ...productoParaPeso,
      cantidadEnCarrito: kg,
      gramosOriginales: gramos,
      subtotal: subtotal
    }

    setCarrito((prev) => [...prev, itemEnCarrito])
    setModalPesoAbierto(false)
    setProductoParaPeso(null)
  }

  function manejarCambioGramos(val) {
    setInputGramos(val)
    if (!val || isNaN(val) || !productoParaPeso) {
      setInputMontoDinero('')
      return
    }
    const montoCalculado = (parseFloat(val) / 1000) * productoParaPeso.precio
    setInputMontoDinero(montoCalculado.toFixed(2))
  }

  function manejarCambioMontoDinero(val) {
    setInputMontoDinero(val)
    if (!val || isNaN(val) || !productoParaPeso || productoParaPeso.precio === 0) {
      setInputGramos('')
      return
    }
    const gramosCalculados = (parseFloat(val) / productoParaPeso.precio) * 1000
    setInputGramos(Math.round(gramosCalculados).toString())
  }

  function eliminarDelCarrito(index) {
    setCarrito(carrito.filter((_, i) => i !== index))
  }

  function modificarCantidadCarrito(index, cambio) {
    setCarrito(
      carrito
        .map((item, i) => {
          if (i === index) {
            if (item.tipo_venta === 'peso') return item
            const nuevaCant = item.cantidadEnCarrito + cambio
            if (nuevaCant > item.stock) {
              alert('Supera el stock disponible.')
              return item
            }
            return nuevaCant > 0 ? { ...item, cantidadEnCarrito: nuevaCant, subtotal: nuevaCant * item.precio } : null
          }
          return item
        })
        .filter(Boolean)
    )
  }

  const totalCarrito = carrito.reduce((sum, i) => sum + (i.subtotal || i.precio * i.cantidadEnCarrito), 0)

  async function completarVenta() {
    if (carrito.length === 0) return
    if (!cajaActual) {
      alert('La caja está cerrada. Debes abrirla antes de cobrar.')
      setPestana('caja')
      return
    }

    setProcesandoVenta(true)

    const ventaObjeto = {
      total: totalCarrito,
      metodo_pago: metodoPago,
      estado: 'completada',
      cliente_id: clienteSeleccionadoPOS ? clienteSeleccionadoPOS.id : null,
      caja_id: cajaActual.id
    }

    const { data: ventaData, error: ventaError } = await supabase
      .from('ventas')
      .insert([ventaObjeto])
      .select('*, clientes(nombre, rut_ci)')

    if (ventaError || !ventaData) {
      alert('Error al procesar la venta.')
      setProcesandoVenta(false)
      return
    }

    const ventaGenerada = ventaData[0]

    const detalles = carrito.map((item) => ({
      venta_id: ventaGenerada.id,
      producto_id: item.id,
      cantidad: item.cantidadEnCarrito,
      precio_unitario: item.precio,
      subtotal: item.subtotal || item.precio * item.cantidadEnCarrito
    }))

    await supabase.from('detalle_ventas').insert(detalles)

    for (const item of carrito) {
      const nuevoStock = Math.max(0, item.stock - item.cantidadEnCarrito)
      await supabase
        .from('productos')
        .update({ stock: nuevoStock })
        .eq('id', item.id)
    }

    const esEfectivo = metodoPago === 'Efectivo'
    const nuevoTotalEfectivo = Number(cajaActual.total_ventas_efectivo || 0) + (esEfectivo ? totalCarrito : 0)
    const nuevoTotalTarjeta = Number(cajaActual.total_ventas_tarjeta || 0) + (!esEfectivo ? totalCarrito : 0)

    await supabase
      .from('sesiones_caja')
      .update({
        total_ventas_efectivo: nuevoTotalEfectivo,
        total_ventas_tarjeta: nuevoTotalTarjeta
      })
      .eq('id', cajaActual.id)

    setCajaActual({
      ...cajaActual,
      total_ventas_efectivo: nuevoTotalEfectivo,
      total_ventas_tarjeta: nuevoTotalTarjeta
    })

    setUltimaVentaRealizada({
      ...ventaGenerada,
      items: [...carrito]
    })

    setCarrito([])
    setClienteSeleccionadoPOS(null)
    setInputRUTPOS('')
    setProcesandoVenta(false)
    obtenerProductos()
    obtenerVentas()
    setModalTicketAbierto(true)
  }

  async function anularVenta(ventaId) {
    if (!confirm('¿Estás seguro de anular esta venta? Se devolverá el stock a los productos correspondientes.')) return

    setProcesandoAnulacion(ventaId)

    const { data: detalles, error: errorDetalles } = await supabase
      .from('detalle_ventas')
      .select('producto_id, cantidad')
      .eq('venta_id', ventaId)

    if (errorDetalles) {
      alert('No se pudieron obtener los detalles de la venta.')
      setProcesandoAnulacion(null)
      return
    }

    if (detalles) {
      for (const d of detalles) {
        if (d.producto_id) {
          const prodActual = productos.find((p) => p.id === d.producto_id)
          if (prodActual) {
            const nuevoStock = prodActual.stock + d.cantidad
            await supabase
              .from('productos')
              .update({ stock: nuevoStock })
              .eq('id', d.producto_id)
          }
        }
      }
    }

    const { error: errorVenta } = await supabase
      .from('ventas')
      .update({ estado: 'anulada' })
      .eq('id', ventaId)

    if (errorVenta) {
      alert('Error al anular la venta.')
    } else {
      alert('Venta anulada correctamente y stock devuelto.')
      obtenerProductos()
      obtenerVentas()
    }

    setProcesandoAnulacion(null)
  }

  function calcularDiasRestantes(fechaStr) {
    const hoy = new Date()
    hoy.setHours(0, 0, 0, 0)
    
    const [year, month, day] = fechaStr.split('-').map(Number)
    const fechaVenc = new Date(year, month - 1, day)
    fechaVenc.setHours(0, 0, 0, 0)

    const diferenciaMs = fechaVenc - hoy
    return Math.ceil(diferenciaMs / (1000 * 60 * 60 * 24))
  }

  const hoyStr = new Date().toISOString().split('T')[0]
  const enSieteDias = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]

  const lotesVencidos = lotes.filter((l) => l.fecha_vencimiento < hoyStr)
  const lotesPorVencer = lotes.filter((l) => l.fecha_vencimiento >= hoyStr && l.fecha_vencimiento <= enSieteDias)

  const productosBajoStock = productos.filter((p) => p.stock > 0 && p.stock <= 5)
  const productosAgotados = productos.filter((p) => p.stock === 0)

  const productosFiltrados = productos.filter((p) => {
    const termino = busqueda.toLowerCase().trim()
    const coincideNombre = p.nombre.toLowerCase().includes(termino)
    const coincideCodigo = p.codigo_barra ? p.codigo_barra.toLowerCase().includes(termino) : false
    const coincideBusqueda = !termino || coincideNombre || coincideCodigo

    if (!coincideBusqueda) return false

    if (filtroStock === 'bajo') return p.stock > 0 && p.stock <= 5
    if (filtroStock === 'agotado') return p.stock === 0
    if (filtroStock === 'vencidos') {
      const prodLotes = lotes.filter((l) => l.producto_id === p.id)
      return prodLotes.some((l) => l.fecha_vencimiento <= enSieteDias)
    }
    return true
  })

  function TextoResumenVencimiento({ prodLotes }) {
    if (!prodLotes || prodLotes.length === 0) return null

    const loteMasProximo = prodLotes[0]
    const dias = calcularDiasRestantes(loteMasProximo.fecha_vencimiento)

    if (dias < 0) return <span className="text-xs font-medium text-red-400 bg-red-950/60 border border-red-800 px-2.5 py-1 rounded-full">Ya venció</span>
    if (dias === 0) return <span className="text-xs font-bold text-red-400 bg-red-950/90 border border-red-600 px-2.5 py-1 rounded-full animate-pulse">Vence hoy</span>
    if (dias <= 7) return <span className="text-xs font-medium text-amber-300 bg-amber-950/60 border border-amber-800 px-2.5 py-1 rounded-full">Faltan {dias} {dias === 1 ? 'día' : 'días'}</span>
    return <span className="text-xs text-purple-300 font-medium bg-purple-950/50 border border-purple-800 px-2.5 py-1 rounded-full">Faltan {dias} días</span>
  }

  const valorTotalInventario = productos.reduce((sum, p) => sum + p.precio * p.stock, 0)
  const ventasValidas = ventas.filter((v) => v.estado !== 'anulada')
  const totalVentasMonto = ventasValidas.reduce((sum, v) => sum + Number(v.total), 0)

  const datosGraficoVentas = ventasValidas.slice(0, 7).reverse().map((v) => ({
    fecha: new Date(v.created_at).toLocaleDateString('es-UY', { day: '2-digit', month: '2-digit' }),
    monto: Number(v.total)
  }))

  return (
    <div className="flex min-h-screen bg-slate-950 text-slate-100 font-sans">
      
      {/* ESTILOS DE IMPRESIÓN TÉRMICA OPTIMIZADOS (72mm EPSON) */}
      <style>{`
        @media print {
          body, html {
            background: #ffffff !important;
            color: #000000 !important;
            margin: 0 !important;
            padding: 0 !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          
          aside, main, nav, .no-imprimir {
            display: none !important;
          }

          #ticket-impresion {
            display: block !important;
            position: absolute !important;
            top: 0 !important;
            left: 0 !important;
            width: 72mm !important;
            margin: 0 !important;
            padding: 2mm !important;
            background: #ffffff !important;
            color: #000000 !important;
            font-family: "Courier New", Courier, monospace !important;
            font-size: 12px !important;
            font-weight: 700 !important;
            line-height: 1.2 !important;
            z-index: 99999 !important;
            -webkit-font-smoothing: none !important;
          }

          #ticket-impresion * {
            color: #000000 !important;
            background: transparent !important;
            font-family: "Courier New", Courier, monospace !important;
            border-color: #000000 !important;
          }
        }
      `}</style>

      {/* BARRA LATERAL (SIDEBAR) */}
      <aside className={`${sidebarColapsado ? 'w-20' : 'w-64'} bg-slate-900 border-r border-slate-800 p-4 flex flex-col justify-between hidden md:flex shrink-0 transition-all duration-300 relative no-imprimir`}>
        <div className="space-y-6">
          
          <div 
            onClick={() => setSidebarColapsado(!sidebarColapsado)}
            className={`flex items-center gap-3 cursor-pointer p-1.5 rounded-xl hover:bg-slate-800/60 transition group ${
              sidebarColapsado ? 'justify-center w-full' : ''
            }`}
            title={sidebarColapsado ? 'Expandir menú' : 'Contraer menú'}
          >
            <div className="p-2 bg-blue-600/20 text-blue-400 rounded-xl border border-blue-500/30 shrink-0 group-hover:scale-105 transition-transform">
              <Package className="w-6 h-6" />
            </div>
            {!sidebarColapsado && (
              <div className="overflow-hidden">
                <h1 className="font-black text-xl text-white tracking-tight">StockPulse</h1>
                <p className="text-[10px] text-slate-400">POS & Inventario</p>
              </div>
            )}
          </div>

          <nav className="space-y-2 pt-2">
            <button
              onClick={() => setPestana('pos')}
              title="Punto de Venta"
              className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-xl font-medium text-sm transition-all duration-200 ${
                sidebarColapsado ? 'justify-center' : ''
              } ${
                pestana === 'pos' ? 'bg-emerald-600 text-white shadow-lg shadow-emerald-600/30' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
            >
              <ShoppingCart className="w-5 h-5 shrink-0" />
              {!sidebarColapsado && <span>Punto de Venta</span>}
            </button>

            <button
              onClick={() => setPestana('caja')}
              title="Control de Caja"
              className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-xl font-medium text-sm transition-all duration-200 ${
                sidebarColapsado ? 'justify-center' : ''
              } ${
                pestana === 'caja' ? 'bg-teal-600 text-white shadow-lg shadow-teal-600/30' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
            >
              <Vault className="w-5 h-5 shrink-0" />
              {!sidebarColapsado && (
                <div className="flex items-center justify-between w-full">
                  <span>Caja / Turno</span>
                  {cajaActual ? (
                    <span className="w-2 h-2 rounded-full bg-emerald-400 animate-ping" />
                  ) : (
                    <span className="w-2 h-2 rounded-full bg-red-400" />
                  )}
                </div>
              )}
            </button>

            <button
              onClick={() => setPestana('productos')}
              title="Inventario"
              className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-xl font-medium text-sm transition-all duration-200 ${
                sidebarColapsado ? 'justify-center' : ''
              } ${
                pestana === 'productos' ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/30' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
            >
              <ShoppingBag className="w-5 h-5 shrink-0" />
              {!sidebarColapsado && <span>Inventario</span>}
            </button>

            <button
              onClick={() => setPestana('clientes')}
              title="Clientes"
              className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-xl font-medium text-sm transition-all duration-200 ${
                sidebarColapsado ? 'justify-center' : ''
              } ${
                pestana === 'clientes' ? 'bg-amber-600 text-white shadow-lg shadow-amber-600/30' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
            >
              <Users className="w-5 h-5 shrink-0" />
              {!sidebarColapsado && <span>Clientes (RUT)</span>}
            </button>

            <button
              onClick={() => setPestana('dashboard')}
              title="Métricas y Ventas"
              className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-xl font-medium text-sm transition-all duration-200 ${
                sidebarColapsado ? 'justify-center' : ''
              } ${
                pestana === 'dashboard' ? 'bg-purple-600 text-white shadow-lg shadow-purple-600/30' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
            >
              <LayoutDashboard className="w-5 h-5 shrink-0" />
              {!sidebarColapsado && <span>Métricas y Ventas</span>}
            </button>

            <button
              onClick={() => setPestana('categorias')}
              title="Categorías"
              className={`w-full flex items-center gap-3 px-3.5 py-3 rounded-xl font-medium text-sm transition-all duration-200 ${
                sidebarColapsado ? 'justify-center' : ''
              } ${
                pestana === 'categorias' ? 'bg-slate-800 text-white' : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
            >
              <Layers className="w-5 h-5 shrink-0" />
              {!sidebarColapsado && <span>Categorías</span>}
            </button>
          </nav>
        </div>

        <div className="pt-4 border-t border-slate-800">
          <div className={`flex items-center space-x-2 text-emerald-400 bg-emerald-950/40 p-2.5 rounded-xl border border-emerald-800/50 text-xs font-semibold ${sidebarColapsado ? 'justify-center' : ''}`}>
            <CheckCircle2 className="w-4 h-4 shrink-0" />
            {!sidebarColapsado && <span>Supabase Activo</span>}
          </div>
        </div>
      </aside>

      {/* CONTENIDO PRINCIPAL */}
      <main className="flex-1 p-4 md:p-8 overflow-y-auto max-w-7xl mx-auto space-y-6 w-full no-imprimir">

        {/* Navegación Móvil */}
        <div className="md:hidden flex bg-slate-900 p-1.5 rounded-2xl border border-slate-800 gap-1 overflow-x-auto">
          <button onClick={() => setPestana('pos')} className={`flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium whitespace-nowrap ${pestana === 'pos' ? 'bg-emerald-600 text-white' : 'text-slate-400'}`}>
            <ShoppingCart className="w-3.5 h-3.5" /> POS
          </button>
          <button onClick={() => setPestana('caja')} className={`flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium whitespace-nowrap ${pestana === 'caja' ? 'bg-teal-600 text-white' : 'text-slate-400'}`}>
            <Vault className="w-3.5 h-3.5" /> Caja
          </button>
          <button onClick={() => setPestana('productos')} className={`flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium whitespace-nowrap ${pestana === 'productos' ? 'bg-blue-600 text-white' : 'text-slate-400'}`}>
            <ShoppingBag className="w-3.5 h-3.5" /> Inventario
          </button>
          <button onClick={() => setPestana('clientes')} className={`flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium whitespace-nowrap ${pestana === 'clientes' ? 'bg-amber-600 text-white' : 'text-slate-400'}`}>
            <Users className="w-3.5 h-3.5" /> Clientes
          </button>
          <button onClick={() => setPestana('dashboard')} className={`flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium whitespace-nowrap ${pestana === 'dashboard' ? 'bg-purple-600 text-white' : 'text-slate-400'}`}>
            <LayoutDashboard className="w-3.5 h-3.5" /> Métricas
          </button>
          <button onClick={() => setPestana('categorias')} className={`flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium whitespace-nowrap ${pestana === 'categorias' ? 'bg-slate-800 text-white' : 'text-slate-400'}`}>
            <Layers className="w-3.5 h-3.5" /> Categorías
          </button>
        </div>

        {/* VISTA: CONTROL DE CAJA / TURNO */}
        {pestana === 'caja' && (
          <div className="space-y-6">
            <div className="flex justify-between items-center bg-slate-900/80 p-6 rounded-2xl border border-slate-800">
              <div>
                <h2 className="text-xl font-bold text-white flex items-center gap-2">
                  <Vault className="w-6 h-6 text-teal-400" /> Control de Turno y Caja
                </h2>
                <p className="text-xs text-slate-400 mt-1">Gestión de aperturas, cierres de turno y arqueo de efectivo</p>
              </div>

              {cajaActual ? (
                <div className="flex items-center gap-2 bg-emerald-950/80 text-emerald-400 border border-emerald-800 px-3.5 py-1.5 rounded-xl text-xs font-bold">
                  <Unlock className="w-4 h-4" /> CAJA ABIERTA (Turno #{cajaActual.id})
                </div>
              ) : (
                <div className="flex items-center gap-2 bg-red-950/80 text-red-400 border border-red-800 px-3.5 py-1.5 rounded-xl text-xs font-bold">
                  <Lock className="w-4 h-4" /> CAJA CERRADA
                </div>
              )}
            </div>

            {!cajaActual ? (
              <div className="bg-slate-900/60 p-8 rounded-2xl border border-slate-800 max-w-lg mx-auto text-center space-y-4 shadow-xl">
                <div className="p-4 bg-teal-500/10 text-teal-400 rounded-2xl w-16 h-16 mx-auto flex items-center justify-center">
                  <Coins className="w-8 h-8" />
                </div>
                <h3 className="text-lg font-bold text-white">Apertura de Caja</h3>
                <p className="text-xs text-slate-400">Ingresá el cambio inicial en efectivo disponible en el cajón para comenzar el turno.</p>

                <form onSubmit={abrirCaja} className="space-y-4 pt-2">
                  <div>
                    <label className="text-xs text-slate-400 font-semibold mb-1 block text-left">Fondo de Caja / Cambio ($)</label>
                    <input
                      type="number"
                      step="0.01"
                      placeholder="Ej: 2000"
                      value={inputMontoInicial}
                      onChange={(e) => setInputMontoInicial(e.target.value)}
                      className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-lg font-bold text-teal-400 focus:outline-none focus:border-teal-500 text-center"
                      required
                    />
                  </div>

                  <button
                    type="submit"
                    disabled={procesandoAperturaCaja}
                    className="w-full bg-teal-600 hover:bg-teal-500 text-white font-bold py-3 rounded-xl flex items-center justify-center gap-2 transition shadow-lg shadow-teal-600/20"
                  >
                    {procesandoAperturaCaja ? <Loader2 className="w-5 h-5 animate-spin" /> : <Unlock className="w-5 h-5" />}
                    Abrir Caja y Comenzar Turno
                  </button>
                </form>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="space-y-4">
                  <div className="bg-slate-900/80 p-5 rounded-2xl border border-slate-800">
                    <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Fondo Inicial</p>
                    <p className="text-2xl font-black text-slate-100 mt-1">${Number(cajaActual.monto_inicial).toFixed(2)}</p>
                  </div>

                  <div className="bg-slate-900/80 p-5 rounded-2xl border border-slate-800">
                    <p className="text-xs font-bold text-emerald-400 uppercase tracking-wider">Ventas en Efectivo</p>
                    <p className="text-2xl font-black text-emerald-400 mt-1">${Number(cajaActual.total_ventas_efectivo || 0).toFixed(2)}</p>
                  </div>

                  <div className="bg-slate-900/80 p-5 rounded-2xl border border-slate-800">
                    <p className="text-xs font-bold text-purple-400 uppercase tracking-wider">Ventas Digitales / Tarjeta</p>
                    <p className="text-2xl font-black text-purple-400 mt-1">${Number(cajaActual.total_ventas_tarjeta || 0).toFixed(2)}</p>
                  </div>
                </div>

                <div className="md:col-span-2 bg-slate-900/80 p-6 rounded-2xl border border-slate-800 flex flex-col justify-between space-y-4">
                  <div className="space-y-4">
                    <h3 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
                      <Lock className="w-5 h-5 text-red-400" /> Cierre de Turno y Arqueo
                    </h3>

                    <div className="bg-slate-950 p-4 rounded-xl border border-slate-800 space-y-2">
                      <div className="flex justify-between text-xs text-slate-400">
                        <span>Fondo Inicial:</span>
                        <span className="font-bold text-slate-200">${Number(cajaActual.monto_inicial).toFixed(2)}</span>
                      </div>
                      <div className="flex justify-between text-xs text-slate-400">
                        <span>(+) Ventas Efectivo:</span>
                        <span className="font-bold text-emerald-400">+${Number(cajaActual.total_ventas_efectivo || 0).toFixed(2)}</span>
                      </div>
                      <div className="flex justify-between text-sm font-bold border-t border-slate-800 pt-2 text-slate-100">
                        <span>Total Esperado en Cajón:</span>
                        <span className="text-teal-400">${(Number(cajaActual.monto_inicial) + Number(cajaActual.total_ventas_efectivo || 0)).toFixed(2)}</span>
                      </div>
                    </div>

                    <form onSubmit={cerrarCaja} className="space-y-4">
                      <div>
                        <label className="text-xs text-slate-400 font-semibold mb-1 block">Efectivo Real Contado en el Cajón ($)</label>
                        <input
                          type="number"
                          step="0.01"
                          placeholder="Ingresá cuánto dinero hay físicamente..."
                          value={inputMontoArqueo}
                          onChange={(e) => setInputMontoArqueo(e.target.value)}
                          className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-lg font-bold text-emerald-400 focus:outline-none focus:border-red-500"
                          required
                        />
                      </div>

                      {inputMontoArqueo !== '' && (
                        <div className="bg-slate-950 p-3 rounded-xl border border-slate-800 text-xs flex justify-between items-center">
                          <span className="text-slate-400">Diferencia de Arqueo:</span>
                          {(() => {
                            const dif = (parseFloat(inputMontoArqueo) || 0) - (Number(cajaActual.monto_inicial) + Number(cajaActual.total_ventas_efectivo || 0))
                            if (dif === 0) return <span className="font-bold text-emerald-400">Cuadra exacto $0.00</span>
                            if (dif > 0) return <span className="font-bold text-blue-400">Sobrante: +${dif.toFixed(2)}</span>
                            return <span className="font-bold text-red-400">Faltante: -${Math.abs(dif).toFixed(2)}</span>
                          })()}
                        </div>
                      )}

                      <button
                        type="submit"
                        disabled={procesandoCierreCaja}
                        className="w-full bg-red-600 hover:bg-red-500 text-white font-bold py-3 rounded-xl flex items-center justify-center gap-2 transition shadow-lg shadow-red-600/20"
                      >
                        {procesandoCierreCaja ? <Loader2 className="w-5 h-5 animate-spin" /> : <Lock className="w-5 h-5" />}
                        Cerrar Turno de Caja
                      </button>
                    </form>
                  </div>
                </div>

              </div>
            )}

            <div className="bg-slate-900/80 p-6 rounded-2xl border border-slate-800 space-y-4">
              <h3 className="text-sm font-bold text-slate-200 uppercase tracking-wider flex items-center gap-2">
                <History className="w-4 h-4 text-teal-400" /> Historial de Turnos de Caja
              </h3>

              {historialCajas.length === 0 ? (
                <p className="text-xs text-slate-500 italic">No hay registros previos de caja.</p>
              ) : (
                <div className="space-y-2">
                  {historialCajas.map((c) => (
                    <div key={c.id} className="flex flex-wrap items-center justify-between bg-slate-950 p-3.5 rounded-xl border border-slate-800 text-xs">
                      <div className="space-y-0.5">
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-slate-200">Turno #{c.id}</span>
                          {c.estado === 'abierta' ? (
                            <span className="bg-emerald-950 text-emerald-400 text-[10px] font-bold px-2 py-0.5 rounded border border-emerald-800">EN CURSO</span>
                          ) : (
                            <span className="bg-slate-800 text-slate-400 text-[10px] font-bold px-2 py-0.5 rounded">CERRADA</span>
                          )}
                        </div>
                        <p className="text-[11px] text-slate-500">
                          Apertura: {new Date(c.fecha_apertura).toLocaleString('es-UY')}
                        </p>
                      </div>

                      <div className="flex items-center gap-6">
                        <div>
                          <p className="text-[10px] text-slate-500">Efectivo Cobrado</p>
                          <p className="font-bold text-emerald-400">${Number(c.total_ventas_efectivo || 0).toFixed(2)}</p>
                        </div>
                        <div>
                          <p className="text-[10px] text-slate-500">Digital / Tarjeta</p>
                          <p className="font-bold text-purple-400">${Number(c.total_ventas_tarjeta || 0).toFixed(2)}</p>
                        </div>
                        {c.estado === 'cerrada' && (
                          <div>
                            <p className="text-[10px] text-slate-500">Diferencia</p>
                            <p className={`font-bold ${Number(c.diferencia) === 0 ? 'text-slate-400' : Number(c.diferencia) > 0 ? 'text-blue-400' : 'text-red-400'}`}>
                              ${Number(c.diferencia || 0).toFixed(2)}
                            </p>
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* VISTA: POS */}
        {pestana === 'pos' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-2 space-y-4">
              
              {!cajaActual && (
                <div className="bg-amber-950/60 border border-amber-800 text-amber-200 p-4 rounded-2xl flex justify-between items-center text-xs">
                  <div className="flex items-center gap-2">
                    <AlertTriangle className="w-5 h-5 text-amber-400 shrink-0" />
                    <span><strong>Atención:</strong> La caja se encuentra cerrada. Abrila para poder habilitar ventas.</span>
                  </div>
                  <button onClick={() => setPestana('caja')} className="bg-amber-600 hover:bg-amber-500 text-white font-bold px-3 py-1.5 rounded-xl transition">
                    Abrir Caja
                  </button>
                </div>
              )}

              <div className="relative">
                <Search className="w-5 h-5 text-slate-400 absolute left-3.5 top-3" />
                <input
                  type="text"
                  placeholder="Escanear código o buscar producto..."
                  value={busquedaPos}
                  onChange={(e) => setBusquedaPos(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-800 rounded-2xl pl-11 pr-4 py-3 text-sm text-white focus:outline-none focus:border-emerald-500 transition"
                  autoFocus
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 max-h-[550px] overflow-y-auto pr-1">
                {productos
                  .filter((p) => {
                    const term = busquedaPos.toLowerCase().trim()
                    return !term || p.nombre.toLowerCase().includes(term) || (p.codigo_barra && p.codigo_barra.includes(term))
                  })
                  .map((p) => (
                    <div
                      key={p.id}
                      onClick={() => solicitarAgregarAlCarrito(p)}
                      className="bg-slate-900/60 hover:bg-slate-800/80 p-4 rounded-2xl border border-slate-800 cursor-pointer transition flex items-center justify-between group"
                    >
                      <div>
                        <div className="flex items-center gap-2">
                          <p className="font-bold text-sm text-slate-100 group-hover:text-emerald-400 transition">{p.nombre}</p>
                          {p.tipo_venta === 'peso' && (
                            <span className="text-[10px] bg-purple-950 text-purple-300 border border-purple-800 px-1.5 py-0.5 rounded flex items-center gap-0.5">
                              <Scale className="w-3 h-3" /> Peso
                            </span>
                          )}
                        </div>
                        <p className="text-xs text-slate-400 mt-1">Stock: {p.stock} {p.tipo_venta === 'peso' ? 'Kg' : 'u.'}</p>
                      </div>
                      <span className="font-black text-emerald-400 text-base">${Number(p.precio).toFixed(2)} {p.tipo_venta === 'peso' ? '/ Kg' : ''}</span>
                    </div>
                  ))}
              </div>
            </div>

            {/* CARRITO POS */}
            <div className="bg-slate-900/90 p-5 rounded-2xl border border-slate-800 flex flex-col justify-between shadow-xl space-y-4">
              <div>
                <h3 className="font-bold text-lg text-slate-100 border-b border-slate-800 pb-3 mb-3 flex items-center gap-2">
                  <ShoppingCart className="w-5 h-5 text-emerald-400" /> Carrito
                </h3>

                <div className="space-y-1.5 mb-4">
                  <label className="text-[11px] text-slate-400 font-bold uppercase tracking-wider block">
                    Buscar RUT / C.I. de Cliente
                  </label>
                  
                  <input
                    type="text"
                    placeholder="Escribí el RUT (ej: 211234560018)..."
                    maxLength={12}
                    value={inputRUTPOS}
                    onChange={(e) => buscarYAutocompletarRUT(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-amber-300 font-mono focus:border-amber-500 focus:outline-none"
                  />

                  {clienteSeleccionadoPOS && (
                    <div className="bg-amber-950/40 border border-amber-800/60 p-2 rounded-xl text-xs text-amber-200 mt-2 flex justify-between items-center">
                      <div>
                        <p className="font-bold text-slate-100">{clienteSeleccionadoPOS.nombre}</p>
                        <p className="text-[10px] text-amber-400 font-mono">RUT: {clienteSeleccionadoPOS.rut_ci}</p>
                      </div>
                      <button 
                        onClick={() => { setClienteSeleccionadoPOS(null); setInputRUTPOS(''); }}
                        className="text-amber-400 hover:text-white text-[10px] bg-amber-900/60 px-2 py-1 rounded-lg border border-amber-800"
                      >
                        Quitar
                      </button>
                    </div>
                  )}
                </div>

                {carrito.length === 0 ? (
                  <p className="text-xs text-slate-500 italic text-center py-8">Haz clic en un producto para iniciar la venta.</p>
                ) : (
                  <div className="space-y-2.5 max-h-[220px] overflow-y-auto pr-1">
                    {carrito.map((item, index) => (
                      <div key={index} className="flex items-center justify-between bg-slate-950 p-3 rounded-xl border border-slate-800">
                        <div>
                          <p className="text-xs font-bold text-slate-200">{item.nombre}</p>
                          {item.tipo_venta === 'peso' ? (
                            <p className="text-[11px] text-purple-400 font-medium">{item.gramosOriginales}g ({item.cantidadEnCarrito} Kg) &times; ${item.precio}/Kg</p>
                          ) : (
                            <p className="text-[11px] text-emerald-400 font-semibold">${Number(item.precio).toFixed(2)} c/u</p>
                          )}
                        </div>

                        <div className="flex items-center gap-3">
                          <span className="text-xs font-black text-emerald-400">${(item.subtotal || item.precio * item.cantidadEnCarrito).toFixed(2)}</span>

                          {item.tipo_venta === 'peso' ? (
                            <button onClick={() => eliminarDelCarrito(index)} className="text-slate-500 hover:text-red-400 p-1"><Trash2 className="w-4 h-4" /></button>
                          ) : (
                            <div className="flex items-center gap-1.5">
                              <button onClick={() => modificarCantidadCarrito(index, -1)} className="text-slate-400 hover:text-red-400"><MinusCircle className="w-4 h-4" /></button>
                              <span className="text-xs font-bold text-white px-1">{item.cantidadEnCarrito}</span>
                              <button onClick={() => modificarCantidadCarrito(index, 1)} className="text-slate-400 hover:text-emerald-400"><PlusCircle className="w-4 h-4" /></button>
                            </div>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* OPCIONES DE PAGO Y COBRO */}
              <div className="border-t border-slate-800 pt-3 space-y-3">
                <div>
                  <label className="text-[11px] text-slate-400 font-bold uppercase tracking-wider mb-1 block">Modo de Pago</label>
                  <div className="grid grid-cols-2 gap-2">
                    {['Efectivo', 'Débito', 'Crédito', 'Transferencia'].map((m) => (
                      <button
                        key={m}
                        type="button"
                        onClick={() => setMetodoPago(m)}
                        className={`py-2 px-3 rounded-xl text-xs font-semibold border transition ${
                          metodoPago === m
                            ? 'bg-emerald-600 text-white border-emerald-500 shadow-md shadow-emerald-950'
                            : 'bg-slate-950 text-slate-400 border-slate-800 hover:border-slate-700'
                        }`}
                      >
                        {m}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="flex justify-between items-center pt-1">
                  <span className="text-slate-400 text-sm font-semibold">Total a cobrar:</span>
                  <span className="text-emerald-400 font-black text-3xl">${totalCarrito.toFixed(2)}</span>
                </div>

                <button
                  onClick={completarVenta}
                  disabled={carrito.length === 0 || procesandoVenta || !cajaActual}
                  className="w-full bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 text-white font-bold py-3.5 rounded-xl flex items-center justify-center gap-2 transition shadow-lg shadow-emerald-600/20"
                >
                  {procesandoVenta ? <Loader2 className="w-5 h-5 animate-spin" /> : <Printer className="w-5 h-5" />}
                  Completar e Imprimir
                </button>
              </div>
            </div>
          </div>
        )}

        {/* VISTA: INVENTARIO */}
        {pestana === 'productos' && (
          <div className="space-y-6">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div onClick={() => setFiltroStock('todos')} className={`p-5 rounded-2xl border cursor-pointer transition-all duration-200 ${filtroStock === 'todos' ? 'bg-blue-950/50 border-blue-500 shadow-lg shadow-blue-950/50' : 'bg-slate-900/60 border-slate-800 hover:border-slate-700 hover:bg-slate-900'}`}>
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Total Productos</span>
                  <div className="p-2 bg-blue-500/10 text-blue-400 rounded-lg"><ShoppingBag className="w-4 h-4" /></div>
                </div>
                <p className="text-3xl font-extrabold mt-3 text-white">{productos.length}</p>
              </div>

              <div onClick={() => setFiltroStock('bajo')} className={`p-5 rounded-2xl border cursor-pointer transition-all duration-200 ${filtroStock === 'bajo' ? 'bg-amber-950/50 border-amber-500 shadow-lg shadow-amber-950/50' : 'bg-slate-900/60 border-slate-800 hover:border-slate-700 hover:bg-slate-900'}`}>
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-amber-400/90 uppercase tracking-wider">Bajo Stock (&le; 5)</span>
                  <div className="p-2 bg-amber-500/10 text-amber-400 rounded-lg"><AlertTriangle className="w-4 h-4" /></div>
                </div>
                <p className="text-3xl font-extrabold mt-3 text-amber-400">{productosBajoStock.length}</p>
              </div>

              <div onClick={() => setFiltroStock('agotado')} className={`p-5 rounded-2xl border cursor-pointer transition-all duration-200 ${filtroStock === 'agotado' ? 'bg-red-950/50 border-red-500 shadow-lg shadow-red-950/50' : 'bg-slate-900/60 border-slate-800 hover:border-slate-700 hover:bg-slate-900'}`}>
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-red-400/90 uppercase tracking-wider">Agotados</span>
                  <div className="p-2 bg-red-500/10 text-red-400 rounded-lg"><XCircle className="w-4 h-4" /></div>
                </div>
                <p className="text-3xl font-extrabold mt-3 text-red-400">{productosAgotados.length}</p>
              </div>

              <div onClick={() => setFiltroStock('vencidos')} className={`p-5 rounded-2xl border cursor-pointer transition-all duration-200 ${filtroStock === 'vencidos' ? 'bg-purple-950/50 border-purple-500 shadow-lg shadow-purple-950/50' : 'bg-slate-900/60 border-slate-800 hover:border-slate-700 hover:bg-slate-900'}`}>
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-purple-300 uppercase tracking-wider">Por Vencer / Vencidos</span>
                  <div className="p-2 bg-purple-500/10 text-purple-400 rounded-lg"><Clock className="w-4 h-4" /></div>
                </div>
                <p className="text-3xl font-extrabold mt-3 text-purple-300">{lotesVencidos.length + lotesPorVencer.length}</p>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row gap-3 items-center justify-between bg-slate-900/60 p-4 rounded-2xl border border-slate-800">
              <div className="relative w-full sm:w-1/2">
                <Search className="w-5 h-5 text-slate-400 absolute left-3.5 top-3" />
                <input
                  type="text"
                  placeholder="Buscar producto por nombre o escanea un código..."
                  value={busqueda}
                  onChange={(e) => setBusqueda(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-11 pr-4 py-2.5 text-sm text-white focus:outline-none focus:border-blue-500 transition"
                />
              </div>

              <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                <button
                  onClick={exportarInventarioExcel}
                  className="bg-emerald-950 hover:bg-emerald-900 text-emerald-400 border border-emerald-800 font-semibold text-xs px-3.5 py-2.5 rounded-xl flex items-center gap-1.5 transition"
                  title="Exportar catálogo a Excel"
                >
                  <FileSpreadsheet className="w-4 h-4" /> Exportar Excel
                </button>

                {filtroStock !== 'todos' && (
                  <button onClick={() => setFiltroStock('todos')} className="text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-2.5 rounded-xl flex items-center gap-1.5 transition">
                    <X className="w-3.5 h-3.5" /> Limpiar
                  </button>
                )}

                <button onClick={() => setModalProductoAbierto(true)} className="bg-blue-600 hover:bg-blue-500 text-white font-semibold text-sm px-4 py-2.5 rounded-xl flex items-center gap-2 transition shadow-lg shadow-blue-600/20">
                  <Plus className="w-4 h-4" /> Nuevo Producto
                </button>
              </div>
            </div>

            {cargandoProductos ? (
              <div className="p-8 text-center text-slate-400 animate-pulse">Cargando catálogo...</div>
            ) : productosFiltrados.length === 0 ? (
              <div className="p-12 text-center bg-slate-900/40 rounded-2xl border border-slate-800 text-slate-400">
                No se encontraron productos registrados.
              </div>
            ) : (
              <div className="grid grid-cols-1 gap-3">
                {productosFiltrados.map((p) => {
                  const prodLotes = lotes.filter((l) => l.producto_id === p.id)
                  const estaDesplegado = prodSeleccionadoLote?.id === p.id

                  return (
                    <div key={p.id} className="bg-slate-900/80 rounded-2xl border border-slate-800 overflow-hidden transition hover:border-slate-700">
                      <div className="p-4 flex flex-col md:flex-row md:items-center justify-between gap-4">
                        {editandoProdId === p.id ? (
                          <div className="grid grid-cols-1 sm:grid-cols-5 gap-2 w-full">
                            <input type="text" value={formProdEdit.nombre} onChange={(e) => setFormProdEdit({ ...formProdEdit, nombre: e.target.value })} className="bg-slate-950 border border-blue-500 rounded-lg px-2.5 py-1.5 text-xs text-white" />
                            <input type="text" value={formProdEdit.codigo_barra} onChange={(e) => setFormProdEdit({ ...formProdEdit, codigo_barra: e.target.value })} className="bg-slate-950 border border-blue-500 rounded-lg px-2.5 py-1.5 text-xs text-white" />
                            <select value={formProdEdit.categoria_id} onChange={(e) => setFormProdEdit({ ...formProdEdit, categoria_id: e.target.value })} className="bg-slate-950 border border-blue-500 rounded-lg px-2.5 py-1.5 text-xs text-white">
                              <option value="">Sin Categoría</option>
                              {categorias.map((c) => (<option key={c.id} value={c.id}>{c.nombre}</option>))}
                            </select>
                            <input type="number" step="0.01" value={formProdEdit.precio} onChange={(e) => setFormProdEdit({ ...formProdEdit, precio: e.target.value })} className="bg-slate-950 border border-blue-500 rounded-lg px-2.5 py-1.5 text-xs text-white" />
                            <div className="flex items-center gap-1 justify-end">
                              <button onClick={() => guardarEdicionProd(p.id)} className="p-1.5 text-emerald-400 hover:bg-slate-800 rounded-lg"><Check className="w-4 h-4" /></button>
                              <button onClick={() => setEditandoProdId(null)} className="p-1.5 text-slate-400 hover:bg-slate-800 rounded-lg"><X className="w-4 h-4" /></button>
                            </div>
                          </div>
                        ) : (
                          <>
                            <div className="flex-1 space-y-1">
                              <div className="flex items-center gap-2">
                                <span className="font-bold text-slate-100 text-base">{p.nombre}</span>
                                <span className="bg-slate-800 px-2.5 py-0.5 rounded-full text-xs text-slate-300 border border-slate-700">
                                  {p.categorias?.nombre || 'Sin Categoría'}
                                </span>
                                {p.tipo_venta === 'peso' ? (
                                  <span className="bg-purple-950/80 text-purple-300 px-2 py-0.5 rounded-full text-[10px] font-semibold border border-purple-800 flex items-center gap-1">
                                    <Scale className="w-3 h-3" /> Por Kg/Peso
                                  </span>
                                ) : (
                                  <span className="bg-blue-950/80 text-blue-300 px-2 py-0.5 rounded-full text-[10px] font-semibold border border-blue-800">
                                    Por Unidad
                                  </span>
                                )}
                              </div>
                              <div className="flex items-center gap-4 text-xs text-slate-400">
                                {p.codigo_barra ? (
                                  <span className="flex items-center gap-1"><Barcode className="w-3.5 h-3.5 text-blue-400" /> {p.codigo_barra}</span>
                                ) : (
                                  <span className="italic">Sin código</span>
                                )}
                                <span className="text-emerald-400 font-bold text-sm">
                                  ${Number(p.precio).toFixed(2)} {p.tipo_venta === 'peso' ? '/ Kg' : ''}
                                </span>
                              </div>
                            </div>

                            <div className="flex flex-wrap items-center gap-3 justify-between md:justify-end">
                              <TextoResumenVencimiento prodLotes={prodLotes} />

                              <div className="flex items-center gap-1.5 bg-slate-950/80 px-2.5 py-1 rounded-xl border border-slate-800">
                                <button onClick={() => modificarStockRapido(p.id, p.stock, -1)} className="text-slate-400 hover:text-red-400 transition">
                                  <MinusCircle className="w-4 h-4" />
                                </button>
                                <span className={`px-2 py-0.5 text-xs font-bold ${p.stock === 0 ? 'text-red-400' : p.stock <= 5 ? 'text-amber-400' : 'text-emerald-400'}`}>
                                  {p.stock} {p.tipo_venta === 'peso' ? 'Kg' : 'u.'}
                                </span>
                                <button onClick={() => modificarStockRapido(p.id, p.stock, 1)} className="text-slate-400 hover:text-emerald-400 transition">
                                  <PlusCircle className="w-4 h-4" />
                                </button>
                              </div>

                              <button onClick={() => setProdSeleccionadoLote(estaDesplegado ? null : p)} className="flex items-center gap-1.5 text-xs bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-1.5 rounded-xl transition">
                                <Calendar className="w-3.5 h-3.5 text-purple-400" /> Lotes ({prodLotes.length})
                                {estaDesplegado ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                              </button>

                              <div className="flex items-center gap-1">
                                <button onClick={() => iniciarEdicionProd(p)} className="p-1.5 text-blue-400 hover:bg-slate-800 rounded-lg transition"><Edit2 className="w-4 h-4" /></button>
                                <button onClick={() => eliminarProducto(p.id, p.nombre)} className="p-1.5 text-red-400 hover:bg-slate-800 rounded-lg transition"><Trash2 className="w-4 h-4" /></button>
                              </div>
                            </div>
                          </>
                        )}
                      </div>

                      {estaDesplegado && (
                        <div className="p-4 bg-slate-950/80 border-t border-slate-800 space-y-3">
                          <h4 className="text-xs font-bold text-purple-300 uppercase tracking-wider flex items-center gap-1.5">
                            <Clock className="w-4 h-4" /> Control de Lotes para "{p.nombre}"
                          </h4>

                          <form onSubmit={agregarLote} className="flex flex-wrap gap-2 items-center">
                            <input type="number" step="0.001" placeholder="Cant. (u o Kg)" value={nuevoLote.cantidad} onChange={(e) => setNuevoLote({ ...nuevoLote, cantidad: e.target.value })} className="bg-slate-900 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-white w-32 focus:outline-none" />
                            <input type="date" value={nuevoLote.fecha_vencimiento} onChange={(e) => setNuevoLote({ ...nuevoLote, fecha_vencimiento: e.target.value })} className="bg-slate-900 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-white focus:outline-none" required />
                            <button type="submit" disabled={guardandoLote || !nuevoLote.fecha_vencimiento} className="bg-purple-600 hover:bg-purple-500 text-white text-xs px-3.5 py-1.5 rounded-xl font-medium flex items-center gap-1 transition">
                              {guardandoLote ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Plus className="w-3.5 h-3.5" />} Lote
                            </button>
                          </form>

                          {prodLotes.length === 0 ? (
                            <p className="text-xs text-slate-500 italic">No hay vencimientos registrados.</p>
                          ) : (
                            <div className="space-y-1.5">
                              {prodLotes.map((lote) => {
                                const diasRestantes = calcularDiasRestantes(lote.fecha_vencimiento)
                                return (
                                  <div key={lote.id} className="flex items-center justify-between bg-slate-900 p-2.5 rounded-xl text-xs border border-slate-800">
                                    <div className="flex items-center gap-3">
                                      <span className="font-semibold text-slate-300">{lote.cantidad} {p.tipo_venta === 'peso' ? 'Kg' : 'u.'}</span>
                                      <span className="text-slate-400">Vence: {lote.fecha_vencimiento}</span>
                                      {diasRestantes < 0 ? (
                                        <span className="text-[10px] font-bold text-red-400 bg-red-950/80 px-2 py-0.5 rounded border border-red-800">Ya venció</span>
                                      ) : diasRestantes === 0 ? (
                                        <span className="text-[10px] font-bold text-white bg-red-600 px-2 py-0.5 rounded animate-pulse">Vence hoy</span>
                                      ) : (
                                        <span className="text-[10px] font-medium text-blue-300 bg-blue-950 px-2 py-0.5 rounded border border-blue-800">Faltan {diasRestantes} días</span>
                                      )}
                                    </div>
                                    <button onClick={() => eliminarLote(lote.id)} className="text-slate-500 hover:text-red-400 transition"><Trash2 className="w-3.5 h-3.5" /></button>
                                  </div>
                                )
                              })}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            )}
          </div>
        )}

        {/* VISTA: CLIENTES */}
        {pestana === 'clientes' && (
          <div className="bg-slate-900/80 p-6 rounded-2xl border border-slate-800 space-y-4">
            <div className="flex justify-between items-center">
              <h3 className="text-base font-bold text-slate-200 flex items-center gap-2">
                <Users className="w-5 h-5 text-amber-400" /> Registro de Clientes / RUT
              </h3>
              <button onClick={() => setModalClienteAbierto(true)} className="bg-amber-600 hover:bg-amber-500 text-white font-semibold text-sm px-4 py-2 rounded-xl flex items-center gap-2 transition">
                <UserPlus className="w-4 h-4" /> Nuevo Cliente
              </button>
            </div>

            {cargandoClientes ? (
              <p className="text-slate-400 animate-pulse text-xs">Cargando...</p>
            ) : clientes.length === 0 ? (
              <p className="text-xs text-slate-500 italic py-8 text-center">No hay clientes registrados.</p>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                {clientes.map((c) => (
                  <div key={c.id} className="bg-slate-950 p-4 rounded-xl border border-slate-800 space-y-1">
                    <p className="font-bold text-sm text-slate-100">{c.nombre}</p>
                    {c.rut_ci && <p className="text-xs text-amber-400 font-mono">RUT / C.I.: {c.rut_ci}</p>}
                    {c.telefono && <p className="text-xs text-slate-400">Tel: {c.telefono}</p>}
                    {c.direccion && <p className="text-xs text-slate-500">Dir: {c.direccion}</p>}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* VISTA: DASHBOARD */}
        {pestana === 'dashboard' && (
          <div className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="bg-slate-900/80 p-6 rounded-2xl border border-slate-800 flex items-center justify-between">
                <div>
                  <p className="text-xs text-slate-400 uppercase font-bold tracking-wider">Valor del Inventario</p>
                  <p className="text-3xl font-black text-emerald-400 mt-2">${valorTotalInventario.toFixed(2)}</p>
                </div>
                <div className="p-3 bg-emerald-500/10 rounded-2xl text-emerald-400"><TrendingUp className="w-8 h-8" /></div>
              </div>

              <div className="bg-slate-900/80 p-6 rounded-2xl border border-slate-800 flex items-center justify-between">
                <div>
                  <p className="text-xs text-slate-400 uppercase font-bold tracking-wider">Ingresos por Ventas</p>
                  <p className="text-3xl font-black text-purple-400 mt-2">${totalVentasMonto.toFixed(2)}</p>
                </div>
                <div className="p-3 bg-purple-500/10 rounded-2xl text-purple-400"><DollarSign className="w-8 h-8" /></div>
              </div>
            </div>

            <div className="bg-slate-900/80 p-6 rounded-2xl border border-slate-800">
              <h3 className="text-xs font-bold text-slate-300 uppercase tracking-wider mb-6">Ventas Recientes</h3>
              {datosGraficoVentas.length === 0 ? (
                <p className="text-xs text-slate-500 italic">Sin datos válidos para el gráfico.</p>
              ) : (
                <div className="h-64 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={datosGraficoVentas}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                      <XAxis dataKey="fecha" stroke="#64748b" fontSize={12} />
                      <YAxis stroke="#64748b" fontSize={12} />
                      <Tooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '12px', color: '#fff' }} />
                      <Bar dataKey="monto" fill="#10b981" radius={[6, 6, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              )}
            </div>

            <div className="bg-slate-900/80 p-6 rounded-2xl border border-slate-800 space-y-4">
              <div className="flex justify-between items-center">
                <h3 className="text-sm font-bold text-slate-200 uppercase tracking-wider flex items-center gap-2">
                  <Receipt className="w-4 h-4 text-purple-400" /> Historial de Transacciones
                </h3>

                <button
                  onClick={exportarVentasExcel}
                  className="bg-purple-950 hover:bg-purple-900 text-purple-300 border border-purple-800 font-semibold text-xs px-3 py-1.5 rounded-xl flex items-center gap-1.5 transition"
                >
                  <Download className="w-3.5 h-3.5" /> Exportar Ventas en Excel
                </button>
              </div>

              {ventas.length === 0 ? (
                <p className="text-xs text-slate-500 italic">No hay registros de ventas.</p>
              ) : (
                <div className="space-y-2">
                  {ventas.map((v) => (
                    <div key={v.id} className="flex flex-wrap items-center justify-between bg-slate-950 p-3.5 rounded-xl border border-slate-800 gap-3">
                      <div className="flex flex-wrap items-center gap-3">
                        <span className="text-xs font-bold text-slate-400">#Venta {v.id}</span>
                        <span className="text-xs text-slate-500">
                          {new Date(v.created_at).toLocaleString('es-UY', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' })}
                        </span>
                        <span className="text-xs bg-slate-900 border border-slate-800 text-slate-300 px-2 py-0.5 rounded font-medium">
                          {v.metodo_pago}
                        </span>
                        {v.clientes?.rut_ci && (
                          <span className="text-[10px] bg-amber-950 text-amber-300 px-2 py-0.5 rounded border border-amber-800 font-mono">
                            RUT: {v.clientes.rut_ci}
                          </span>
                        )}
                        {v.estado === 'anulada' ? (
                          <span className="bg-red-950 text-red-400 text-[10px] font-bold px-2 py-0.5 rounded border border-red-800">ANULADA</span>
                        ) : (
                          <span className="bg-emerald-950 text-emerald-400 text-[10px] font-bold px-2 py-0.5 rounded border border-emerald-800">COMPLETADA</span>
                        )}
                      </div>

                      <div className="flex items-center gap-3">
                        <span className={`font-black text-sm ${v.estado === 'anulada' ? 'line-through text-slate-500' : 'text-emerald-400'}`}>
                          ${Number(v.total).toFixed(2)}
                        </span>

                        {/* BOTÓN REIMPRIMIR TICKET */}
                        <button
                          onClick={() => reimprimirTicketVenta(v)}
                          title="Reimprimir Ticket de esta venta"
                          className="bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs px-2.5 py-1.5 rounded-xl font-medium border border-slate-700 flex items-center gap-1 transition"
                        >
                          <Printer className="w-3.5 h-3.5 text-emerald-400" /> Reimprimir
                        </button>

                        {v.estado !== 'anulada' && (
                          <button
                            onClick={() => anularVenta(v.id)}
                            disabled={procesandoAnulacion === v.id}
                            className="bg-red-950 hover:bg-red-900 text-red-300 text-xs px-2.5 py-1.5 rounded-xl font-medium border border-red-800 flex items-center gap-1 transition"
                          >
                            {procesandoAnulacion === v.id ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <RotateCcw className="w-3.5 h-3.5" />} Anular
                          </button>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* VISTA: CATEGORÍAS */}
        {pestana === 'categorias' && (
          <div className="bg-slate-900/80 p-6 rounded-2xl border border-slate-800 space-y-4">
            <div className="flex justify-between items-center">
              <h3 className="text-base font-bold text-slate-200">Categorías de Productos</h3>
              <button onClick={() => setModalCategoriaAbierto(true)} className="bg-blue-600 hover:bg-blue-500 text-white font-semibold text-sm px-4 py-2 rounded-xl flex items-center gap-2 transition">
                <Plus className="w-4 h-4" /> Nueva Categoría
              </button>
            </div>

            {cargandoCategorias ? (
              <p className="text-slate-400 animate-pulse text-xs">Cargando...</p>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                {categorias.map((cat) => (
                  <div key={cat.id} className="bg-slate-950 p-4 rounded-xl border border-slate-800 flex justify-between items-center">
                    {editandoCatId === cat.id ? (
                      <div className="flex items-center gap-2 w-full">
                        <input type="text" value={nombreCatEditado} onChange={(e) => setNombreCatEditado(e.target.value)} className="flex-1 bg-slate-900 border border-blue-500 text-white rounded-lg px-3 py-1 text-sm" autoFocus />
                        <button onClick={() => guardarEdicionCat(cat.id)} className="p-1 text-emerald-400 hover:bg-slate-800 rounded"><Check className="w-4 h-4" /></button>
                        <button onClick={() => setEditandoCatId(null)} className="p-1 text-slate-400 hover:bg-slate-800 rounded"><X className="w-4 h-4" /></button>
                      </div>
                    ) : (
                      <>
                        <span className="font-semibold text-slate-200 text-sm">{cat.nombre}</span>
                        <div className="flex items-center gap-1">
                          <button onClick={() => { setEditandoCatId(cat.id); setNombreCatEditado(cat.nombre); }} className="p-1.5 text-blue-400 hover:bg-slate-800 rounded-lg"><Edit2 className="w-4 h-4" /></button>
                          <button onClick={() => eliminarCategoria(cat.id, cat.nombre)} className="p-1.5 text-red-400 hover:bg-slate-800 rounded-lg"><Trash2 className="w-4 h-4" /></button>
                        </div>
                      </>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

      </main>

      {/* MODAL: IMPRESIÓN DE TICKET TÉRMICO */}
      {modalTicketAbierto && ultimaVentaRealizada && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl w-full max-w-sm shadow-2xl space-y-4">
            
            {/* FORMATO TICKET TÉRMICO (72mm EPSON) */}
            <div id="ticket-impresion" className="bg-white text-black p-4 rounded-lg font-mono text-xs space-y-2">
              <div className="text-center border-b border-dashed border-gray-400 pb-2">
                <h2 className="font-extrabold text-sm uppercase tracking-wide">StockPulse POS</h2>
                <p className="text-[10px]">Comercio & Fiambrería</p>
                <p className="text-[10px] mt-1">Ticket N° #{ultimaVentaRealizada.id}</p>
                <p className="text-[10px]">{new Date(ultimaVentaRealizada.created_at).toLocaleString('es-UY')}</p>
              </div>

              {ultimaVentaRealizada.clientes && (
                <div className="border-b border-dashed border-gray-400 pb-2 text-[10px]">
                  <p className="font-bold">Cliente: {ultimaVentaRealizada.clientes.nombre}</p>
                  {ultimaVentaRealizada.clientes.rut_ci && <p>RUT/C.I.: {ultimaVentaRealizada.clientes.rut_ci}</p>}
                </div>
              )}

              <div className="space-y-1.5 border-b border-dashed border-gray-400 pb-2 pt-1">
                {ultimaVentaRealizada.items.map((item, idx) => (
                  <div key={idx} className="flex justify-between items-start">
                    <div>
                      <p className="font-bold text-[11px]">{item.nombre}</p>
                      <p className="text-[10px] text-gray-600">
                        {item.tipo_venta === 'peso' ? `${item.gramosOriginales || Math.round(item.cantidadEnCarrito * 1000)}g (${item.cantidadEnCarrito} Kg)` : `${item.cantidadEnCarrito} u.`} &times; ${item.precio}
                      </p>
                    </div>
                    <span className="font-bold text-[11px]">${(item.subtotal || item.precio * item.cantidadEnCarrito).toFixed(2)}</span>
                  </div>
                ))}
              </div>

              <div className="space-y-1 text-right pt-1">
                <div className="flex justify-between font-extrabold text-xs">
                  <span>TOTAL COBRADO:</span>
                  <span>${Number(ultimaVentaRealizada.total).toFixed(2)}</span>
                </div>
                <p className="text-[10px] text-gray-600">Forma de Pago: {ultimaVentaRealizada.metodo_pago}</p>
              </div>

              <div className="text-center pt-3 border-t border-dashed border-gray-300 text-[10px] text-gray-500">
                <p>*** ¡Muchas gracias por su compra! ***</p>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-800 no-imprimir">
              <button onClick={() => setModalTicketAbierto(false)} className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-semibold">
                Cerrar
              </button>
              <button onClick={() => window.print()} className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5">
                <Printer className="w-4 h-4" /> Imprimir Ticket
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODALES ADICIONALES (CLIENTE, PESO, PRODUCTO, CATEGORÍA) */}
      {modalClienteAbierto && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl w-full max-w-md shadow-2xl space-y-4">
            <div className="flex justify-between items-center border-b border-slate-800 pb-3">
              <h3 className="text-lg font-bold text-white">Nuevo Cliente</h3>
              <button onClick={() => setModalClienteAbierto(false)} className="text-slate-400 hover:text-white"><X className="w-5 h-5" /></button>
            </div>

            <form onSubmit={agregarCliente} className="space-y-3">
              <div>
                <label className="text-xs text-slate-400 font-semibold mb-1 block">Nombre / Razón Social</label>
                <input type="text" value={formCliente.nombre} onChange={(e) => setFormCliente({ ...formCliente, nombre: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-amber-500" required />
              </div>
              <div>
                <label className="text-xs text-slate-400 font-semibold mb-1 block">RUT / C.I.</label>
                <input type="text" value={formCliente.rut_ci} onChange={(e) => setFormCliente({ ...formCliente, rut_ci: e.target.value })} placeholder="Ej: 211234560018" className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-amber-500 font-mono" />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-xs text-slate-400 font-semibold mb-1 block">Teléfono</label>
                  <input type="text" value={formCliente.telefono} onChange={(e) => setFormCliente({ ...formCliente, telefono: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-amber-500" />
                </div>
                <div>
                  <label className="text-xs text-slate-400 font-semibold mb-1 block">Dirección</label>
                  <input type="text" value={formCliente.direccion} onChange={(e) => setFormCliente({ ...formCliente, direccion: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-amber-500" />
                </div>
              </div>

              <div className="pt-3 flex justify-end gap-2">
                <button type="button" onClick={() => setModalClienteAbierto(false)} className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-semibold">Cancelar</button>
                <button type="submit" disabled={guardandoCliente} className="px-4 py-2 bg-amber-600 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5">
                  {guardandoCliente ? <Loader2 className="w-4 h-4 animate-spin" /> : <Plus className="w-4 h-4" />} Guardar Cliente
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {modalPesoAbierto && productoParaPeso && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl w-full max-w-md shadow-2xl space-y-4">
            <div className="flex justify-between items-center border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <Scale className="w-5 h-5 text-purple-400" />
                <h3 className="text-base font-bold text-white">Venta por Peso: {productoParaPeso.nombre}</h3>
              </div>
              <button onClick={() => setModalPesoAbierto(false)} className="text-slate-400 hover:text-white"><X className="w-5 h-5" /></button>
            </div>

            <p className="text-xs text-slate-400">
              Precio base: <span className="text-emerald-400 font-bold">${productoParaPeso.precio} / Kg</span>
            </p>

            <form onSubmit={agregarPesoAlCarrito} className="space-y-4">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold mb-1 block">Cantidad en Gramos</label>
                  <input type="number" placeholder="Ej: 250" value={inputGramos} onChange={(e) => manejarCambioGramos(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2.5 text-sm text-white focus:outline-none focus:border-purple-500 font-bold" autoFocus />
                </div>
                <div>
                  <label className="text-xs text-slate-400 font-semibold mb-1 block">O Monto en $</label>
                  <input type="number" step="0.01" placeholder="Ej: 150" value={inputMontoDinero} onChange={(e) => manejarCambioMontoDinero(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2.5 text-sm text-white focus:outline-none focus:border-emerald-500 font-bold text-emerald-400" />
                </div>
              </div>

              {inputGramos > 0 && (
                <div className="bg-slate-950 p-3 rounded-xl border border-slate-800 text-xs space-y-1">
                  <div className="flex justify-between text-slate-400">
                    <span>Peso equivalente:</span>
                    <span className="font-bold text-slate-200">{(parseFloat(inputGramos) / 1000).toFixed(3)} Kg</span>
                  </div>
                  <div className="flex justify-between text-slate-400">
                    <span>Subtotal a cobrar:</span>
                    <span className="font-extrabold text-emerald-400 text-sm">${((parseFloat(inputGramos) / 1000) * productoParaPeso.precio).toFixed(2)}</span>
                  </div>
                </div>
              )}

              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setModalPesoAbierto(false)} className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-semibold">Cancelar</button>
                <button type="submit" className="px-4 py-2 bg-purple-600 hover:bg-purple-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5"><Plus className="w-4 h-4" /> Agregar al Carrito</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {modalProductoAbierto && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl w-full max-w-lg shadow-2xl space-y-4">
            <div className="flex justify-between items-center border-b border-slate-800 pb-3">
              <h3 className="text-lg font-bold text-white">Nuevo Producto</h3>
              <button onClick={() => setModalProductoAbierto(false)} className="text-slate-400 hover:text-white"><X className="w-5 h-5" /></button>
            </div>

            <form onSubmit={agregarProducto} className="space-y-3">
              <div>
                <label className="text-xs text-slate-400 font-semibold mb-1 block">Nombre</label>
                <input type="text" value={formProd.nombre} onChange={(e) => setFormProd({ ...formProd, nombre: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-blue-500" required />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold mb-1 block">Cód. Barras</label>
                  <input type="text" value={formProd.codigo_barra} onChange={(e) => setFormProd({ ...formProd, codigo_barra: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-xs text-slate-400 font-semibold mb-1 block">Tipo de Venta</label>
                  <select value={formProd.tipo_venta} onChange={(e) => setFormProd({ ...formProd, tipo_venta: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-blue-500 font-bold text-purple-300">
                    <option value="unidad">Por Unidad</option>
                    <option value="peso">Por Peso (Kg/Gramos)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold mb-1 block">{formProd.tipo_venta === 'peso' ? 'Precio por Kg ($)' : 'Precio Unitario ($)'}</label>
                  <input type="number" step="0.01" value={formProd.precio} onChange={(e) => setFormProd({ ...formProd, precio: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-xs text-slate-400 font-semibold mb-1 block">{formProd.tipo_venta === 'peso' ? 'Stock Inicial (Kg)' : 'Stock Inicial (u)'}</label>
                  <input type="number" step="0.001" value={formProd.stock} onChange={(e) => setFormProd({ ...formProd, stock: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-blue-500" />
                </div>
              </div>

              <div>
                <label className="text-xs text-slate-400 font-semibold mb-1 block">Categoría</label>
                <select value={formProd.categoria_id} onChange={(e) => setFormProd({ ...formProd, categoria_id: e.target.value })} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-blue-500">
                  <option value="">Sin Categoría</option>
                  {categorias.map((c) => (<option key={c.id} value={c.id}>{c.nombre}</option>))}
                </select>
              </div>

              <div className="pt-3 flex justify-end gap-2">
                <button type="button" onClick={() => setModalProductoAbierto(false)} className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-semibold">Cancelar</button>
                <button type="submit" disabled={guardandoProd} className="px-4 py-2 bg-blue-600 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5">
                  {guardandoProd ? <Loader2 className="w-4 h-4 animate-spin" /> : <Plus className="w-4 h-4" />} Guardar
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {modalCategoriaAbierto && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl w-full max-w-md shadow-2xl space-y-4">
            <div className="flex justify-between items-center border-b border-slate-800 pb-3">
              <h3 className="text-lg font-bold text-white">Nueva Categoría</h3>
              <button onClick={() => setModalCategoriaAbierto(false)} className="text-slate-400 hover:text-white"><X className="w-5 h-5" /></button>
            </div>

            <form onSubmit={agregarCategoria} className="space-y-4">
              <div>
                <label className="text-xs text-slate-400 font-semibold mb-1 block">Nombre de Categoría</label>
                <input type="text" value={nuevoNombreCat} onChange={(e) => setNuevoNombreCat(e.target.value)} placeholder="Ej: Fiambrería, Lácteos..." className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-sm text-white focus:outline-none focus:border-blue-500" required />
              </div>

              <div className="flex justify-end gap-2">
                <button type="button" onClick={() => setModalCategoriaAbierto(false)} className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-semibold">Cancelar</button>
                <button type="submit" disabled={guardandoCat} className="px-4 py-2 bg-blue-600 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5">
                  {guardandoCat ? <Loader2 className="w-4 h-4 animate-spin" /> : <Plus className="w-4 h-4" />} Guardar
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

    </div>
  )
}

export default App