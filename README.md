📦 StockPulse — POS & Sistema de Control de Inventario
StockPulse es una aplicación Web Full-Stack de Punto de Venta (POS) y gestión de inventarios diseñada para optimizar el control de stock, acelerar el proceso de facturación/caja y brindar métricas clave de ventas en tiempo real para comercios y minoristas.

🚀 Demo en vivo: stockpulse-xi-nine.vercel.app

✨ Características Principales
🛒 Punto de Venta Dinámico: Procesamiento de ventas en tiempo real con soporte para búsqueda rápida por código o nombre.

⚖️ Ventas por Peso y Unidad: Gestión adaptable de productos fraccionados (por kilo/gramo) o por unidades individuales.

💳 Múltiples Métodos de Pago: Registro de transacciones con Pago Efectivo, Débito, Crédito y Transferencia.

🧾 Emisión e Impresión de Ticket: Generación e impresión de comprobantes de compra optimizados para impresoras térmicas ESC/POS.

🧮 Control de Caja y Turnos: Apertura, arqueo y cierre de caja registradora con historial detallado de movimientos.

📊 Métricas y Analíticas: Gráficos e indicadores interactivos de ventas diarias, ingresos y productos más vendidos.

📥 Exportación de Datos: Exportación de informes y catálogos en formato Excel (.xlsx).

🆔 Gestión de Clientes (RUT/C.I.): Identificación rápida de clientes recurrentes para facturación personalizada.

🛠️ Tecnologías Utilizadas
Frontend
React.js — Biblioteca UI basada en componentes.

Tailwind CSS — Framework CSS de utilidad para diseño responsivo y modo oscuro.

Lucide React — Set de iconografía profesional.

Recharts — Visualización de datos y gráficos analíticos.

XLSX (SheetJS) — Procesamiento y exportación de archivos Excel.

Backend & Base de Datos
Supabase — Plataforma Backend-as-a-Service (BaaS).

PostgreSQL — Base de datos relacional para persistencia de inventario, ventas y transacciones.

Despliegue & DevOps
Git & GitHub — Control de versiones.

Vercel — Despliegue e integración continua (CI/CD).

🚀 Instalación y Configuración Local
Si deseas ejecutar este proyecto en tu máquina local, sigue estos pasos:

1:  Clona el repositorio:
    git clone https://github.com/Ttonay99/stockpulse.git
    cd stockpulse

2:  Instala las dependencias:
    npm install

3:  Configura las Variables de Entorno:
    Crea un archivo .env en la raíz del proyecto con tus credenciales de Supabase:
    VITE_SUPABASE_URL=tu_supabase_project_url
    VITE_SUPABASE_ANON_KEY=tu_supabase_anon_key

4:  Inicia el servidor de desarrollo:
    npm run dev

📄 Licencia
Este proyecto fue desarrollado por Antony (Ttonay99) como parte de su portafolio profesional Full-Stack. Libre para uso educativo y de referencia.
