# 📦 StockPulse — POS & Inventory Management System

[![Vercel Deployment](https://img.shields.io/badge/Vercel-000000?style=for-the-badge&logo=vercel&logoColor=white)](https://stockpulse-xi-nine.vercel.app/)
[![React](https://img.shields.io/badge/React-61DAFB?style=for-the-badge&logo=react&logoColor=black)](https://react.dev/)
[![Tailwind CSS](https://img.shields.io/badge/Tailwind_CSS-38B2AC?style=for-the-badge&logo=tailwind-css&logoColor=white)](https://tailwindcss.com/)
[![Supabase](https://img.shields.io/badge/Supabase-3ECF8E?style=for-the-badge&logo=supabase&logoColor=white)](https://supabase.com/)

[Lee esto en Español](README.md)

StockPulse is a Full-Stack Web Point of Sale (POS) and inventory control application designed to streamline stock tracking, speed up checkout operations, and provide real-time sales metrics for retail businesses.

🚀 Live Demo: https://stockpulse-xi-nine.vercel.app/

---

## ✨ Key Features

- 🛒 Dynamic Point of Sale: Real-time sales processing with fast barcode scanning and product search.
- ⚖️ Sales by Weight & Unit: Flexible management for bulk items (by Kg/g) or single units.
- 💳 Multiple Payment Methods: Record transactions using Cash, Debit, Credit, and Bank Transfer.
- 🧾 Receipt Printing: Generation and printing of thermal receipts optimized for ESC/POS devices.
- 🧮 Cash Register & Shift Control: Opening, closing, and reconciliation of cash registers with full audit logs.
- 📊 Analytics & Metrics: Interactive charts displaying daily revenue, sales trends, and top-performing products.
- 📥 Data Export: Export reports and product catalogs directly to Excel (.xlsx).
- 🆔 Customer Management (RUT/C.I.): Quick identification of recurring customers for personalized receipts.

---

## 🛠️ Tech Stack

### Frontend
- React.js — Component-based UI library.
- Tailwind CSS — Utility-first CSS framework with responsive design and dark mode.
- Lucide React — Modern icon set.
- Recharts — Data visualization and analytical charts.
- XLSX (SheetJS) — Excel spreadsheet processing and export.

### Backend & Database
- Supabase — Backend-as-a-Service (BaaS) platform.
- PostgreSQL — Relational database for inventory, sales, and transaction persistence.

### Deployment & DevOps
- Git & GitHub — Version control.
- Vercel — Continuous integration and deployment (CI/CD).

---

## 🚀 Local Installation & Setup

To run this project locally, follow these steps:

1. Clone the repository:
git clone https://github.com/Ttonay99/stockpulse.git
cd stockpulse

2. Install dependencies:
npm install

3. Environment Variables Setup:
Create a .env file in the root directory with your Supabase credentials:
VITE_SUPABASE_URL=your_supabase_project_url
VITE_SUPABASE_ANON_KEY=your_supabase_anon_key

4. Run the development server:
npm run dev

---

## 📄 License

Developed by Antony (Ttonay99) as part of a professional Full-Stack portfolio. Free for educational and reference purposes.
