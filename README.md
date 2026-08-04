<div align="center">

# 🏭 Program WZ — Business Management Software

**A WPF desktop application for tracking outbound shipments and returns between a manufacturer and subcontractors**

[![C#](https://img.shields.io/badge/C%23-.NET_4.7.2-512BD4?style=flat-square&logo=csharp)](https://dotnet.microsoft.com/)
[![WPF](https://img.shields.io/badge/UI-WPF-0078D6?style=flat-square&logo=windows)](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)
[![SQL Server](https://img.shields.io/badge/DB-SQL_Server-CC2927?style=flat-square&logo=microsoftsqlserver)](https://www.microsoft.com/sql-server)
[![Excel](https://img.shields.io/badge/Export-Excel-217346?style=flat-square&logo=microsoftexcel)](https://www.microsoft.com/excel)

<br/>

> Tracks which parts were sent to which subcontractor, how many were returned,
> how many are missing, and generates **WZ documents** (Wydanie Zewnętrzne —
> the Polish standard external release form) with Excel export and PDF preview.

</div>

---

## 📋 Table of Contents

- [What Is a WZ Document](#-what-is-a-wz-document)
- [Features](#-features)
- [Screenshots](#-screenshots)
- [How It Works](#-how-it-works)
- [Database Schema](#-database-schema)
- [Tech Stack](#-tech-stack)
- [Requirements](#-requirements)
- [Getting Started](#-getting-started)
- [Project Structure](#-project-structure)

---

## 📄 What Is a WZ Document

**WZ** (*Wydanie Zewnętrzne* — External Release) is a standard Polish logistics document
that confirms goods have been transferred from one company to another.

In this context, the manufacturer sends production parts (details) to an external
subcontractor for processing (e.g. laser cutting, painting, machining),
and the WZ document records what was sent, when, and in what quantity.
When the parts return, the system records how many came back and how many are missing.

---

## ✨ Features

### 📦 Order Management
- Register outbound shipments: part → company, date, quantity
- Record returns: how many came back, when, and how many are defective/missing
- Partial returns supported via `InnerOrders` — receive part of an order in multiple batches
- Mark orders as **Finished** manually or automatically when all parts return
- Toggle between **active orders** and **finished orders** view

### 🏢 Company Registry
- Add and edit subcontractor companies (name, short name, address, NIP tax ID)
- Assign any order to a registered company

### 🔩 Item (Part) Catalog
- Register production parts by code and surface area (dm²)
- Edit part data; attach an image code reference

### 📊 Report
- Summary report per part: total quantity received, total surface area processed
- Filter by date range

### 📤 Excel Export
- Fill a pre-defined Excel template with current order data
- Opens automatically in Excel after export via `Microsoft.Office.Interop.Excel`

### 🖨️ WZ Document Print
- Generate and preview the official WZ release document (PDF)
- PDF rendered to PNG for in-app preview (`SautinSoft.PdfFocus`)
- F3 keyboard shortcut opens the preview panel

### ⚡ Async UI
- All database operations run in `Task.Run()` — the UI never freezes

---

## 🖼️ Screenshots

> All screenshots are in the `Program WZ Documentation/` folder.

| Screen | Description |
|---|---|
| `Główne okno.png` | Main window — active orders data grid |
| `Dodaj firmę.png` | Add company dialog |
| `Dodaj przedmiot.png` | Add item (part) dialog |
| `Dodaj zamówienie.png` | Add order dialog |
| `Zmień dane firmę.png` | Edit company dialog |
| `Zmiana danych przedmiotu.png` | Edit item dialog |
| `Podgląd na F3.png` | F3 PDF preview panel |
| `Raport.png` | Summary report window |
| `Wydruk WZ_1.png` | WZ document printout — page 1 |
| `Wydruk WZ_2.png` | WZ document printout — page 2 |

---

## 🔬 How It Works

```
Manufacturer sends parts to subcontractor
               │
               ▼
         AddOrder dialog
         ┌───────────────────────────────┐
         │ Part code (ItemCode)          │
         │ Company (ShortName)           │
         │ Date sent (DateOut)           │
         │ Quantity sent (CountOut)      │
         │ Notes (Description)           │
         └──────────────┬────────────────┘
                        │ DBWork.AddInOrders()
                        ▼
              ItemOrders row created
              CountIn = 0, Finished = false
                        │
                        ▼
         Main grid shows active order
         [Print] [Part] [dm²] [DateOut] [Sent] [DateIn] [Received] [Missing] [Area] [Done] [Company] [Notes]
                        │
         Parts come back from subcontractor
                        │
                        ▼
         User fills DateIn + CountIn in main grid
         → Save button → DBWork.UpdateIfTakeDetail()
               │
               ├── Validates: CountIn ≤ CountOut - already received
               ├── Validates: DateIn ≥ DateOut
               ├── Creates InnerOrder record (partial delivery log)
               ├── Updates: CountIn += batch, CountFails = CountOut - CountIn
               ├── Updates: TotalDm = CountIn × Item.Size
               └── If CountIn == CountOut → Finished = true (auto-close)
                        │
                        ▼
               Order moves to "Finished" view
               Available in Report and Excel export
```

---

## 🗄️ Database Schema

**Database name:** `Firma`  
**ORM:** LINQ to SQL (`Classes.dbml`)

### `Item` — Parts catalog

| Column | Type | Description |
|---|---|---|
| `Id` | int PK | Auto-increment |
| `ItemCode` | string | Part code / name |
| `Size` | double | Surface area in dm² |
| `Image` | string | Image code reference |

### `ItemOrders` — Shipment orders

| Column | Type | Description |
|---|---|---|
| `Id` | int PK | Auto-increment |
| `ItemId` | FK → Item | Which part |
| `CompanyId` | FK → Company | Which subcontractor |
| `DateOut` | DateTime | Date sent |
| `CountOut` | int | Quantity sent |
| `DateIn` | DateTime? | Date returned (nullable) |
| `CountIn` | int | Quantity received so far |
| `CountFails` | int | Missing = CountOut − CountIn |
| `TotalDm` | double | Total surface = CountIn × Item.Size |
| `Finished` | bool | Order fully completed |
| `Description` | string | Notes |
| `PrintList` | bool | Include in print |

### `InnerOrders` — Partial delivery log

| Column | Description |
|---|---|
| `OrderId` | FK → ItemOrders |
| `ItemId` | FK → Item |
| `CountIn` | Quantity in this batch |
| `DateIn` | Date of this batch |

### `Companies` — Subcontractor registry

| Column | Description |
|---|---|
| `Name` | Full company name |
| `ShortName` | Short identifier (used in dropdowns) |
| `Addres` | Address |
| `NIP` | Polish tax identification number |

---

## 🛠 Tech Stack

| Technology | Purpose |
|---|---|
| C# / WPF (.NET 4.7.2) | Desktop UI framework |
| LINQ to SQL (`Classes.dbml`) | ORM — auto-generated from DB schema |
| Dapper 2.0 + Dapper.Contrib | Lightweight SQL queries |
| SQL Server / LocalDB | Database (`Firma` catalog) |
| Microsoft.Office.Interop.Excel | Excel template export |
| SautinSoft.PdfFocus | PDF → PNG conversion for in-app WZ preview |
| MD5Encryptor | MD5 utility class |
| `Task.Run()` + `Dispatcher.Invoke()` | Async DB ops without UI freeze |

---

## ⚙️ Requirements

- Windows 10 / 11
- .NET Framework 4.7.2
- Visual Studio 2019+
- SQL Server or SQL Server LocalDB
- **Microsoft Excel** installed (required for Excel export via COM Interop)

---

## 🚀 Getting Started

### 1 — Create the database

Connect to SQL Server (or LocalDB) and create a database named `Firma`.  
Run migrations or create the tables manually based on `Classes.dbml`.

### 2 — Configure the connection string

Open `App.config` and update the connection strings:

```xml
<connectionStrings>
  <!-- Remote SQL Server (production) -->
  <add name="MyConnection"
       connectionString="Server=YOUR_SERVER,1433;Database=Firma;User Id=user;Password=yourpassword;"
       providerName="System.Data.SqlClient" />

  <!-- LocalDB (development) -->
  <add name="MyConnection1"
       connectionString="Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Firma;Integrated Security=SSPI;" />
</connectionStrings>
```

The active connection is set in `MainWindow.xaml.cs`:
```csharp
static string connectionString = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Firma;Integrated Security=SSPI;";
```

Change this line to switch between local and remote.

### 3 — Build and run

Open `program.sln` in Visual Studio → Build → Run (`F5`).

---

## 📁 Project Structure

```
BUSINESS_MANAGEMENT_SOFTWARE/
├── Program WZ Documentation/       # UI screenshots (10 screens)
│
└── program/
    ├── MainWindow.xaml / .cs       # Main window — orders data grid, toolbar,
    │                               # receive parts, F3 preview, async refresh
    ├── AddInDB.xaml / .cs          # Add new item (part) to catalog
    ├── AddOrder.xaml / .cs         # Add new outbound order
    ├── AddCompany.xaml / .cs       # Register new subcontractor company
    ├── EditCompany.xaml / .cs      # Edit existing company details
    ├── Edit Item.xaml / .cs        # Edit item data, view order history
    ├── Raport.xaml / .cs           # Summary report — quantity + surface by part
    │
    ├── Classes/
    │   ├── DBWork.cs               # All database operations:
    │   │                           #   ShowAllDataInItemOrders()
    │   │                           #   AddInOrders()
    │   │                           #   UpdateIfTakeDetail()
    │   │                           #   AddCompanyInDB()
    │   │                           #   UpdateInfoInGrid()
    │   │                           #   DeleteAllOrdersInDB()
    │   ├── _ExelDoc.cs             # Excel export via Office COM Interop
    │   ├── _Raport.cs              # Report row model (part, count, surface)
    │   ├── MD5Encryptor.cs         # MD5 hash utility
    │   └── PdfToPng.cs             # PDF → PNG via SautinSoft.PdfFocus
    │
    ├── Classes.dbml                # LINQ to SQL schema designer
    ├── Classes.designer.cs         # Auto-generated ORM (Item, ItemOrder,
    │                               # InnerOrder, Company entities)
    └── App.config                  # Connection strings
```

---

<div align="center">

Made with C# · WPF · LINQ to SQL · SQL Server · Microsoft Excel Interop

</div>
