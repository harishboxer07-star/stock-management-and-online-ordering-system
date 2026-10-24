import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

PROJECT_ROOT = r"c:\Users\DHASWEEN\OneDrive\Desktop\smart-grocery-ai-production project"
OUTPUT_FILE = os.path.join(PROJECT_ROOT, "SmartGrocery_AI_Source_Code.docx")

# Color palette
NAVY_PRIMARY = RGBColor(30, 58, 138)      # #1E3A8A
BLUE_ACCENT = RGBColor(37, 99, 235)       # #2563EB
DARK_TEXT = RGBColor(15, 23, 42)          # #0F172A
SLATE_GRAY = RGBColor(100, 116, 139)      # #64748B
LIGHT_SLATE = RGBColor(241, 245, 249)     # #F1F5F9
WHITE = RGBColor(255, 255, 255)
GREEN_ACCENT = RGBColor(16, 185, 129)     # #10B981

LIGHT_BG_HEX = "F8FAFC"
HEADER_BG_HEX = "0F172A"
SUBHEADER_BG_HEX = "1E293B"
META_BG_HEX = "F1F5F9"
BORDER_HEX = "CBD5E1"
BLUE_BORDER_HEX = "3B82F6"
ACCENT_BLUE_HEX = "2563EB"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_borders(cell, top="CBD5E1", bottom="CBD5E1", left="CBD5E1", right="CBD5E1", left_sz="4", other_sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="{other_sz}" w:space="0" w:color="{top}"/>
            <w:left w:val="single" w:sz="{left_sz}" w:space="0" w:color="{left}"/>
            <w:bottom w:val="single" w:sz="{other_sz}" w:space="0" w:color="{bottom}"/>
            <w:right w:val="single" w:sz="{other_sz}" w:space="0" w:color="{right}"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def add_field(p, field_name):
    run = p.add_run()
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), field_name)
    run._r.append(fld)

def sanitize_file_content(path, content):
    """Mask real secrets in env files for security in documentation."""
    if path.endswith('.env'):
        lines = []
        for line in content.splitlines():
            if 'MONGODB_URI=mongodb+srv://' in line:
                lines.append('# MONGODB_URI=mongodb+srv://<USERNAME>:<PASSWORD>@cluster0.mongodb.net/smart_grocery_db?retryWrites=true&w=majority')
            elif line.startswith('JWT_SECRET='):
                lines.append('JWT_SECRET=smart_grocery_super_secret_jwt_key_2026')
            else:
                lines.append(line)
        return '\n'.join(lines)
    return content

SECTIONS = [
    {
        "title": "Section 1: Monorepo & Deployment Infrastructure",
        "description": "Root project configuration, container orchestration, Docker multi-stage builds, cloud platform deployment descriptors, and automated platform startup scripts.",
        "files": [
            {
                "path": "README.md",
                "lang": "Markdown",
                "role": "Project Technical Documentation",
                "desc": "Complete system overview, architecture breakdown, quickstart commands, persona credentials table, and deployment instructions for Render, Vercel, and Docker."
            },
            {
                "path": "package.json",
                "lang": "JSON",
                "role": "Monorepo Build Manifest",
                "desc": "Top-level npm package manifest orchestrating unified backend and frontend installation, concurrent development scripts, and production build pipelines."
            },
            {
                "path": "render.yaml",
                "lang": "YAML",
                "role": "Render.com Cloud Blueprint",
                "desc": "Infrastructure-as-Code (IaC) configuration for building and hosting the unified Express API, Socket.IO WebSockets, and static frontend on Render."
            },
            {
                "path": "docker-compose.yml",
                "lang": "YAML",
                "role": "Docker Multi-Container Orchestration",
                "desc": "Docker Compose manifest defining interconnected services: SmartGrocery Node.js fullstack container, MongoDB database instance, and network volumes."
            },
            {
                "path": "Dockerfile",
                "lang": "Dockerfile",
                "role": "Unified Multi-Stage Dockerfile",
                "desc": "Two-stage container build process: Stage 1 compiles the React 19 / Vite client; Stage 2 bundles Node.js backend dependencies and serves the unified application."
            },
            {
                "path": "start.bat",
                "lang": "Batch Script",
                "role": "Windows Automated Launcher",
                "desc": "One-click Windows execution script that detects dependencies, builds frontend static assets if missing, and initializes the local production server."
            },
            {
                "path": "start.sh",
                "lang": "Bash Script",
                "role": "Linux / macOS Automated Launcher",
                "desc": "Automated Unix shell script for environment validation, dependency installation, Vite compilation, and node server process initialization."
            },
            {
                "path": ".gitignore",
                "lang": "Git Configuration",
                "role": "Version Control Exclusions",
                "desc": "Rules excluding node_modules, build artifacts, environment secrets, and IDE caches from Git version control."
            },
            {
                "path": ".dockerignore",
                "lang": "Docker Configuration",
                "role": "Docker Build Exclusions",
                "desc": "Build context optimization rules preventing local node_modules, logs, and Git files from inflating container images."
            },
            {
                "path": ".node-version",
                "lang": "Text / Config",
                "role": "Runtime Engine Pinning",
                "desc": "Declares the required LTS Node.js engine version for automated cloud builders."
            }
        ]
    },
    {
        "title": "Section 2: Backend Server, Configuration & Seed Data",
        "description": "Backend Express server initialization, Socket.IO WebSocket real-time engine, MongoDB Mongoose connectivity, and database bootstrapping with mock catalog and personas.",
        "files": [
            {
                "path": "backend/package.json",
                "lang": "JSON",
                "role": "Backend Dependencies Manifest",
                "desc": "Declares backend libraries including Express, Mongoose, Socket.IO, PDFKit, JWT, bcryptjs, and Google Gemini AI SDK."
            },
            {
                "path": "backend/.env.example",
                "lang": "Env Configuration",
                "role": "Environment Variables Blueprint",
                "desc": "Template for essential backend environment variables including PORT, MONGODB_URI, JWT_SECRET, and GEMINI_API_KEY."
            },
            {
                "path": "backend/Dockerfile",
                "lang": "Dockerfile",
                "role": "Backend Containerfile",
                "desc": "Standalone Alpine Node container configuration for isolated backend microservice deployment."
            },
            {
                "path": "backend/.dockerignore",
                "lang": "Docker Configuration",
                "role": "Backend Docker Exclusions",
                "desc": "Excludes backend node_modules and local development artifacts from Docker builds."
            },
            {
                "path": "backend/src/server.js",
                "lang": "JavaScript (Node.js)",
                "role": "Main Server Entry & WebSockets",
                "desc": "Primary HTTP application server initializing Express, CORS, Socket.IO server, WebSocket real-time broadcast hub, REST route mounting, and SPA fallback."
            },
            {
                "path": "backend/src/config/db.js",
                "lang": "JavaScript (Node.js)",
                "role": "Database Connection & Seeder",
                "desc": "Mongoose connection manager handling MongoDB Atlas connection states, reconnect logic, and automated catalog seeding on cold start."
            },
            {
                "path": "backend/src/seed/seedData.js",
                "lang": "JavaScript (Node.js)",
                "role": "Comprehensive Database Seed",
                "desc": "Pre-populates MongoDB with demo personas (Admin, Cashier, Delivery, Customer), master product catalog across multiple categories, FEFO batches with varying expiry dates, and sample orders."
            }
        ]
    },
    {
        "title": "Section 3: Backend Data Models (Mongoose Schemas)",
        "description": "Object Data Modeling (ODM) schemas enforcing data integrity, validation rules, indexes, and virtual fields across all business domains.",
        "files": [
            {
                "path": "backend/src/models/User.js",
                "lang": "JavaScript (Mongoose)",
                "role": "User & Persona Schema",
                "desc": "User entity supporting roles ('admin', 'staff', 'delivery', 'customer'), bcrypt password hashing pre-save hooks, and password comparison methods."
            },
            {
                "path": "backend/src/models/Product.js",
                "lang": "JavaScript (Mongoose)",
                "role": "Master Catalog Product Schema",
                "desc": "Core product record storing SKU, name, category, base price, unit, minimum threshold for reorder, and estimated shelf life."
            },
            {
                "path": "backend/src/models/Batch.js",
                "lang": "JavaScript (Mongoose)",
                "role": "FEFO Inventory Batch Schema",
                "desc": "Critical inventory batch schema tracking lot number, manufacturing date, expiry date, current stock, barcode (EAN/Code-128), and dynamic clearance markdown percentage."
            },
            {
                "path": "backend/src/models/Order.js",
                "lang": "JavaScript (Mongoose)",
                "role": "Sales Order & POS Receipt Schema",
                "desc": "Orders collection handling line items, batch-level allocations, subtotal, taxes, discount deductions, payment method (UPI, Card, Cash), and delivery lifecycle."
            },
            {
                "path": "backend/src/models/PurchaseOrder.js",
                "lang": "JavaScript (Mongoose)",
                "role": "Supplier Purchase Order Schema",
                "desc": "Purchase order management schema linking supplier IDs, requested items, expected delivery dates, and receiving confirmation statuses."
            },
            {
                "path": "backend/src/models/Supplier.js",
                "lang": "JavaScript (Mongoose)",
                "role": "Vendor & Supplier Profile Schema",
                "desc": "Supplier registry schema storing company name, contact person, email, phone, address, categories supplied, and average lead time in days."
            },
            {
                "path": "backend/src/models/DailySale.js",
                "lang": "JavaScript (Mongoose)",
                "role": "Historical Sales Aggregations",
                "desc": "Aggregated daily product sales records utilized by the AI forecasting engine for regression analysis and demand trend modeling."
            },
            {
                "path": "backend/src/models/AuditLog.js",
                "lang": "JavaScript (Mongoose)",
                "role": "Immutable Security & Audit Trail",
                "desc": "Tamper-evident audit log storing timestamps, user ID, user role, action category, resource affected, client IP, and before/after mutation diffs."
            }
        ]
    },
    {
        "title": "Section 4: Backend Security & Compliance Middleware",
        "description": "Request validation, authentication interceptors, role-based authorization guards, and automated audit trail capture.",
        "files": [
            {
                "path": "backend/src/middleware/auth.js",
                "lang": "JavaScript (Node.js)",
                "role": "JWT Authentication & RBAC Guard",
                "desc": "Validates incoming Authorization Bearer tokens, decrypts JWT payloads, attaches the user object to the request, and enforces role access boundaries."
            },
            {
                "path": "backend/src/middleware/audit.js",
                "lang": "JavaScript (Node.js)",
                "role": "Audit Logging Middleware",
                "desc": "Express interceptor that captures high-impact state-changing HTTP requests (POST, PUT, DELETE) and records detailed audit log entries."
            }
        ]
    },
    {
        "title": "Section 5: Backend Business Logic & Core Services",
        "description": "Algorithmic engines: Google Gemini generative AI, heuristic demand forecasting, FEFO batch allocation with clearance pricing, and PDF invoice generation.",
        "files": [
            {
                "path": "backend/src/services/aiService.js",
                "lang": "JavaScript (Node.js)",
                "role": "Heuristic Simulation AI Engine",
                "desc": "Offline fallback engine providing statistical moving-average demand forecasting, heuristic recipe suggestions, and executive summaries when external AI APIs are unavailable."
            },
            {
                "path": "backend/src/services/geminiService.js",
                "lang": "JavaScript (Node.js)",
                "role": "Google Gemini Generative AI SDK",
                "desc": "Integrates with Google Gemini API (gemini-1.5-flash / pro) for natural language grocery assistance, conversational recipe generation from cart items, and automated supply-chain executive briefs."
            },
            {
                "path": "backend/src/services/inventoryService.js",
                "lang": "JavaScript (Node.js)",
                "role": "FEFO & Clearance Markdown Engine",
                "desc": "Core inventory algorithm: evaluates expiry dates, applies automated tiered markdowns (e.g., 20% off for 3-5 days, 50% off for 1-2 days), and executes First-Expired, First-Out batch picking during checkout."
            },
            {
                "path": "backend/src/services/invoiceService.js",
                "lang": "JavaScript (Node.js)",
                "role": "PDF Tax Invoice Generator",
                "desc": "PDFKit document generator that dynamically compiles formatted PDF tax invoices complete with store branding, cashier details, itemized batch numbers, GST breakdown, and payment confirmation."
            }
        ]
    },
    {
        "title": "Section 6: Backend API Controllers",
        "description": "Request dispatchers orchestrating business services, database queries, and response payloads across all platform domains.",
        "files": [
            {
                "path": "backend/src/controllers/authController.js",
                "lang": "JavaScript (Node.js)",
                "role": "Authentication Controller",
                "desc": "Handles customer registration, credential verification, JWT token signing, profile retrieval, and persona rapid switching."
            },
            {
                "path": "backend/src/controllers/productController.js",
                "lang": "JavaScript (Node.js)",
                "role": "Product Catalog Controller",
                "desc": "Manages product catalog CRUD, category queries, search indexing, and real-time aggregate stock calculation across active batches."
            },
            {
                "path": "backend/src/controllers/orderController.js",
                "lang": "JavaScript (Node.js)",
                "role": "Order Processing & POS Controller",
                "desc": "Executes atomic order creation with FEFO batch stock deduction, handles payment status changes, triggers invoice generation, and updates delivery statuses."
            },
            {
                "path": "backend/src/controllers/batchController.js",
                "lang": "JavaScript (Node.js)",
                "role": "FEFO Batch Inventory Controller",
                "desc": "Handles stock receiving/inwarding, barcode lookups, expiry inspection, manual and automated clearance markdown adjustments, and batch disposal."
            },
            {
                "path": "backend/src/controllers/reorderController.js",
                "lang": "JavaScript (Node.js)",
                "role": "Smart Reorder Controller",
                "desc": "Computes recommended reorder quantities based on current stock vs. reorder thresholds, and generates official Purchase Orders for vendor fulfillment."
            },
            {
                "path": "backend/src/controllers/supplierController.js",
                "lang": "JavaScript (Node.js)",
                "role": "Supplier Registry Controller",
                "desc": "Manages vendor records, supplier performance metrics, contact details, and catalog category affiliations."
            },
            {
                "path": "backend/src/controllers/aiController.js",
                "lang": "JavaScript (Node.js)",
                "role": "AI Features Controller",
                "desc": "API endpoints connecting the frontend with Google Gemini AI and heuristic models for chat assistance, recipe generation, and 7-day predictive demand forecasting."
            },
            {
                "path": "backend/src/controllers/analyticsController.js",
                "lang": "JavaScript (Node.js)",
                "role": "Executive Analytics Controller",
                "desc": "Aggregates key business metrics including total revenue, gross margin, at-risk expiring inventory value, top-selling SKUs, and sales trends."
            },
            {
                "path": "backend/src/controllers/auditController.js",
                "lang": "JavaScript (Node.js)",
                "role": "Audit Log Controller",
                "desc": "Serves filtered audit logs for compliance review, allowing administrators to filter by role, date range, action type, and severity."
            }
        ]
    },
    {
        "title": "Section 7: Backend API Route Definitions",
        "description": "Express router manifests connecting HTTP endpoints with corresponding controller actions and security middleware.",
        "files": [
            {
                "path": "backend/src/routes/authRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "Auth Route Manifest",
                "desc": "Defines endpoints for user registration (/register), login (/login), and user profile verification (/me)."
            },
            {
                "path": "backend/src/routes/productRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "Product Route Manifest",
                "desc": "Defines public catalog endpoints and protected staff/admin product creation and modification routes."
            },
            {
                "path": "backend/src/routes/orderRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "Order Route Manifest",
                "desc": "Defines order submission, POS checkout, invoice PDF streaming, and delivery status update endpoints."
            },
            {
                "path": "backend/src/routes/batchRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "Batch Route Manifest",
                "desc": "Defines FEFO batch inwarding, barcode lookup (/barcode/:code), and markdown update routes."
            },
            {
                "path": "backend/src/routes/reorderRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "Reorder & PO Route Manifest",
                "desc": "Defines routes for automated reorder recommendations and purchase order status transitions."
            },
            {
                "path": "backend/src/routes/supplierRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "Supplier Route Manifest",
                "desc": "Defines vendor registry CRUD endpoints accessible by staff and administrators."
            },
            {
                "path": "backend/src/routes/aiRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "AI Route Manifest",
                "desc": "Defines AI conversational assistant (/chat), recipe generator (/recipes), demand forecast (/forecast), and executive brief (/brief) routes."
            },
            {
                "path": "backend/src/routes/analyticsRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "Analytics Route Manifest",
                "desc": "Defines executive dashboard analytical endpoints including revenue trends and wastage metrics."
            },
            {
                "path": "backend/src/routes/auditRoutes.js",
                "lang": "JavaScript (Express Router)",
                "role": "Audit Route Manifest",
                "desc": "Defines protected administrator endpoints for querying system audit trails."
            }
        ]
    },
    {
        "title": "Section 8: Frontend Build Configuration & HTML Shell",
        "description": "Vite bundler configuration, Tailwind CSS design system tokens, PostCSS processing, NGINX reverse proxy, and single-page application HTML shell.",
        "files": [
            {
                "path": "frontend/package.json",
                "lang": "JSON",
                "role": "Frontend Dependencies Manifest",
                "desc": "Declares client dependencies including React 19, Lucide React icons, Axios, Socket.IO Client, Tailwind CSS, and Vite."
            },
            {
                "path": "frontend/vite.config.js",
                "lang": "JavaScript (Vite)",
                "role": "Vite Bundler Configuration",
                "desc": "Configures React fast-refresh plugin, build optimizations, asset chunking, and local API proxy to the Node.js backend."
            },
            {
                "path": "frontend/tailwind.config.js",
                "lang": "JavaScript (Tailwind)",
                "role": "Tailwind Design System Tokens",
                "desc": "Custom theme configuration defining brand color palettes (emerald, violet, amber, slate), custom border-radii, and animation keyframes."
            },
            {
                "path": "frontend/postcss.config.js",
                "lang": "JavaScript (PostCSS)",
                "role": "PostCSS Pipeline",
                "desc": "Configures Tailwind CSS and Autoprefixer processors for cross-browser styling compatibility."
            },
            {
                "path": "frontend/index.html",
                "lang": "HTML5",
                "role": "Single-Page Application Shell",
                "desc": "HTML5 entry document importing Google Fonts (Inter, Plus Jakarta Sans), setting responsive viewport meta tags, and mounting the root React DOM node."
            },
            {
                "path": "frontend/vercel.json",
                "lang": "JSON",
                "role": "Vercel Edge Rewrite Rules",
                "desc": "Configures SPA rewrite fallback rules for seamless client-side routing on Vercel Edge networks."
            },
            {
                "path": "frontend/nginx.conf",
                "lang": "NGINX Config",
                "role": "Production NGINX Web Server",
                "desc": "Production NGINX web server configuration providing gzip compression, client-side route fallback to index.html, and API reverse proxying."
            },
            {
                "path": "frontend/Dockerfile",
                "lang": "Dockerfile",
                "role": "Frontend Production Containerfile",
                "desc": "Multi-stage Dockerfile that builds the Vite static bundle and serves it via a lightweight Alpine NGINX container."
            },
            {
                "path": "frontend/.gitignore",
                "lang": "Git Configuration",
                "role": "Frontend Version Control Exclusions",
                "desc": "Excludes frontend node_modules, dist build outputs, and local development caches."
            },
            {
                "path": "frontend/.dockerignore",
                "lang": "Docker Configuration",
                "role": "Frontend Docker Exclusions",
                "desc": "Excludes build caches and local node_modules from frontend Docker context."
            },
            {
                "path": "frontend/.oxlintrc.json",
                "lang": "JSON",
                "role": "Linter Configuration",
                "desc": "Configures fast Oxlint linter rules for code cleanliness and syntax validation."
            },
            {
                "path": "frontend/README.md",
                "lang": "Markdown",
                "role": "Frontend Documentation",
                "desc": "Client-specific setup instructions, available npm scripts, and environment variable guidelines."
            }
        ]
    },
    {
        "title": "Section 9: Frontend Main Application & Global Styles",
        "description": "React root mounting, master application component with navigation shell, responsive layout wrappers, and global design styles.",
        "files": [
            {
                "path": "frontend/src/main.jsx",
                "lang": "JavaScript (React 19 JSX)",
                "role": "React 19 DOM Root Mount",
                "desc": "Mounts the React application into the root DOM element and wraps it with StrictMode and Context Providers."
            },
            {
                "path": "frontend/src/App.jsx",
                "lang": "JavaScript (React 19 JSX)",
                "role": "Master Application Shell & Router",
                "desc": "Central layout component featuring the omni-present navbar, quick persona switcher banner, modal controllers (Barcode Scanner, Gemini AI), and dynamic view routing."
            },
            {
                "path": "frontend/src/index.css",
                "lang": "CSS",
                "role": "Global Design Styles & Utilities",
                "desc": "Tailwind layer directives, custom utility classes, glassmorphism blur effects, custom scrollbar styling, and base typography."
            },
            {
                "path": "frontend/src/App.css",
                "lang": "CSS",
                "role": "Component Keyframe Animations",
                "desc": "Keyframe animations for pulse badges, scanner beam effects, AI glowing gradients, and modal transitions."
            }
        ]
    },
    {
        "title": "Section 10: Frontend State Management (Contexts)",
        "description": "Global React Context providers handling authentication state, user persona switching, shopping cart state, and real-time inventory validation.",
        "files": [
            {
                "path": "frontend/src/context/AuthContext.jsx",
                "lang": "JavaScript (React Context)",
                "role": "Authentication & Persona Context",
                "desc": "Manages global user session, JWT storage in localStorage, rapid 1-click persona switching (Admin, Staff, Delivery, Customer), and role authorization helpers."
            },
            {
                "path": "frontend/src/context/CartContext.jsx",
                "lang": "JavaScript (React Context)",
                "role": "Shopping Cart & Stock Context",
                "desc": "Manages shopping cart line items, quantity increment/decrement, coupon code application, stock limit validation, and local persistence."
            }
        ]
    },
    {
        "title": "Section 11: Frontend Shared Components & Modals",
        "description": "Reusable UI components: responsive navigation header, role-based access control gates, alerts drawer, hardware barcode scanner, and Gemini conversational assistant.",
        "files": [
            {
                "path": "frontend/src/components/Navbar.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Omnichannel Navigation Bar",
                "desc": "Top navigation component displaying brand identity, role-specific portal tabs, real-time stock alert badge, active persona badge, and shopping cart counter."
            },
            {
                "path": "frontend/src/components/RoleGate.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Declarative RBAC Guard Component",
                "desc": "Security wrapper component that checks user permissions against required roles before rendering protected views or displaying an unauthorized screen."
            },
            {
                "path": "frontend/src/components/AlertsDrawer.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Live Inventory Alerts Drawer",
                "desc": "Slide-out notification drawer displaying real-time WebSocket alerts for critically low stock levels and items nearing expiry."
            },
            {
                "path": "frontend/src/components/BarcodeScannerModal.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Interactive Barcode Scanner Modal",
                "desc": "Simulated hardware barcode scanner with visual laser scanning animation, quick-click barcode chips, and manual barcode entry for instant cashier checkout."
            },
            {
                "path": "frontend/src/components/GeminiAssistantModal.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Gemini AI Assistant & Chef Modal",
                "desc": "Multi-mode generative AI assistant modal supporting conversational grocery shopping Q&A, dietary meal planning, and automated recipe generation from active cart items."
            }
        ]
    },
    {
        "title": "Section 12: Frontend Application Pages & Portals",
        "description": "Full-featured application pages: Customer Storefront, In-Store Cashier POS, Cart Checkout, FEFO Inventory, AI Demand Forecasting, Purchase Orders, Suppliers, Admin BI Dashboard, Audit Trail, and Delivery Agent Portal.",
        "files": [
            {
                "path": "frontend/src/pages/Storefront.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Customer E-Commerce Catalog",
                "desc": "Customer-facing grocery shopping page featuring category filtering, live search, FEFO clearance discount badges, stock availability status, and one-click add to cart."
            },
            {
                "path": "frontend/src/pages/PosCounter.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "High-Speed Cashier POS Terminal",
                "desc": "In-store cashier point-of-sale interface featuring rapid barcode SKU scanning, line-item quantity adjustments, multi-mode payment collection, and instant PDF receipt generation."
            },
            {
                "path": "frontend/src/pages/CartCheckout.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Cart, Coupon & Checkout Portal",
                "desc": "Complete checkout experience with line item breakdown, coupon code discounts, delivery address entry, simulated payment gateway (UPI/Card/COD), and order confirmation."
            },
            {
                "path": "frontend/src/pages/InventoryManagement.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "FEFO Batch Inventory Controller",
                "desc": "Inventory management hub featuring batch receiving, expiry countdowns with color-coded warning badges (Fresh, Expiring Soon, Critical), manual clearance markdown controls, and barcode labels."
            },
            {
                "path": "frontend/src/pages/AIForecasting.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "7-Day AI Demand Forecasting",
                "desc": "Predictive analytics page presenting 7-day SKU demand projections, stockout probability estimates, seasonal trend charts, and automated reorder recommendations."
            },
            {
                "path": "frontend/src/pages/PurchaseOrders.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Supplier Purchase Order Hub",
                "desc": "Procurement portal for creating, approving, and tracking vendor purchase orders, with automated receiving workflow and inventory replenishment."
            },
            {
                "path": "frontend/src/pages/SuppliersPage.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Supplier Directory & Relations",
                "desc": "Vendor registry page managing supplier contacts, supported product categories, average delivery lead times, and reliability ratings."
            },
            {
                "path": "frontend/src/pages/AdminDashboard.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Executive BI Analytics & Gemini Brief",
                "desc": "Executive overview dashboard featuring revenue KPI cards, sales trajectory charts, expiry loss analytics, inventory health gauges, and one-click Gemini AI Executive Brief generation."
            },
            {
                "path": "frontend/src/pages/StaffAuditLogs.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Compliance & Security Audit Trail",
                "desc": "Audit investigation dashboard with real-time search, severity filtering, role categorization, and CSV export for compliance verification."
            },
            {
                "path": "frontend/src/pages/DeliveryPortal.jsx",
                "lang": "JavaScript (React JSX)",
                "role": "Delivery Driver Dispatch Portal",
                "desc": "Mobile-optimized delivery portal for drivers to view assigned dispatch routes, customer delivery addresses, contact details, and update delivery milestones."
            }
        ]
    },
    {
        "title": "Section 13: Frontend API Client & Fallback Engine",
        "description": "Network communication layer: Axios HTTP client, Socket.IO client event handling, and comprehensive standalone mock data generator.",
        "files": [
            {
                "path": "frontend/src/services/api.js",
                "lang": "JavaScript",
                "role": "Centralized API & WebSocket Client",
                "desc": "Axios HTTP wrapper configuring base URLs, auth token request interceptors, response error handling, and Socket.IO real-time event subscriptions with fallback logic."
            },
            {
                "path": "frontend/src/services/mockFallback.js",
                "lang": "JavaScript",
                "role": "Autonomous Offline Mock Engine",
                "desc": "Extensive in-memory mock engine providing realistic fallback responses for products, FEFO batches, orders, AI forecasting, suppliers, and analytics when the backend server is unreachable."
            }
        ]
    }
]

def build_word_document():
    print("Initializing Microsoft Word Document...")
    doc = docx.Document()

    # Set document margins (0.75 in / 54 pt)
    for s in doc.sections:
        s.top_margin = Inches(0.75)
        s.bottom_margin = Inches(0.75)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)
        s.page_width = Inches(8.5)
        s.page_height = Inches(11.0)

    section = doc.sections[0]
    section.different_first_page_header_footer = True

    # -------------------------------------------------------------
    # COVER PAGE
    # -------------------------------------------------------------
    print("Generating Cover Page...")
    
    # Top spacing
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(36)
    p_top.paragraph_format.space_after = Pt(12)

    # Super-title badge
    p_badge = doc.add_paragraph()
    p_badge.paragraph_format.space_before = Pt(0)
    p_badge.paragraph_format.space_after = Pt(6)
    run_badge = p_badge.add_run("ENTERPRISE QUICK-COMMERCE & RETAIL PLATFORM")
    run_badge.font.name = "Segoe UI"
    run_badge.font.size = Pt(11)
    run_badge.font.bold = True
    run_badge.font.color.rgb = BLUE_ACCENT

    # Main Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(8)
    run_title = p_title.add_run("SmartGrocery AI")
    run_title.font.name = "Segoe UI"
    run_title.font.size = Pt(32)
    run_title.font.bold = True
    run_title.font.color.rgb = NAVY_PRIMARY

    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(24)
    run_sub = p_sub.add_run("Complete System Architecture & Production Source Code")
    run_sub.font.name = "Segoe UI"
    run_sub.font.size = Pt(15)
    run_sub.font.color.rgb = SLATE_GRAY

    # Accent decorative bar table
    bar_tbl = doc.add_table(rows=1, cols=1)
    bar_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    bar_cell = bar_tbl.cell(0, 0)
    bar_cell.width = Inches(7.0)
    set_cell_background(bar_cell, ACCENT_BLUE_HEX)
    bar_p = bar_cell.paragraphs[0]
    bar_p.paragraph_format.space_before = Pt(2)
    bar_p.paragraph_format.space_after = Pt(2)
    bar_run = bar_p.add_run(" ")
    bar_run.font.size = Pt(2)

    # Metadata card table
    p_meta_space = doc.add_paragraph()
    p_meta_space.paragraph_format.space_before = Pt(20)
    p_meta_space.paragraph_format.space_after = Pt(4)

    meta_items = [
        ("Platform Edition", "Production Enterprise Retail & Quick-Commerce v1.0.0"),
        ("Frontend Technology", "React 19, Vite Bundler, Tailwind CSS Design System, Lucide React"),
        ("Backend Technology", "Node.js (v24), Express REST API, Socket.IO WebSockets, PDFKit"),
        ("Database Architecture", "MongoDB Atlas (Mongoose ODM) with Time-Series & Indexing"),
        ("Artificial Intelligence", "Google Gemini AI (1.5 Flash / Pro) + Heuristic Simulation Engine"),
        ("Inventory Strategy", "First-Expired, First-Out (FEFO) with Automated Clearance Markdowns"),
        ("Hardware / POS", "Dynamic EAN/Code-128 Barcode Engine & PDF Tax Invoice Generation"),
        ("Security & Compliance", "JSON Web Tokens (JWT), 4-Tier RBAC, Immutable Audit Logging"),
        ("Deployment Target", "Docker Multi-Stage Container, Render.com Unified Service, Vercel Edge"),
        ("Compiled Date", "September 2026"),
        ("Scope of Document", "100% Comprehensive Monorepo Source Code (Backend, Frontend, Configs)")
    ]

    meta_tbl = doc.add_table(rows=len(meta_items), cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False

    for idx, (label, val) in enumerate(meta_items):
        c0 = meta_tbl.cell(idx, 0)
        c1 = meta_tbl.cell(idx, 1)
        c0.width = Inches(2.2)
        c1.width = Inches(4.8)

        # Alternating background
        row_bg = META_BG_HEX if idx % 2 == 0 else "FFFFFF"
        set_cell_background(c0, row_bg)
        set_cell_background(c1, row_bg)
        set_cell_margins(c0, top=70, bottom=70, left=120, right=120)
        set_cell_margins(c1, top=70, bottom=70, left=120, right=120)
        set_cell_borders(c0, top="E2E8F0", bottom="E2E8F0", left="E2E8F0", right="E2E8F0", other_sz="2")
        set_cell_borders(c1, top="E2E8F0", bottom="E2E8F0", left="E2E8F0", right="E2E8F0", other_sz="2")

        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(label)
        r0.font.name = "Segoe UI"
        r0.font.size = Pt(9.5)
        r0.font.bold = True
        r0.font.color.rgb = DARK_TEXT

        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(val)
        r1.font.name = "Segoe UI"
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = DARK_TEXT

    # Quick metric callout blocks
    doc.add_paragraph().paragraph_format.space_before = Pt(20)

    stats_tbl = doc.add_table(rows=1, cols=3)
    stats_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    stats_tbl.autofit = False
    
    stats_data = [
        ("80+ Files", "Complete Source Code Files Included"),
        ("10,600+ Lines", "Clean, Documented Production Code"),
        ("4 Personas", "Admin, Cashier, Delivery, Customer")
    ]

    for s_idx, (stat_title, stat_desc) in enumerate(stats_data):
        scell = stats_tbl.cell(0, s_idx)
        scell.width = Inches(2.33)
        set_cell_background(scell, "F8FAFC")
        set_cell_borders(scell, top="3B82F6", bottom="CBD5E1", left="CBD5E1", right="CBD5E1", left_sz="2", other_sz="2")
        set_cell_margins(scell, top=100, bottom=100, left=100, right=100)

        sp = scell.paragraphs[0]
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(2)
        srun_top = sp.add_run(stat_title)
        srun_top.font.name = "Segoe UI"
        srun_top.font.size = Pt(16)
        srun_top.font.bold = True
        srun_top.font.color.rgb = NAVY_PRIMARY

        sp_sub = scell.add_paragraph()
        sp_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp_sub.paragraph_format.space_before = Pt(0)
        sp_sub.paragraph_format.space_after = Pt(0)
        srun_sub = sp_sub.add_run(stat_desc)
        srun_sub.font.name = "Segoe UI"
        srun_sub.font.size = Pt(8.5)
        srun_sub.font.color.rgb = SLATE_GRAY

    # End Cover Page
    doc.add_page_break()

    # -------------------------------------------------------------
    # CONFIGURE HEADERS & FOOTERS (For subsequent pages)
    # -------------------------------------------------------------
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run("SmartGrocery AI • System Architecture & Complete Source Code")
    hrun.font.name = "Segoe UI"
    hrun.font.size = Pt(8)
    hrun.font.color.rgb = RGBColor(148, 163, 184)

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    frun1 = fp.add_run("SmartGrocery AI Enterprise Platform\t\tPage ")
    frun1.font.name = "Segoe UI"
    frun1.font.size = Pt(8.5)
    frun1.font.color.rgb = SLATE_GRAY
    add_field(fp, "PAGE")
    frun2 = fp.add_run(" of ")
    frun2.font.name = "Segoe UI"
    frun2.font.size = Pt(8.5)
    frun2.font.color.rgb = SLATE_GRAY
    add_field(fp, "NUMPAGES")

    # -------------------------------------------------------------
    # ARCHITECTURAL OVERVIEW & SYSTEM DESIGN
    # -------------------------------------------------------------
    print("Generating Architectural Overview...")
    
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("Architectural Overview & System Blueprint")
    r.font.name = "Segoe UI"
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = NAVY_PRIMARY

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_before = Pt(4)
    p_intro.paragraph_format.space_after = Pt(8)
    p_intro.paragraph_format.line_spacing = 1.15
    r_intro = p_intro.add_run(
        "SmartGrocery AI is an omnichannel enterprise retail and quick-commerce platform built using a distributed full-stack architecture. "
        "The system fuses modern web interfaces (React 19, Vite, Tailwind CSS) with a high-throughput Node.js microservice layer, "
        "real-time bi-directional WebSockets (Socket.IO), and artificial intelligence capabilities powered by Google Gemini AI."
    )
    r_intro.font.name = "Segoe UI"
    r_intro.font.size = Pt(10)
    r_intro.font.color.rgb = DARK_TEXT

    # System Highlights
    highlights = [
        ("1. Google Gemini AI Assistant & Chef", 
         "Features an interactive natural language assistant for customers. Generates personalized culinary recipes and dietary meal plans based on items currently inside the shopping cart, reducing grocery waste and driving basket size."),
        ("2. First-Expired, First-Out (FEFO) Inventory Engine", 
         "Unlike conventional FIFO (first-in-first-out), FEFO prioritizes batches nearing expiration regardless of arrival date. Automatically calculates real-time clearance markdowns (20% to 50% discount) to clear at-risk stock before spoilage occurs."),
        ("3. AI Demand Forecasting & Predictive Analytics", 
         "Evaluates historical sales velocity and seasonality using regression models and Gemini AI prompts to forecast 7-day SKU demand, anticipate stockouts, and recommend proactive purchase orders."),
        ("4. Omnichannel In-Store POS & Dynamic Barcode Scanning", 
         "Enables cashiers to scan hardware or virtual EAN/Code-128 barcodes, apply real-time clearance discounts, accept simulated UPI/Card/Cash payments, and generate official downloadable PDF tax invoices via PDFKit."),
        ("5. Real-Time WebSockets Synchronization", 
         "A centralized Socket.IO event bus broadcasts stock level updates, critical low-stock alerts, order status transitions, and delivery milestones simultaneously across all cashier terminals, admin dashboards, and customer screens."),
        ("6. 4-Tier Role-Based Access Control (RBAC) & Audit Logs", 
         "Provides granular permission boundaries for Admin, Cashier/Staff, Delivery Driver, and Customer personas, backed by tamper-evident, immutable audit trail logging for all sensitive inventory and order mutations.")
    ]

    for title, body in highlights:
        hp = doc.add_paragraph()
        hp.paragraph_format.space_before = Pt(8)
        hp.paragraph_format.space_after = Pt(2)
        hrun = hp.add_run(title)
        hrun.font.name = "Segoe UI"
        hrun.font.size = Pt(11)
        hrun.font.bold = True
        hrun.font.color.rgb = BLUE_ACCENT

        bp = doc.add_paragraph()
        bp.paragraph_format.space_before = Pt(0)
        bp.paragraph_format.space_after = Pt(6)
        bp.paragraph_format.line_spacing = 1.15
        brun = bp.add_run(body)
        brun.font.name = "Segoe UI"
        brun.font.size = Pt(9.5)
        brun.font.color.rgb = DARK_TEXT

    # Persona Table
    hp_persona = doc.add_paragraph()
    hp_persona.paragraph_format.space_before = Pt(12)
    hp_persona.paragraph_format.space_after = Pt(4)
    r_per = hp_persona.add_run("System Personas & Access Permissions")
    r_per.font.name = "Segoe UI"
    r_per.font.size = Pt(13)
    r_per.font.bold = True
    r_per.font.color.rgb = NAVY_PRIMARY

    persona_table = doc.add_table(rows=5, cols=4)
    persona_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    persona_headers = ["Persona Role", "Demo Email", "Password", "System Permissions & Capabilities"]
    for c_idx, text in enumerate(persona_headers):
        cell = persona_table.cell(0, c_idx)
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.font.name = "Segoe UI"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = WHITE

    personas = [
        ("Admin", "admin@grocery.com", "admin123", "Executive BI dashboard, Gemini AI briefs, Purchase Orders, Staff management, Compliance audit logs"),
        ("Staff / Cashier", "staff@grocery.com", "staff123", "In-store POS counter, Barcode scanning, FEFO batch inwarding, Markdown adjustments"),
        ("Delivery Driver", "delivery@grocery.com", "delivery123", "Dispatch route portal, Customer addresses, Status updates (Out for delivery, Delivered)"),
        ("Customer", "customer@gmail.com", "customer123", "Storefront shopping, Gemini Chef recipe generation, Cart checkout, Order tracking")
    ]

    for r_idx, row in enumerate(personas, start=1):
        for c_idx, val in enumerate(row):
            cell = persona_table.cell(r_idx, c_idx)
            bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
            set_cell_borders(cell, top="E2E8F0", bottom="E2E8F0", left="E2E8F0", right="E2E8F0", other_sz="2")
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Segoe UI"
            r.font.size = Pt(8.5)
            r.font.color.rgb = DARK_TEXT
            if c_idx == 0:
                r.font.bold = True

    # -------------------------------------------------------------
    # MASTER SOURCE CODE INDEX
    # -------------------------------------------------------------
    print("Generating Master File Index Table...")
    doc.add_page_break()

    h_idx = doc.add_paragraph()
    h_idx.paragraph_format.space_before = Pt(12)
    h_idx.paragraph_format.space_after = Pt(4)
    r_idx_title = h_idx.add_run("Master Source Code Directory & File Index")
    r_idx_title.font.name = "Segoe UI"
    r_idx_title.font.size = Pt(20)
    r_idx_title.font.bold = True
    r_idx_title.font.color.rgb = NAVY_PRIMARY

    p_idx_desc = doc.add_paragraph()
    p_idx_desc.paragraph_format.space_before = Pt(0)
    p_idx_desc.paragraph_format.space_after = Pt(12)
    r_desc = p_idx_desc.add_run("The complete codebase is organized below into logical functional modules, detailing each file's relative path, primary language, role, and line count.")
    r_desc.font.name = "Segoe UI"
    r_desc.font.size = Pt(9.5)
    r_desc.font.color.rgb = SLATE_GRAY

    # Collect all file stats
    total_project_files = 0
    total_project_lines = 0
    all_file_entries = []

    for s_idx, sec in enumerate(SECTIONS, start=1):
        for f in sec["files"]:
            rel_path = f["path"]
            full_path = os.path.join(PROJECT_ROOT, rel_path)
            line_count = 0
            file_size_kb = 0
            if os.path.exists(full_path):
                try:
                    with open(full_path, 'r', encoding='utf-8', errors='ignore') as fp:
                        lines = fp.readlines()
                        line_count = len(lines)
                        file_size_kb = round(os.path.getsize(full_path) / 1024, 1)
                except Exception as e:
                    line_count = 0
            total_project_files += 1
            total_project_lines += line_count
            all_file_entries.append({
                "section": sec["title"].split(":")[0],
                "path": rel_path,
                "role": f["role"],
                "lang": f["lang"],
                "lines": line_count,
                "size_kb": file_size_kb
            })

    # File index table
    idx_table = doc.add_table(rows=len(all_file_entries) + 1, cols=5)
    idx_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    idx_table.autofit = False

    cols_width = [Inches(0.6), Inches(2.6), Inches(2.3), Inches(0.8), Inches(0.7)]
    idx_headers = ["#", "File Relative Path", "Role / Component Description", "Language", "Lines"]

    for c_idx, htext in enumerate(idx_headers):
        cell = idx_table.cell(0, c_idx)
        cell.width = cols_width[c_idx]
        set_cell_background(cell, "0F172A")
        set_cell_margins(cell, top=70, bottom=70, left=70, right=70)
        p = cell.paragraphs[0]
        r = p.add_run(htext)
        r.font.name = "Segoe UI"
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = WHITE
        if c_idx in [0, 4]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for idx, item in enumerate(all_file_entries, start=1):
        bg_col = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
        row_cells = [
            str(idx),
            item["path"],
            item["role"],
            item["lang"],
            f'{item["lines"]:,}'
        ]
        for c_idx, val in enumerate(row_cells):
            cell = idx_table.cell(idx, c_idx)
            cell.width = cols_width[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
            set_cell_borders(cell, top="E2E8F0", bottom="E2E8F0", left="E2E8F0", right="E2E8F0", other_sz="2")
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Consolas" if c_idx == 1 else "Segoe UI"
            r.font.size = Pt(7.5) if c_idx == 1 else Pt(8)
            r.font.color.rgb = DARK_TEXT
            if c_idx == 1:
                r.font.bold = True
            if c_idx in [0, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # -------------------------------------------------------------
    # SOURCE CODE SECTIONS & CODE LISTINGS
    # -------------------------------------------------------------
    print("Generating Source Code Listings...")

    file_counter = 0

    for sec_idx, sec in enumerate(SECTIONS, start=1):
        doc.add_page_break()

        # Section Header Banner Table
        sec_banner = doc.add_table(rows=1, cols=1)
        sec_banner.alignment = WD_TABLE_ALIGNMENT.CENTER
        sec_banner.autofit = False
        s_cell = sec_banner.cell(0, 0)
        s_cell.width = Inches(7.0)
        set_cell_background(s_cell, "1E3A8A")
        set_cell_margins(s_cell, top=140, bottom=140, left=180, right=180)
        set_cell_borders(s_cell, top="1E3A8A", bottom="1E3A8A", left="2563EB", right="1E3A8A", left_sz="36", other_sz="4")

        sp = s_cell.paragraphs[0]
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(2)
        srun_title = sp.add_run(sec["title"])
        srun_title.font.name = "Segoe UI"
        srun_title.font.size = Pt(16)
        srun_title.font.bold = True
        srun_title.font.color.rgb = WHITE

        sp_desc = s_cell.add_paragraph()
        sp_desc.paragraph_format.space_before = Pt(2)
        sp_desc.paragraph_format.space_after = Pt(0)
        srun_desc = sp_desc.add_run(sec["description"])
        srun_desc.font.name = "Segoe UI"
        srun_desc.font.size = Pt(9.5)
        srun_desc.font.color.rgb = RGBColor(226, 232, 240)

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

        # Loop through files in this section
        for f_idx, f_info in enumerate(sec["files"], start=1):
            file_counter += 1
            rel_path = f_info["path"]
            full_path = os.path.join(PROJECT_ROOT, rel_path)

            print(f"[{file_counter}/80] Processing: {rel_path}...")

            # Read file content
            content = ""
            lines = []
            if os.path.exists(full_path):
                try:
                    with open(full_path, 'r', encoding='utf-8', errors='ignore') as fp:
                        content = fp.read()
                        content = sanitize_file_content(rel_path, content)
                        lines = content.splitlines()
                except Exception as ex:
                    content = f"// Error reading file: {str(ex)}"
                    lines = [content]
            else:
                content = "// File not found at path"
                lines = [content]

            num_lines = len(lines)
            size_kb = round(len(content.encode('utf-8')) / 1024, 1)

            # File Header Box
            f_meta_tbl = doc.add_table(rows=2, cols=1)
            f_meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            f_meta_tbl.autofit = False

            # Top Row: File Name & Stats
            r0_cell = f_meta_tbl.cell(0, 0)
            r0_cell.width = Inches(7.0)
            set_cell_background(r0_cell, "0F172A")
            set_cell_margins(r0_cell, top=90, bottom=90, left=140, right=140)
            set_cell_borders(r0_cell, top="0F172A", bottom="1E293B", left="2563EB", right="0F172A", left_sz="24", other_sz="4")

            p_fhead = r0_cell.paragraphs[0]
            p_fhead.paragraph_format.space_before = Pt(0)
            p_fhead.paragraph_format.space_after = Pt(0)
            
            rf_icon = p_fhead.add_run("📄 ")
            rf_icon.font.size = Pt(10)
            
            rf_path = p_fhead.add_run(rel_path)
            rf_path.font.name = "Consolas"
            rf_path.font.size = Pt(11)
            rf_path.font.bold = True
            rf_path.font.color.rgb = WHITE

            rf_meta = p_fhead.add_run(f"   |   {f_info['lang']}   •   {num_lines:,} lines ({size_kb} KB)")
            rf_meta.font.name = "Segoe UI"
            rf_meta.font.size = Pt(8.5)
            rf_meta.font.color.rgb = RGBColor(148, 163, 184)

            # Bottom Row: Role & Architectural Description
            r1_cell = f_meta_tbl.cell(1, 0)
            r1_cell.width = Inches(7.0)
            set_cell_background(r1_cell, "F1F5F9")
            set_cell_margins(r1_cell, top=70, bottom=70, left=140, right=140)
            set_cell_borders(r1_cell, top="E2E8F0", bottom="CBD5E1", left="2563EB", right="CBD5E1", left_sz="24", other_sz="2")

            p_fdesc = r1_cell.paragraphs[0]
            p_fdesc.paragraph_format.space_before = Pt(0)
            p_fdesc.paragraph_format.space_after = Pt(0)
            
            r_role = p_fdesc.add_run(f"{f_info['role']}: ")
            r_role.font.name = "Segoe UI"
            r_role.font.size = Pt(8.5)
            r_role.font.bold = True
            r_role.font.color.rgb = DARK_TEXT

            r_fbody = p_fdesc.add_run(f_info['desc'])
            r_fbody.font.name = "Segoe UI"
            r_fbody.font.size = Pt(8.5)
            r_fbody.font.color.rgb = DARK_TEXT

            # Spacing between header box and code box
            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_before = Pt(2)
            p_space.paragraph_format.space_after = Pt(2)

            # Code Box (1-cell table with line numbering)
            code_tbl = doc.add_table(rows=1, cols=1)
            code_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            code_tbl.autofit = False

            code_cell = code_tbl.cell(0, 0)
            code_cell.width = Inches(7.0)
            set_cell_background(code_cell, "F8FAFC")
            set_cell_margins(code_cell, top=80, bottom=80, left=100, right=100)
            set_cell_borders(code_cell, top="E2E8F0", bottom="E2E8F0", left="3B82F6", right="E2E8F0", left_sz="16", other_sz="2")

            # Determine line number padding width
            num_digits = max(3, len(str(num_lines)))

            for l_idx, line in enumerate(lines, start=1):
                p_code = code_cell.paragraphs[0] if l_idx == 1 else code_cell.add_paragraph()
                p_code.paragraph_format.space_before = Pt(0)
                p_code.paragraph_format.space_after = Pt(0)
                p_code.paragraph_format.line_spacing = 1.0

                # Line number prefix
                r_num = p_code.add_run(f"{l_idx:>{num_digits}} | ")
                r_num.font.name = "Consolas"
                r_num.font.size = Pt(8.0)
                r_num.font.color.rgb = RGBColor(148, 163, 184)

                # Code content
                r_text = p_code.add_run(line)
                r_text.font.name = "Consolas"
                r_text.font.size = Pt(8.0)
                r_text.font.color.rgb = DARK_TEXT

            # Bottom margin after code table
            p_after = doc.add_paragraph()
            p_after.paragraph_format.space_before = Pt(8)
            p_after.paragraph_format.space_after = Pt(16)

    print(f"Saving final document to: {OUTPUT_FILE}...")
    doc.save(OUTPUT_FILE)
    file_size_mb = round(os.path.getsize(OUTPUT_FILE) / (1024 * 1024), 2)
    print(f"Document successfully created! File size: {file_size_mb} MB")
    print(f"Total files documented: {total_project_files}")
    print(f"Total lines of code included: {total_project_lines:,}")

if __name__ == "__main__":
    build_word_document()
