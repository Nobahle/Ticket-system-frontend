from flask import Flask, render_template, request, session, redirect, url_for, send_file
import sqlite3
import string
import os
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.barcharts import VerticalBarChart
import openpyxl
from io import BytesIO

IT_KEYWORDS = [
    "password","computer","wifi","laptop","printer","email",
    "system","software","network","login","internet",
    "server","bug","error","update","pc"
]

FINANCE_KEYWORDS = [
    "salary","payment","payslip","invoice","refund",
    "budget","expense","tax","bank","billing"
]

HR_KEYWORDS = [
    "leave","holiday","vacation","sick","promotion",
    "training","resignation","contract","benefits"
]

OPERATIONS_KEYWORDS = [
    "office","chair","desk","aircon","maintenance",
    "cleaning","electricity","water","parking","security"
]

URGENT_WORDS = [
    "urgent","asap","immediately","now","critical",
    "system down","not working","failed","error"
]

FRIENDLY_WORDS = [
    "hi","hello","please","could you","thank you","kindly"
]


def init_db():
    conn = sqlite3.connect("tickets.db")
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        email TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tickets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ticket_text TEXT,
        category TEXT,
        tone TEXT,
        response TEXT,
        user_id INTEGER,
        status TEXT DEFAULT 'Open',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users(id)
    )
    """)

    conn.commit()
    conn.close()

# NOTE: init_db() will run after get_db and migrate_db are defined.

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "your-secret-key-here")

# Login required decorator
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# Admin required decorator
def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT role FROM users WHERE id = ?", (session['user_id'],))
        user = cursor.fetchone()
        conn.close()
        
        if not user or user[0] != 'admin':
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

# Define departments
DEPARTMENTS = ["IT", "Finance", "HR", "Operations"]

# DATABASE
DB_PATH = os.environ.get(
    "DB_PATH",
    "/tmp/tickets.db" if os.environ.get("VERCEL", "0") == "1" else os.path.join(os.path.abspath(os.path.dirname(__file__)), "tickets.db")
)

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def migrate_db():
    conn = get_db()
    cursor = conn.cursor()

    # Ensure column user_id exists
    cursor.execute("PRAGMA table_info('tickets')")
    cols = [row[1] for row in cursor.fetchall()]

    if "user_id" not in cols:
        cursor.execute("ALTER TABLE tickets ADD COLUMN user_id INTEGER")

    if "status" not in cols:
        cursor.execute("ALTER TABLE tickets ADD COLUMN status TEXT DEFAULT 'Open'")

    conn.commit()
    conn.close()

# Run initialization and migration when app starts
init_db()
migrate_db()

# TICKET CLASSIFICATION
def classify_ticket(ticket):
    text = ticket.lower()
    text = text.translate(str.maketrans('', '', string.punctuation))

    detected_departments = set()

    for word in IT_KEYWORDS:
        if word in text:
            detected_departments.add("IT")

    for word in FINANCE_KEYWORDS:
        if word in text:
            detected_departments.add("Finance")

    for word in HR_KEYWORDS:
        if word in text:
            detected_departments.add("HR")

    for word in OPERATIONS_KEYWORDS:
        if word in text:
            detected_departments.add("Operations")

    if not detected_departments:
        return ["Unrecognized"]
    else:
        return list(detected_departments)

# TONE DETECTION
def detect_tone(ticket):
    text = ticket.lower()
    if any(word in text for word in URGENT_WORDS):
        return "Urgent"
    if any(word in text for word in FRIENDLY_WORDS):
        return "Friendly"
    return "Formal"

# RESPONSE - AI-Enhanced with Context Awareness
def generate_response(categories, tone="Formal", ticket_text=""):
    if isinstance(categories, str):
        categories = [categories]

    tone = tone.capitalize()
    if tone not in ["Friendly", "Formal", "Urgent"]:
        tone = "Formal"

    ticket_lower = ticket_text.lower()
    
    if categories == ["Unrecognized"]:
        if tone == "Friendly":
            return "Thank you for reaching out! 😊 We received your message and our team is reviewing it carefully to route it to the right department. We'll be in touch shortly!"
        if tone == "Urgent":
            return "⚠️ Your urgent request has been received! We're prioritizing this and assigning it to the appropriate team immediately. You can expect a quick response."
        return "Your request has been received and logged into our system. We will review it and route it to the appropriate department for resolution."

    categories_str = ", ".join(categories)
    
    # Context-aware responses based on ticket content
    if "IT" in categories and "wifi" in ticket_lower:
        if tone == "Urgent":
            return f"🔴 High Priority: Your connectivity issue has been escalated to {categories_str} team. They will diagnose and resolve this immediately."
        elif tone == "Friendly":
            return f"✓ No worries! Your WiFi/connectivity issue has been sent to our {categories_str} experts. They'll get you back online quickly!"
        else:
            return f"Your technical issue has been assigned to {categories_str}. They will troubleshoot and provide a solution promptly."
    
    if "Finance" in categories and "salary" in ticket_lower:
        if tone == "Urgent":
            return f"🔴 Priority: Your financial matter has been marked urgent and sent to {categories_str}. Expect immediate attention."
        elif tone == "Friendly":
            return f"✓ Got it! Your payroll concern is with our {categories_str} team now. We appreciate your patience!"
        else:
            return f"Your financial inquiry has been forwarded to {categories_str}. They will review and address this promptly."
    
    if "HR" in categories and "leave" in ticket_lower:
        if tone == "Urgent":
            return f"🔴 Time-sensitive: Your leave request has been fast-tracked to {categories_str} for immediate processing."
        elif tone == "Friendly":
            return f"✓ Perfect! Your leave request is now with our {categories_str} team. We'll process it as soon as possible!"
        else:
            return f"Your leave request has been submitted to {categories_str}. Processing typically takes 1-2 business days."
    
    # Generic context-aware responses
    if tone == "Friendly":
        return f"✓ Excellent! Your request has been assigned to {categories_str}. We appreciate you reaching out and will help you soon!"
    elif tone == "Urgent":
        return f"🔴 URGENT: Your request is now with {categories_str} with high priority status. Expect quick action."
    else:
        return f"Your request has been successfully assigned to {categories_str}. They will handle this according to protocol."

# HOME PAGE (Login page redirect)
@app.route("/")
def home():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

# LOGIN PAGE
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ?", (username,))
        user = cursor.fetchone()
        conn.close()
        
        if not user:
            return render_template("login.html", error="Invalid username or password")

        if not check_password_hash(user[3], password):
            return render_template("login.html", error="Invalid username or password")

        # Auto-detect role from database
        session['user_id'] = user[0]
        session['username'] = user[1]
        session['role'] = user[4]  # Role comes from DB, not user input
        return redirect(url_for('dashboard'))

    return render_template("login.html")

# REGISTER PAGE
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form.get("username")
        email = request.form.get("email")
        password = request.form.get("password")
        confirm_password = request.form.get("confirm_password")
        role = "user"  # Default role is user
        
        # Validation
        if not username or not email or not password:
            return render_template("register.html", error="All fields are required")
        
        if password != confirm_password:
            return render_template("register.html", error="Passwords do not match")
        
        if len(password) < 6:
            return render_template("register.html", error="Password must be at least 6 characters")
        
        conn = get_db()
        cursor = conn.cursor()
        
        try:
            hashed_password = generate_password_hash(password)
            cursor.execute(
                "INSERT INTO users (username, email, password, role) VALUES (?, ?, ?, ?)",
                (username, email, hashed_password, role)
            )
            conn.commit()
            conn.close()
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            conn.close()
            return render_template("register.html", error="Username or email already exists")
    
    return render_template("register.html")

# LOGOUT
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for('login'))

# INITIALIZE ADMIN (First time setup)
@app.route("/init_admin")
def init_admin():
    conn = get_db()
    cursor = conn.cursor()
    
    # Check if admin already exists
    cursor.execute("SELECT * FROM users WHERE role = 'admin'")
    admin = cursor.fetchone()
    
    if admin:
        conn.close()
        return "Admin already exists!"
    
    # Create default admin user
    hashed_password = generate_password_hash("Admin123")
    try:
        cursor.execute(
            "INSERT INTO users (username, email, password, role) VALUES (?, ?, ?, ?)",
            ("administrator", "administrator@ticketsystem.com", hashed_password, "admin")
        )
        conn.commit()
        conn.close()
        return "Admin user created successfully! Username: administrator | Password: Admin123<br><a href='/login'>Go to Login</a>"
    except sqlite3.IntegrityError:
        conn.close()
        return "Admin user already exists! If role is wrong, visit /fix_admin"

# FIX ADMIN ROLE (if administrator was registered as user)
@app.route("/fix_admin")
def fix_admin():
    conn = get_db()
    cursor = conn.cursor()
    
    # Force administrator to be admin role
    cursor.execute("UPDATE users SET role = 'admin' WHERE username = 'administrator'")
    conn.commit()
    
    # Verify
    cursor.execute("SELECT username, role FROM users WHERE username = 'administrator'")
    result = cursor.fetchone()
    conn.close()
    
    if result:
        return f"✅ Fixed! Administrator role set to: {result[1]}<br><a href='/login'>Go to Login</a>"
    else:
        return "❌ Administrator account not found. Please run /init_admin first."

# DASHBOARD
@app.route("/dashboard")
@login_required
def dashboard():
    conn = get_db()
    cursor = conn.cursor()
    
    analytics = None
    weekly_insights = None
    
    if session.get('role') == 'admin':
        # Admin sees all tickets with user info
        cursor.execute("""
            SELECT t.*, u.username FROM tickets t 
            LEFT JOIN users u ON t.user_id = u.id 
            ORDER BY t.created_at DESC
        """)
        tickets = cursor.fetchall()
        
        # Calculate analytics per department
        analytics = {}
        for dept in DEPARTMENTS:
            cursor.execute("""
                SELECT status, COUNT(*) FROM tickets 
                WHERE category LIKE ? 
                GROUP BY status
            """, (f'%{dept}%',))
            stats = cursor.fetchall()
            total = sum(count for _, count in stats)
            analytics[dept] = {
                'total': total,
                'open': dict(stats).get('Open', 0),
                'in_progress': dict(stats).get('In Progress', 0),
                'closed': dict(stats).get('Closed', 0)
            }
        
        # Calculate weekly insights (last 7 days)
        cursor.execute("""
            SELECT COUNT(*) FROM tickets 
            WHERE DATE(created_at) >= DATE('now', '-7 days')
        """)
        weekly_total = cursor.fetchone()[0]
        
        cursor.execute("""
            SELECT status, COUNT(*) FROM tickets 
            WHERE DATE(created_at) >= DATE('now', '-7 days')
            GROUP BY status
        """)
        weekly_status = dict(cursor.fetchall())
        
        cursor.execute("""
            SELECT COUNT(DISTINCT user_id) FROM tickets 
            WHERE DATE(created_at) >= DATE('now', '-7 days')
        """)
        weekly_users = cursor.fetchone()[0]
        
        cursor.execute("""
            SELECT COUNT(*) FROM tickets 
            WHERE status = 'Closed' 
            AND DATE(created_at) >= DATE('now', '-7 days')
        """)
        weekly_resolved = cursor.fetchone()[0]
        
        cursor.execute("""
            SELECT category, COUNT(*) as count FROM tickets 
            WHERE DATE(created_at) >= DATE('now', '-7 days')
            GROUP BY category
            ORDER BY count DESC
            LIMIT 1
        """)
        top_dept_result = cursor.fetchone()
        top_dept = top_dept_result[0] if top_dept_result else "None"
        
        weekly_insights = {
            'total': weekly_total,
            'open': weekly_status.get('Open', 0),
            'in_progress': weekly_status.get('In Progress', 0),
            'closed': weekly_status.get('Closed', 0),
            'users_submitted': weekly_users,
            'resolved': weekly_resolved,
            'top_department': top_dept,
            'closure_rate': f"{(weekly_resolved / weekly_total * 100):.0f}%" if weekly_total > 0 else "0%"
        }
    else:
        # User sees only their tickets
        cursor.execute("SELECT * FROM tickets WHERE user_id = ? ORDER BY created_at DESC", (session['user_id'],))
        tickets = cursor.fetchall()
    
    conn.close()
    
    return render_template("dashboard.html", tickets=tickets, username=session['username'], role=session['role'], analytics=analytics, weekly_insights=weekly_insights)

# SUBMIT
@app.route("/submit", methods=["POST"])
@login_required
def submit():
    ticket = request.form["ticket"]

    categories = classify_ticket(ticket)
    tone = detect_tone(ticket)
    
    # If category is unrecognized, show selection form
    if categories == ["Unrecognized"]:
        session['ticket_text'] = ticket
        session['ticket_tone'] = tone
        return render_template("category_selection.html", departments=DEPARTMENTS)
    
    categories_str = ",".join(categories)
    response = generate_response(categories, tone, ticket)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO tickets (ticket_text, category, tone, response, user_id, status) VALUES (?, ?, ?, ?, ?, ?)",
        (ticket, categories_str, tone, response, session['user_id'], "Open")
    )
    conn.commit()
    conn.close()

    return render_template("result.html", categories=categories, response=response, tone=tone)

# VIEW
@app.route("/view")
@login_required
def view():
    conn = get_db()
    cursor = conn.cursor()

    try:
        if session['role'] == 'admin':
            cursor.execute("SELECT * FROM tickets ORDER BY created_at DESC")
        else:
            cursor.execute("SELECT * FROM tickets WHERE user_id = ? ORDER BY created_at DESC", (session['user_id'],))

        tickets = cursor.fetchall()
    except sqlite3.DatabaseError as db_err:
        conn.close()
        return render_template("view.html", tickets=[], role=session.get('role', 'user'), error=str(db_err))

    conn.close()

    return render_template("view.html", tickets=tickets, role=session.get('role', 'user'))

# SELECT CATEGORY
@app.route("/select_category", methods=["POST"])
@login_required
def select_category():
    selected_categories = request.form.getlist("category")
    ticket = session.get('ticket_text', '')
    tone = session.get('ticket_tone', 'Neutral')
    
    # Validate the selected categories
    if not selected_categories:
        return render_template("category_selection.html", departments=DEPARTMENTS, error="Please select at least one department")

    for cat in selected_categories:
        if cat not in DEPARTMENTS:
            return render_template("category_selection.html", departments=DEPARTMENTS, error="Invalid category selected")
    
    response = generate_response(selected_categories, tone, session.get('ticket_text', ''))
    categories_str = ",".join(selected_categories)
    
    # Save to database with user-selected categories
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO tickets (ticket_text, category, tone, response, user_id, status) VALUES (?, ?, ?, ?, ?, ?)",
        (ticket, categories_str, tone, response, session['user_id'], "Open")
    )
    conn.commit()
    conn.close()
    
    # Clear session
    session.pop('ticket_text', None)
    session.pop('ticket_tone', None)
    
    return render_template("result.html", categories=selected_categories, response=response, tone=tone)

# DELETE
@app.route("/delete_ticket/<int:ticket_id>", methods=["POST"])
@login_required
def delete_ticket(ticket_id):
    conn = get_db()
    cursor = conn.cursor()
    
    # Check if ticket exists and belongs to user (or user is admin)
    if session['role'] == 'admin':
        cursor.execute("SELECT id FROM tickets WHERE id = ?", (ticket_id,))
    else:
        cursor.execute("SELECT id FROM tickets WHERE id = ? AND user_id = ?", (ticket_id, session['user_id']))
    
    ticket = cursor.fetchone()
    
    if ticket:
        cursor.execute("DELETE FROM tickets WHERE id = ?", (ticket_id,))
        conn.commit()
        conn.close()
        return redirect(url_for('view'))
    else:
        conn.close()
        return "Ticket not found or access denied", 403

# ADMIN: update ticket status
@app.route("/update_status/<int:ticket_id>", methods=["POST"])
@login_required
def update_status(ticket_id):
    if session.get('role') != 'admin':
        return "Access denied", 403

    new_status = request.form.get('status')
    allowed_statuses = ['Open', 'In Progress', 'Closed']
    if new_status not in allowed_statuses:
        return "Invalid status", 400

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE tickets SET status = ? WHERE id = ?", (new_status, ticket_id))
    conn.commit()
    conn.close()

    return redirect(url_for('view'))

# REPORTS PAGE
@app.route("/reports", methods=["GET", "POST"])
@login_required
@admin_required
def reports():
    if request.method == "POST":
        department = request.form.get('department', 'All')
        date_from = request.form.get('date_from')
        date_to = request.form.get('date_to')
        report_format = request.form.get('format', 'pdf')
        
        # Generate report based on format
        if report_format == 'pdf':
            return generate_pdf_report(department, date_from, date_to)
        elif report_format == 'csv':
            return generate_csv_report(department, date_from, date_to)
    
    return render_template("reports.html", departments=DEPARTMENTS)

# Generate PDF Report
def generate_pdf_report(department, date_from, date_to):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics import renderPDF
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=0.5*inch, rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle('Title', parent=styles['Title'], fontSize=24, spaceAfter=30, alignment=1)
    heading1_style = ParagraphStyle('Heading1', parent=styles['Heading1'], fontSize=18, spaceAfter=20, textColor=colors.darkblue)
    heading2_style = ParagraphStyle('Heading2', parent=styles['Heading2'], fontSize=14, spaceAfter=15, textColor=colors.darkgreen)
    normal_style = styles['Normal']
    
    elements = []
    
    # Get data
    conn = get_db()
    cursor = conn.cursor()
    
    query = """
        SELECT t.id, t.ticket_text, t.category, t.tone, t.status, t.created_at, u.username
        FROM tickets t
        LEFT JOIN users u ON t.user_id = u.id
    """
    params = []
    
    if department != 'All':
        query += " WHERE t.category LIKE ?"
        params.append(f'%{department}%')
    
    if date_from and date_to:
        if 'WHERE' in query:
            query += " AND DATE(t.created_at) BETWEEN ? AND ?"
        else:
            query += " WHERE DATE(t.created_at) BETWEEN ? AND ?"
        params.extend([date_from, date_to])
    
    query += " ORDER BY t.created_at DESC"
    
    cursor.execute(query, params)
    tickets = cursor.fetchall()
    
    # Calculate statistics
    total_tickets = len(tickets)
    open_count = sum(1 for t in tickets if t[4] == 'Open')
    in_progress_count = sum(1 for t in tickets if t[4] == 'In Progress')
    closed_count = sum(1 for t in tickets if t[4] == 'Closed')
    
    # Department breakdown
    dept_stats = {}
    for ticket in tickets:
        dept = ticket[2] if ticket[2] else 'Unassigned'
        if dept not in dept_stats:
            dept_stats[dept] = {'total': 0, 'open': 0, 'in_progress': 0, 'closed': 0}
        dept_stats[dept]['total'] += 1
        if ticket[4] == 'Open':
            dept_stats[dept]['open'] += 1
        elif ticket[4] == 'In Progress':
            dept_stats[dept]['in_progress'] += 1
        elif ticket[4] == 'Closed':
            dept_stats[dept]['closed'] += 1
    
    # ===== EXECUTIVE SUMMARY =====
    elements.append(Paragraph("EXECUTIVE SUMMARY", title_style))
    elements.append(Spacer(1, 20))
    
    report_title = f"Ticket System Performance Report - {department} Department"
    elements.append(Paragraph(report_title, heading1_style))
    
    date_range = f"Report Period: {date_from} to {date_to}" if date_from and date_to else "Report Period: All Time"
    elements.append(Paragraph(date_range, normal_style))
    elements.append(Spacer(1, 15))
    
    # Key metrics
    summary_title = Paragraph("This report provides a comprehensive analysis of ticket system performance for the selected period.", ParagraphStyle('SummaryTitle', parent=styles['BodyText'], fontSize=12, leading=16, spaceAfter=15))
    elements.append(summary_title)
    
    metrics = [
        ('Total Tickets', str(total_tickets), colors.HexColor('#3b82f6')),
        ('Open Tickets', str(open_count), colors.HexColor('#ef4444')),
        ('In Progress', str(in_progress_count), colors.HexColor('#f59e0b')),
        ('Closed Tickets', str(closed_count), colors.HexColor('#10b981')),
        ('Closure Rate', f"{(closed_count / total_tickets * 100):.1f}%" if total_tickets > 0 else '0%', colors.HexColor('#6366f1'))
    ]
    
    card_cells = []
    for label, value, bg in metrics:
        card = Table([
            [Paragraph(f"<b>{label}</b>", ParagraphStyle('CardLabel', parent=styles['BodyText'], fontSize=9, textColor=colors.white))],
            [Paragraph(value, ParagraphStyle('CardValue', parent=styles['BodyText'], fontSize=16, textColor=colors.white))]
        ], colWidths=2.1*inch)
        card.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), bg),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.white),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]))
        card_cells.append(card)
    
    elements.append(Table([card_cells], colWidths=[2.1*inch] * len(card_cells), hAlign='LEFT', spaceBefore=10, spaceAfter=20))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("Key performance indicators are shown in the cards above. These summarize ticket volume, urgency, and closure success.", ParagraphStyle('SummaryFooter', parent=styles['BodyText'], fontSize=10, textColor=colors.grey, spaceAfter=10)))
    elements.append(PageBreak())
    
    # ===== DEPARTMENT OVERVIEW =====
    elements.append(Paragraph("DEPARTMENT OVERVIEW", title_style))
    elements.append(Spacer(1, 20))
    
    # Overall statistics table
    overview_data = [
        ['Metric', 'Count', 'Percentage'],
        ['Total Tickets', str(total_tickets), '100%'],
        ['Open Tickets', str(open_count), f"{(open_count/total_tickets*100):.1f}%" if total_tickets > 0 else "0%"],
        ['In Progress', str(in_progress_count), f"{(in_progress_count/total_tickets*100):.1f}%" if total_tickets > 0 else "0%"],
        ['Closed Tickets', str(closed_count), f"{(closed_count/total_tickets*100):.1f}%" if total_tickets > 0 else "0%"]
    ]
    
    overview_table = Table(overview_data, colWidths=[2*inch, 1*inch, 1.5*inch])
    overview_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
    ]))
    
    elements.append(overview_table)
    elements.append(Spacer(1, 20))
    
    # Department breakdown table
    dept_data = [['Department', 'Total', 'Open', 'In Progress', 'Closed', 'Closure Rate']]
    for dept, stats in sorted(dept_stats.items()):
        closure_rate = f"{(stats['closed']/stats['total']*100):.1f}%" if stats['total'] > 0 else "0%"
        dept_data.append([
            dept,
            str(stats['total']),
            str(stats['open']),
            str(stats['in_progress']),
            str(stats['closed']),
            closure_rate
        ])
    
    dept_table = Table(dept_data, colWidths=[1.5*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch, 1*inch])
    dept_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
    ]))
    
    elements.append(Paragraph("Department Performance Breakdown", heading2_style))
    elements.append(dept_table)
    elements.append(PageBreak())
    
    # ===== VISUAL BREAKDOWN =====
    elements.append(Paragraph("VISUAL BREAKDOWN", title_style))
    elements.append(Spacer(1, 20))
    
    if total_tickets > 0:
        drawing = Drawing(500, 280)
        
        # Pie Chart
        pie = Pie()
        pie.x = 30
        pie.y = 50
        pie.width = 150
        pie.height = 150
        pie.data = [open_count if open_count > 0 else 0.1, in_progress_count if in_progress_count > 0 else 0.1, closed_count if closed_count > 0 else 0.1]
        pie.labels = ['Open', 'In Progress', 'Closed']
        pie.slices.strokeWidth = 0.5
        pie.slices.strokeColor = colors.white
        
        # Color slices
        for i in range(len(pie.slices)):
            if i == 0:
                pie.slices[i].fillColor = colors.HexColor('#ef4444')
            elif i == 1:
                pie.slices[i].fillColor = colors.HexColor('#f97316')
            else:
                pie.slices[i].fillColor = colors.HexColor('#22c55e')
        
        drawing.add(pie)
        
        # Bar Chart
        bar = VerticalBarChart()
        bar.x = 240
        bar.y = 40
        bar.width = 230
        bar.height = 170
        bar.data = [[open_count, in_progress_count, closed_count]]
        bar.categoryAxis.categoryNames = ['Open', 'In Progress', 'Closed']
        bar.valueAxis.valueMin = 0
        bar.valueAxis.valueMax = max(total_tickets, 1) * 1.1
        bar.barWidth = 25
        
        # Style the bars
        bar.bars[0][0].fillColor = colors.HexColor('#ef4444')
        bar.bars[0][1].fillColor = colors.HexColor('#f97316')
        bar.bars[0][2].fillColor = colors.HexColor('#22c55e')
        
        # Add some padding and labels
        bar.categoryAxis.labels.angle = 0
        bar.title = "Ticket Status Distribution"
        
        drawing.add(bar)
        
        elements.append(drawing)
        elements.append(Spacer(1, 15))
        elements.append(Paragraph("Pie chart (left) shows proportion of ticket states. Bar chart (right) displays volume comparison.", ParagraphStyle('VisualNote', parent=styles['BodyText'], fontSize=10, textColor=colors.grey, spaceAfter=10)))
    else:
        elements.append(Paragraph("No ticket data available for visualization.", normal_style))
    
    elements.append(PageBreak())
    
    # ===== DEPARTMENT DEEP-DIVE REPORTS =====
    elements.append(Paragraph("DEPARTMENT DEEP-DIVE REPORTS", title_style))
    elements.append(Spacer(1, 20))
    
    for dept, stats in sorted(dept_stats.items()):
        elements.append(Paragraph(f"Department: {dept}", heading1_style))
        
        dept_summary = f"""
        <b>Performance Summary for {dept}:</b><br/>
        • Total Tickets: {stats['total']}<br/>
        • Open Tickets: {stats['open']}<br/>
        • In Progress: {stats['in_progress']}<br/>
        • Closed Tickets: {stats['closed']}<br/>
        • Closure Rate: {f"{(stats['closed']/stats['total']*100):.1f}%" if stats['total'] > 0 else "0%"}<br/>
        """
        elements.append(Paragraph(dept_summary, normal_style))
        elements.append(Spacer(1, 15))
        
        # Department tickets table
        dept_tickets = [t for t in tickets if (t[2] == dept or (dept == 'Unassigned' and not t[2]))]
        if dept_tickets:
            ticket_data = [['ID', 'Ticket Text', 'Status', 'Tone', 'Created', 'User']]
            for ticket in dept_tickets[:20]:  # Limit to 20 per department
                ticket_data.append([
                    str(ticket[0]),
                    ticket[1][:40] + '...' if len(ticket[1]) > 40 else ticket[1],
                    ticket[4],
                    ticket[3],
                    ticket[5][:10],
                    ticket[6] or 'System'
                ])
            
            ticket_table = Table(ticket_data, colWidths=[0.5*inch, 2.5*inch, 0.8*inch, 0.8*inch, 1*inch, 1*inch])
            ticket_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
            ]))
            
            elements.append(ticket_table)
            elements.append(Spacer(1, 20))
        
        if dept != list(sorted(dept_stats.keys()))[-1]:  # Don't add page break after last department
            elements.append(PageBreak())
    
    # ===== COMPLETE TICKET LOG =====
    elements.append(Paragraph("COMPLETE TICKET LOG", title_style))
    elements.append(Spacer(1, 20))
    
    if tickets:
        log_data = [['ID', 'Ticket Text', 'Department', 'Status', 'Tone', 'Created', 'User']]
        for ticket in tickets:
            log_data.append([
                str(ticket[0]),
                ticket[1][:50] + '...' if len(ticket[1]) > 50 else ticket[1],
                ticket[2] or 'Unassigned',
                ticket[4],
                ticket[3],
                ticket[5],
                ticket[6] or 'System'
            ])
        
        log_table = Table(log_data, colWidths=[0.4*inch, 2.5*inch, 1*inch, 0.7*inch, 0.7*inch, 1.2*inch, 1*inch])
        log_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkred),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
        ]))
        
        elements.append(log_table)
    else:
        elements.append(Paragraph("No tickets found for the selected criteria.", normal_style))
    
    conn.close()
    
    doc.build(elements)
    buffer.seek(0)
    
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"executive_report_{department}_{datetime.now().strftime('%Y%m%d')}.pdf",
        mimetype='application/pdf'
    )

# Generate CSV Report
def generate_csv_report(department, date_from, date_to):
    from io import StringIO
    import csv
    
    output = StringIO()
    writer = csv.writer(output)
    
    # Header
    writer.writerow(['ID', 'Ticket Text', 'Category', 'Tone', 'Status', 'Created At', 'User'])
    
    # Get data
    conn = get_db()
    cursor = conn.cursor()
    
    query = """
        SELECT t.id, t.ticket_text, t.category, t.tone, t.status, t.created_at, u.username
        FROM tickets t
        LEFT JOIN users u ON t.user_id = u.id
    """
    params = []
    
    if department != 'All':
        query += " WHERE t.category LIKE ?"
        params.append(f'%{department}%')
    
    if date_from and date_to:
        if 'WHERE' in query:
            query += " AND DATE(t.created_at) BETWEEN ? AND ?"
        else:
            query += " WHERE DATE(t.created_at) BETWEEN ? AND ?"
        params.extend([date_from, date_to])
    
    query += " ORDER BY t.created_at DESC"
    
    cursor.execute(query, params)
    tickets = cursor.fetchall()
    
    for ticket in tickets:
        writer.writerow([
            ticket[0],
            ticket[1],
            ticket[2],
            ticket[3],
            ticket[4],
            ticket[5],
            ticket[6] or 'System'
        ])
    
    conn.close()
    
    output.seek(0)
    buffer = BytesIO()
    buffer.write(output.getvalue().encode('utf-8'))
    buffer.seek(0)
    
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"ticket_report_{department}_{datetime.now().strftime('%Y%m%d')}.csv",
        mimetype='text/csv'
    )

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
        debug=os.environ.get("FLASK_DEBUG", "false").lower() in ("1", "true", "yes"),
    )
