"""
Contractor CRM & Database Management Engine for TDSC Platform
Handles:
1. SQLite schema migration & indexing for 54,000+ contractors
2. Batch import from Google Sheets (.xlsx / .csv) & custom user uploads
3. Fast paginated search, multi-district, entity type, caste, and category filtering
4. Full CRUD operations: Add, Edit, Delete, View
5. Styled Excel export (.xlsx)
6. Google Sheets live sync
7. Instant WhatsApp / Email tender suggestion linking
"""

import os
import re
import csv
import json
import sqlite3
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import urllib.request

DB_PATH = os.path.join(os.path.dirname(__file__), 'etender_subscription.db')
GOOGLE_SHEET_EXPORT_URL = 'https://docs.google.com/spreadsheets/d/1DWLXkr0ro699vd7NCaXG_AQKiFo4dyb84wBZX111-HE/export?format=xlsx'

DISTRICT_MAPPINGS = {
    'Araria': ['araria'],
    'Arwal': ['arwal'],
    'Aurangabad': ['aurangabad'],
    'Banka': ['banka'],
    'Begusarai': ['begusarai'],
    'Bhagalpur': ['bhagalpur'],
    'Bhojpur': ['bhojpur', 'ara'],
    'Buxar': ['buxar'],
    'Darbhanga': ['darbhanga'],
    'East Champaran': ['purba champaran', 'east champaran', 'motihari', 'purba champarar'],
    'Gaya': ['gaya'],
    'Gopalganj': ['gopalganj'],
    'Jamui': ['jamui'],
    'Jehanabad': ['jehanabad', 'jahanabad'],
    'Kaimur (Bhabua)': ['kaimur', 'bhabua'],
    'Katihar': ['katihar'],
    'Khagaria': ['khagaria', 'kh'],
    'Kishanganj': ['kishanganj'],
    'Lakhisarai': ['lakhisarai'],
    'Madhepura': ['madhepura'],
    'Madhubani': ['madhubani'],
    'Munger': ['munger'],
    'Muzaffarpur': ['muzaffarpur', 'muzafferpur'],
    'Nalanda': ['nalanda', 'biharsharif'],
    'Nawada': ['nawada'],
    'Patna': ['patna'],
    'Purnia': ['purnia', 'purnea'],
    'Rohtas': ['rohtas', 'sasaram'],
    'Saharsa': ['saharsa'],
    'Samastipur': ['samastipur'],
    'Saran': ['saran', 'chhapra'],
    'Sheikhpura': ['sheikhpura'],
    'Sheohar': ['sheohar'],
    'Sitamarhi': ['sitamarhi'],
    'Siwan': ['siwan'],
    'Supaul': ['supaul'],
    'Vaishali': ['vaishali', 'hajipur'],
    'West Champaran': ['pashchim champaran', 'west champaran', 'bettiah']
}

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    conn.row_factory = sqlite3.Row
    return conn

def init_contractors_schema():
    """Ensure contractors table has all Google Sheet fields and performant indexes."""
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS contractors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contractor_id TEXT UNIQUE NOT NULL,
        company_name TEXT NOT NULL,
        contact_person TEXT,
        mobile TEXT,
        email TEXT,
        district TEXT,
        category TEXT DEFAULT 'General Civil',
        work_class TEXT DEFAULT 'Class 1',
        budget_slab TEXT DEFAULT '50Lto2Cr',
        status TEXT DEFAULT 'active',
        source TEXT DEFAULT 'TDSC Office System',
        notes TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );
    """)
    
    cursor.execute("PRAGMA table_info(contractors);")
    existing_cols = set(r['name'] for r in cursor.fetchall())
    
    extra_cols = [
        ('application_id', 'TEXT'),
        ('director_name', 'TEXT'),
        ('dob_doa', 'TEXT'),
        ('address', 'TEXT'),
        ('cast', 'TEXT'),
        ('gender', 'TEXT'),
        ('entity_type', 'TEXT'),
        ('pan_no', 'TEXT'),
        ('calling_remarks', 'TEXT')
    ]
    
    for col, ctype in extra_cols:
        if col not in existing_cols:
            cursor.execute(f"ALTER TABLE contractors ADD COLUMN {col} {ctype};")
            
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_con_app_id ON contractors(application_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_con_mobile ON contractors(mobile);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_con_district ON contractors(district);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_con_company ON contractors(company_name);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_con_pan ON contractors(pan_no);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_con_entity ON contractors(entity_type);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_con_cast ON contractors([cast]);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_con_category ON contractors(category);")
    
    conn.commit()
    conn.close()

def clean_text(v):
    if v is None:
        return ""
    text = re.sub(r'[\r\n\t]+', ' ', str(v))
    return text.strip()

def clean_mobile(v):
    if not v:
        return ""
    digits = re.sub(r'[^0-9]', '', str(v))
    if len(digits) == 12 and digits.startswith('91'):
        digits = digits[2:]
    elif len(digits) > 10:
        digits = digits[-10:]
    return digits if len(digits) == 10 else digits

def normalize_district(raw: str) -> str:
    if not raw:
        return 'Other / State-wide'
    clean = raw.lower().strip()
    for std_name, kws in DISTRICT_MAPPINGS.items():
        for kw in kws:
            if kw in clean:
                return std_name
    return 'Other / State-wide'

def import_contractors_from_workbook(file_path: str, source_label: str = 'Excel Upload') -> dict:
    """Imports contractors from an Excel workbook (.xlsx / .csv) across all sheets with UPSERT."""
    init_contractors_schema()
    conn = get_db()
    cursor = conn.cursor()
    
    total_scanned = 0
    total_added = 0
    total_updated = 0
    districts_seen = set()
    now_iso = datetime.now().isoformat()
    
    # Check if CSV
    if file_path.lower().endswith('.csv'):
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            reader = csv.reader(f)
            headers = [clean_text(c).lower() for c in next(reader, [])]
            for r in reader:
                if not any(r):
                    continue
                total_scanned += 1
                app_id = clean_text(r[1]) if len(r) > 1 else ''
                comp_name = clean_text(r[2]) if len(r) > 2 else ''
                father = clean_text(r[3]) if len(r) > 3 else ''
                dob = clean_text(r[4]) if len(r) > 4 else ''
                addr = clean_text(r[5]) if len(r) > 5 else ''
                raw_dist = clean_text(r[6]) if len(r) > 6 else ''
                dist = normalize_district(raw_dist)
                cast = clean_text(r[7]).upper() if len(r) > 7 else ''
                gender = clean_text(r[8]).upper() if len(r) > 8 else ''
                email = clean_text(r[9]).lower() if len(r) > 9 else ''
                entity = clean_text(r[10]).title() if len(r) > 10 else 'Individual'
                mob = clean_mobile(r[11]) if len(r) > 11 else ''
                pan = clean_text(r[12]).upper() if len(r) > 12 else ''
                notes = clean_text(r[13]) if len(r) > 13 else ''
                
                if not comp_name and not mob:
                    continue
                if not comp_name:
                    comp_name = f"Contractor ({father or mob})"
                    
                districts_seen.add(dist)
                c_code = app_id if app_id else f"CON-{dist[:3].upper()}-{mob or total_scanned}"
                
                cursor.execute("SELECT id FROM contractors WHERE contractor_id = ? OR (application_id != '' AND application_id = ?);", (c_code, app_id))
                existing = cursor.fetchone()
                if existing:
                    cursor.execute("""
                    UPDATE contractors SET
                        company_name = ?, contact_person = ?, director_name = ?, mobile = COALESCE(NULLIF(?, ''), mobile),
                        email = COALESCE(NULLIF(?, ''), email), district = ?, dob_doa = ?, address = ?,
                        [cast] = ?, gender = ?, entity_type = ?, pan_no = ?, calling_remarks = ?, updated_at = ?
                    WHERE id = ?;
                    """, (comp_name, father or comp_name, father, mob, email, dist, dob, addr, cast, gender, entity, pan, notes, now_iso, existing['id']))
                    total_updated += 1
                else:
                    cursor.execute("""
                    INSERT INTO contractors (
                        contractor_id, application_id, company_name, contact_person, director_name,
                        mobile, email, district, dob_doa, address, [cast], gender, entity_type,
                        pan_no, calling_remarks, category, work_class, budget_slab, status, source, notes, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'General Civil', 'Class 1', '50Lto2Cr', 'active', ?, ?, ?, ?);
                    """, (c_code, app_id, comp_name, father or comp_name, father, mob, email, dist, dob, addr, cast, gender, entity, pan, notes, source_label, notes, now_iso, now_iso))
                    total_added += 1
                    
                if (total_added + total_updated) % 2000 == 0:
                    conn.commit()
    else: # Excel XLSX
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
                
            header_idx = -1
            for i, r in enumerate(rows[:6]):
                text = ' '.join(str(c or '').lower() for c in r)
                if 'application' in text or 'firm' in text or 'mobile' in text or 'applicant' in text:
                    header_idx = i
                    break
            if header_idx == -1:
                continue
                
            headers = [clean_text(c).lower() for c in rows[header_idx]]
            
            def get_col_idx(keywords):
                for idx, h in enumerate(headers):
                    for kw in keywords:
                        if kw in h:
                            return idx
                return -1

            app_id_idx = get_col_idx(['application', 'reg'])
            name_idx = get_col_idx(['applicant/firm', 'company name', 'firm', 'applicant name', 'name'])
            father_idx = get_col_idx(['father', 'director'])
            dob_idx = get_col_idx(['dob', 'doa'])
            addr_idx = get_col_idx(['address'])
            dist_idx = get_col_idx(['district'])
            cast_idx = get_col_idx(['cast', 'caste', 'category'])
            gender_idx = get_col_idx(['gender', 'sex'])
            email_idx = get_col_idx(['email'])
            entity_idx = get_col_idx(['entity'])
            mobile_idx = get_col_idx(['mobile', 'phone', 'contact'])
            pan_idx = get_col_idx(['pan'])
            notes_idx = 13 if len(headers) > 13 else -1

            for r in rows[header_idx+1:]:
                total_scanned += 1
                if not any(r):
                    continue
                    
                app_id = clean_text(r[app_id_idx]) if app_id_idx != -1 and app_id_idx < len(r) else ''
                comp_name = clean_text(r[name_idx]) if name_idx != -1 and name_idx < len(r) else ''
                mob = clean_mobile(r[mobile_idx]) if mobile_idx != -1 and mobile_idx < len(r) else ''
                
                if not comp_name and not mob:
                    continue
                    
                father = clean_text(r[father_idx]) if father_idx != -1 and father_idx < len(r) else ''
                dob = clean_text(r[dob_idx]) if dob_idx != -1 and dob_idx < len(r) else ''
                addr = clean_text(r[addr_idx]) if addr_idx != -1 and addr_idx < len(r) else ''
                
                raw_dist = clean_text(r[dist_idx]) if dist_idx != -1 and dist_idx < len(r) else ''
                if not raw_dist or len(raw_dist) > 30:
                    raw_dist = sheet_name
                district = normalize_district(raw_dist)
                districts_seen.add(district)

                cast = clean_text(r[cast_idx]).upper() if cast_idx != -1 and cast_idx < len(r) else ''
                gender = clean_text(r[gender_idx]).upper() if gender_idx != -1 and gender_idx < len(r) else ''
                email = clean_text(r[email_idx]).lower() if email_idx != -1 and email_idx < len(r) else ''
                entity = clean_text(r[entity_idx]).title() if entity_idx != -1 and entity_idx < len(r) else 'Individual'
                pan = clean_text(r[pan_idx]).upper() if pan_idx != -1 and pan_idx < len(r) else ''
                notes = clean_text(r[notes_idx]) if notes_idx != -1 and notes_idx < len(r) else ''

                if not comp_name:
                    comp_name = f"Contractor ({father or mob})"
                contact_person = father or comp_name

                c_code = app_id if app_id else f"CON-{district[:3].upper()}-{mob or total_scanned}"

                cursor.execute("SELECT id FROM contractors WHERE contractor_id = ? OR (application_id != '' AND application_id = ?);", (c_code, app_id))
                existing = cursor.fetchone()

                if existing:
                    cid = existing['id']
                    cursor.execute("""
                    UPDATE contractors SET
                        company_name = ?, contact_person = ?, director_name = ?, mobile = COALESCE(NULLIF(?, ''), mobile),
                        email = COALESCE(NULLIF(?, ''), email), district = ?, dob_doa = ?, address = ?,
                        [cast] = ?, gender = ?, entity_type = ?, pan_no = ?, calling_remarks = ?, updated_at = ?
                    WHERE id = ?;
                    """, (comp_name, contact_person, father, mob, email, district, dob, addr,
                          cast, gender, entity, pan, notes, now_iso, cid))
                    total_updated += 1
                else:
                    cursor.execute("""
                    INSERT INTO contractors (
                        contractor_id, application_id, company_name, contact_person, director_name,
                        mobile, email, district, dob_doa, address, [cast], gender, entity_type,
                        pan_no, calling_remarks, category, work_class, budget_slab, status, source, notes, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'General Civil', 'Class 1', '50Lto2Cr', 'active', ?, ?, ?, ?);
                    """, (c_code, app_id, comp_name, contact_person, father, mob, email, district,
                          dob, addr, cast, gender, entity, pan, notes, source_label, notes, now_iso, now_iso))
                    total_added += 1

                if (total_added + total_updated) % 2000 == 0:
                    conn.commit()

    conn.commit()
    conn.close()
    return {
        "success": True,
        "scanned": total_scanned,
        "added": total_added,
        "updated": total_updated,
        "total": total_added + total_updated,
        "districts_count": len(districts_seen),
        "districts": sorted(list(districts_seen))
    }

def get_contractors_list(search: str = "", district: str = "", cast: str = "", entity_type: str = "",
                         category: str = "", page: int = 1, limit: int = 50, sort: str = "name_asc") -> dict:
    """Fast indexed pagination and query for contractor database."""
    init_contractors_schema()
    conn = get_db()
    cursor = conn.cursor()
    
    where_clauses = ["1=1"]
    params = []
    
    if search:
        search_kw = f"%{search.strip().lower()}%"
        where_clauses.append("""(
            LOWER(company_name) LIKE ? OR
            LOWER(contact_person) LIKE ? OR
            LOWER(director_name) LIKE ? OR
            mobile LIKE ? OR
            LOWER(application_id) LIKE ? OR
            LOWER(pan_no) LIKE ? OR
            LOWER(district) LIKE ? OR
            LOWER(address) LIKE ? OR
            LOWER(calling_remarks) LIKE ?
        )""")
        params.extend([search_kw] * 9)
        
    if district and district.lower() != 'all':
        where_clauses.append("LOWER(district) = ?")
        params.append(district.strip().lower())
        
    if cast and cast.lower() != 'all':
        where_clauses.append("UPPER([cast]) = ?")
        params.append(cast.strip().upper())
        
    if entity_type and entity_type.lower() != 'all':
        where_clauses.append("LOWER(entity_type) = ?")
        params.append(entity_type.strip().lower())

    if category and category.lower() != 'all':
        where_clauses.append("LOWER(category) = ?")
        params.append(category.strip().lower())
        
    where_sql = " AND ".join(where_clauses)
    
    cursor.execute(f"SELECT COUNT(*) as cnt FROM contractors WHERE {where_sql};", params)
    total_matching = cursor.fetchone()['cnt']
    
    sort_sql = "ORDER BY company_name ASC"
    if sort == "name_desc":
        sort_sql = "ORDER BY company_name DESC"
    elif sort == "district_asc":
        sort_sql = "ORDER BY district ASC, company_name ASC"
    elif sort == "app_id_asc":
        sort_sql = "ORDER BY application_id ASC"
    elif sort == "recent":
        sort_sql = "ORDER BY id DESC"
        
    offset = max(0, (page - 1) * limit)
    cursor.execute(f"""
    SELECT id, contractor_id, application_id, company_name, contact_person, director_name,
           mobile, email, district, dob_doa, address, [cast], gender, entity_type,
           pan_no, calling_remarks, category, work_class, budget_slab, status, source, notes, updated_at
    FROM contractors
    WHERE {where_sql}
    {sort_sql}
    LIMIT ? OFFSET ?;
    """, params + [limit, offset])
    
    rows = [dict(r) for r in cursor.fetchall()]
    
    cursor.execute("SELECT COUNT(DISTINCT district) as dist_cnt, COUNT(*) as total_db, COUNT(CASE WHEN mobile != '' AND mobile IS NOT NULL THEN 1 END) as mob_cnt FROM contractors;")
    stat_row = cursor.fetchone()
    
    cursor.execute("SELECT DISTINCT district FROM contractors WHERE district != '' AND district IS NOT NULL ORDER BY district ASC;")
    districts = [r['district'] for r in cursor.fetchall() if r['district']]

    cursor.execute("SELECT DISTINCT category FROM contractors WHERE category != '' AND category IS NOT NULL ORDER BY category ASC;")
    categories = [r['category'] for r in cursor.fetchall() if r['category']]

    cursor.execute("SELECT DISTINCT entity_type FROM contractors WHERE entity_type != '' AND entity_type IS NOT NULL ORDER BY entity_type ASC;")
    entity_types = [r['entity_type'] for r in cursor.fetchall() if r['entity_type']]

    cursor.execute("SELECT DISTINCT [cast] FROM contractors WHERE [cast] != '' AND [cast] IS NOT NULL ORDER BY [cast] ASC;")
    castes = [r['cast'] for r in cursor.fetchall() if r['cast']]
    
    conn.close()
    return {
        "success": True,
        "total": total_matching,
        "page": page,
        "limit": limit,
        "contractors": rows,
        "stats": {
            "total_contractors": stat_row['total_db'],
            "verified_mobiles": stat_row['mob_cnt'],
            "districts_count": stat_row['dist_cnt']
        },
        "districts": districts,
        "categories": categories,
        "entity_types": entity_types,
        "castes": castes
    }

def add_contractor_entry(data: dict) -> dict:
    """Add a new contractor manually via web form."""
    init_contractors_schema()
    conn = get_db()
    cursor = conn.cursor()
    
    now_iso = datetime.now().isoformat()
    app_id = clean_text(data.get('application_id') or data.get('contractor_id'))
    comp_name = clean_text(data.get('company_name') or data.get('companyName'))
    if not comp_name:
        raise ValueError("Contractor or Firm Name is required.")
        
    mob = clean_mobile(data.get('mobile'))
    c_code = app_id if app_id else f"CON-MANUAL-{mob or int(datetime.now().timestamp())}"
    dist = normalize_district(clean_text(data.get('district')))
    
    cursor.execute("""
    INSERT INTO contractors (
        contractor_id, application_id, company_name, contact_person, director_name,
        mobile, email, district, dob_doa, address, [cast], gender, entity_type,
        pan_no, calling_remarks, category, work_class, budget_slab, status, source, notes, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', 'Manual Web Entry', ?, ?, ?);
    """, (
        c_code,
        app_id,
        comp_name,
        clean_text(data.get('contact_person') or data.get('contactPerson')) or comp_name,
        clean_text(data.get('director_name') or data.get('directorName')),
        mob,
        clean_text(data.get('email')).lower(),
        dist,
        clean_text(data.get('dob_doa') or data.get('dobDoa')),
        clean_text(data.get('address')),
        clean_text(data.get('cast')).upper(),
        clean_text(data.get('gender')).upper() or 'MALE',
        clean_text(data.get('entity_type') or data.get('entityType')).title() or 'Individual',
        clean_text(data.get('pan_no') or data.get('panNo')).upper(),
        clean_text(data.get('calling_remarks') or data.get('callingRemarks')),
        clean_text(data.get('category')) or 'General Civil',
        clean_text(data.get('work_class') or data.get('workClass')) or 'Class 1',
        clean_text(data.get('budget_slab') or data.get('budgetSlab')) or '50Lto2Cr',
        clean_text(data.get('notes')),
        now_iso, now_iso
    ))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return {"success": True, "id": new_id, "contractor_id": c_code, "message": f"Contractor '{comp_name}' added successfully!"}

def update_contractor_entry(cid: int, data: dict) -> dict:
    """Update contractor details."""
    init_contractors_schema()
    conn = get_db()
    cursor = conn.cursor()
    
    now_iso = datetime.now().isoformat()
    dist = normalize_district(clean_text(data.get('district')))
    cursor.execute("""
    UPDATE contractors SET
        application_id = COALESCE(NULLIF(?, ''), application_id),
        company_name = COALESCE(NULLIF(?, ''), company_name),
        contact_person = COALESCE(NULLIF(?, ''), contact_person),
        director_name = ?,
        mobile = COALESCE(NULLIF(?, ''), mobile),
        email = ?,
        district = ?,
        dob_doa = ?,
        address = ?,
        [cast] = ?,
        gender = ?,
        entity_type = ?,
        pan_no = ?,
        calling_remarks = ?,
        category = COALESCE(NULLIF(?, ''), category),
        work_class = COALESCE(NULLIF(?, ''), work_class),
        budget_slab = COALESCE(NULLIF(?, ''), budget_slab),
        notes = ?,
        updated_at = ?
    WHERE id = ?;
    """, (
        clean_text(data.get('application_id') or data.get('contractor_id')),
        clean_text(data.get('company_name') or data.get('companyName')),
        clean_text(data.get('contact_person') or data.get('contactPerson')),
        clean_text(data.get('director_name') or data.get('directorName')),
        clean_mobile(data.get('mobile')),
        clean_text(data.get('email')).lower(),
        dist,
        clean_text(data.get('dob_doa') or data.get('dobDoa')),
        clean_text(data.get('address')),
        clean_text(data.get('cast')).upper(),
        clean_text(data.get('gender')).upper(),
        clean_text(data.get('entity_type') or data.get('entityType')).title(),
        clean_text(data.get('pan_no') or data.get('panNo')).upper(),
        clean_text(data.get('calling_remarks') or data.get('callingRemarks')),
        clean_text(data.get('category')),
        clean_text(data.get('work_class') or data.get('workClass')),
        clean_text(data.get('budget_slab') or data.get('budgetSlab')),
        clean_text(data.get('notes')),
        now_iso, cid
    ))
    conn.commit()
    conn.close()
    return {"success": True, "message": "Contractor updated successfully!"}

def delete_contractor_entry(cid: int) -> dict:
    """Delete a contractor by primary key id."""
    init_contractors_schema()
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM contractors WHERE id = ?;", (cid,))
    conn.commit()
    conn.close()
    return {"success": True, "message": "Contractor deleted successfully!"}

def generate_contractors_excel_export(filters: dict = None, output_path: str = None) -> str:
    """Generate styled Excel workbook matching the Google Sheets schema."""
    init_contractors_schema()
    if output_path is None:
        output_path = os.path.join(os.path.dirname(__file__), "Contractors_Master_Database.xlsx")
        
    conn = get_db()
    cursor = conn.cursor()
    
    where_sql = "1=1"
    params = []
    if filters and filters.get('district') and filters.get('district').lower() != 'all':
        where_sql += " AND LOWER(district) = ?"
        params.append(filters.get('district').strip().lower())
        
    cursor.execute(f"""
    SELECT id, application_id, company_name, director_name, dob_doa, address, district,
           [cast], gender, email, entity_type, mobile, pan_no, calling_remarks
    FROM contractors
    WHERE {where_sql}
    ORDER BY district ASC, company_name ASC
    LIMIT 35000;
    """, params)
    rows = cursor.fetchall()
    conn.close()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contractor Master Registry"
    ws.views.sheetView[0].showGridLines = True
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    thin_border = Border(
        left=Side(style='thin', color='CBD5E0'), right=Side(style='thin', color='CBD5E0'),
        top=Side(style='thin', color='CBD5E0'), bottom=Side(style='thin', color='CBD5E0')
    )
    
    headers = [
        "SL. No.", "Application ID", "Applicant / Firm / Company Name",
        "Father / Director Name", "DoB / DoA", "Address", "District",
        "Cast", "Gender", "Email ID", "Entity Type", "Mobile No", "PAN No", "Calling Remarks"
    ]
    
    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    
    for idx, r in enumerate(rows, 1):
        row_vals = [
            idx,
            r['application_id'] or f"CON-{r['id']}",
            r['company_name'] or '',
            r['director_name'] or '',
            r['dob_doa'] or '',
            r['address'] or '',
            r['district'] or '',
            r['cast'] or '',
            r['gender'] or '',
            r['email'] or '',
            r['entity_type'] or '',
            r['mobile'] or '',
            r['pan_no'] or '',
            r['calling_remarks'] or ''
        ]
        ws.append(row_vals)
        r_idx = idx + 1
        ws.row_dimensions[r_idx].height = 20
        for col_idx in range(1, len(row_vals) + 1):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [1, 2, 7, 12, 13]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
    wb.save(output_path)
    return output_path

def sync_from_google_sheet(sheet_url: str = None) -> dict:
    """Download Google Sheet XLSX export and sync into SQLite."""
    url = sheet_url if sheet_url else GOOGLE_SHEET_EXPORT_URL
    temp_path = os.path.join(os.path.dirname(__file__), 'temp_google_sheet_sync.xlsx')
    
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=45) as resp:
        content = resp.read()
        with open(temp_path, 'wb') as f:
            f.write(content)
            
    res = import_contractors_from_workbook(temp_path, source_label='Google Sheet Sync')
    try:
        if os.path.exists(temp_path):
            os.remove(temp_path)
    except Exception:
        pass
    return res
