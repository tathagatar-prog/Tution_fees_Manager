import streamlit as st
import sqlite3, pandas as pd, os, base64, calendar, io
from datetime import date, datetime
from pathlib import Path
from contextlib import contextmanager

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "tuition_centre.db"
LOGO_PATH = APP_DIR / "logo.png"
EXCEL_PATH = APP_DIR / "Tution Payment Status.xlsx"

CENTRE = "Tutoring Computer Science & Application"
ADDRESS = "496/2, Amarabati Rd, Choto Bazar, Nabapally, Barasat, Kolkata, West Bengal 700126"
PHONE = "070035 85265"

st.set_page_config(page_title="Tuition Centre Manager", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
html, body, [class*="css"] {font-family:'DM Sans',sans-serif;}
.block-container {padding-top:1.1rem; padding-bottom:2rem; max-width:1400px;}
[data-testid="stSidebar"] {background:linear-gradient(180deg,#10294b 0%,#173d69 100%);}
[data-testid="stSidebar"] * {color:white !important;}
[data-testid="stMetric"] {background:#f5f8fc; padding:16px 18px; border-radius:14px; border:1px solid #e5edf7;}
div.stButton > button {border-radius:10px; font-weight:600; min-height:2.6rem;}
div[data-testid="stForm"] {border:1px solid #e7edf5; border-radius:14px; padding:18px;}
[data-testid="stDataFrame"] {border-radius:12px; overflow:hidden;}
h1,h2,h3 {letter-spacing:-.4px;}
.small-muted {color:#64748b;font-size:13px;}
@media(max-width:700px){.block-container{padding-left: .8rem;padding-right:.8rem;} h1{font-size:1.7rem!important;} }
</style>
""", unsafe_allow_html=True)

@contextmanager
def db():
    con=sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory=sqlite3.Row
    try:
        yield con
        con.commit()
    finally: con.close()

def init_db():
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS students(
          id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, category TEXT NOT NULL,
          course TEXT DEFAULT '', monthly_fee REAL DEFAULT 0, phone TEXT DEFAULT '',
          guardian TEXT DEFAULT '', join_date TEXT DEFAULT '', active INTEGER DEFAULT 1,
          created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS payments(
          id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER NOT NULL,
          fee_month TEXT NOT NULL, amount REAL NOT NULL, payment_date TEXT NOT NULL,
          mode TEXT DEFAULT 'Cash', receipt_no TEXT UNIQUE, note TEXT DEFAULT '',
          FOREIGN KEY(student_id) REFERENCES students(id));
        CREATE INDEX IF NOT EXISTS idx_payments_month ON payments(fee_month);
        CREATE TABLE IF NOT EXISTS expenses(
          id INTEGER PRIMARY KEY AUTOINCREMENT, expense_date TEXT NOT NULL,
          category TEXT NOT NULL, amount REAL NOT NULL, description TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
        """)
        n=c.execute("SELECT COUNT(*) FROM students").fetchone()[0]
        imported=c.execute("SELECT value FROM settings WHERE key='initial_import_done'").fetchone()
    if n==0 and not imported and EXCEL_PATH.exists():
        import_excel()
    # default receipt counter is derived from existing records

def category_from_sheet(sheet, subject):
    s=(str(subject or "")).upper()
    if "BTECH" in sheet.upper(): return "B.Tech"
    if "ITI" in s or "CST" in s: return "ITI/CST"
    if "BSC" in s: return "B.Sc"
    if "MCA" in s or "BCA" in s: return "BCA/MCA"
    if "XI" in s or "XII" in s or "CLASS" in sheet.upper(): return "Other/XI/XII"
    return "Other/XI/XII"

def clean(v):
    if pd.isna(v): return ""
    return str(v).replace("\\n"," ").strip()

def import_excel():
    try:
        xls=pd.ExcelFile(EXCEL_PATH)
        with db() as c:
            for sheet in xls.sheet_names:
                raw=pd.read_excel(EXCEL_PATH, sheet_name=sheet, header=None)
                if len(raw)<4: continue
                # The workbook has a title row, a column-label row and a month-header row.
                headers=[clean(x).upper() for x in raw.iloc[2].tolist()]
                for ridx in range(3,len(raw)):
                    row=raw.iloc[ridx]
                    name=""
                    subject=""
                    amount=0
                    if sheet=="IX to XII":
                        name=clean(row.iloc[0]); subject=clean(row.iloc[1]); amount=row.iloc[2] if len(row)>2 else 0
                        month_start=3
                    elif sheet=="BCA-MCA-BSC":
                        name=clean(row.iloc[1]); subject=clean(row.iloc[2]); amount=row.iloc[3] if len(row)>3 else 0
                        month_start=4
                    else:
                        name=clean(row.iloc[0]); subject=clean(row.iloc[1]); amount=row.iloc[2] if len(row)>2 else 0
                        month_start=3
                    if not name or name.upper() in ("STUDENT NAME","NAN"): continue
                    try: fee=float(amount) if not pd.isna(amount) else 0.0
                    except: fee=0.0
                    cat=category_from_sheet(sheet,subject)
                    course=subject
                    cur=c.execute("INSERT INTO students(name,category,course,monthly_fee,active) VALUES(?,?,?,?,1)",
                                  (name,cat,course,fee))
                    sid=cur.lastrowid
                    # Each populated date cell is retained as a payment for the month named by its Excel column.
                    for col in range(month_start,len(row)):
                        val=row.iloc[col]
                        if pd.isna(val) or clean(val)=="": continue
                        try:
                            dt=pd.to_datetime(val, errors="coerce")
                            if pd.isna(dt): continue
                        except Exception:
                            continue
                        head=headers[col] if col<len(headers) else ""
                        month_num=list(calendar.month_name).index(head.title()) if head.title() in list(calendar.month_name) else dt.month
                        fee_month=f"{dt.year:04d}-{month_num:02d}"
                        # A populated date is proof of a recorded payment; unknown fees are flagged in note.
                        amt=fee if fee>0 else 0.0
                        note="Imported from original Excel register"
                        if fee<=0: note += " — fee amount missing; please verify"
                        receipt=f"IMP-{sid:03d}-{dt.strftime('%Y%m')}-{col:02d}"
                        c.execute("""INSERT OR IGNORE INTO payments
                          (student_id,fee_month,amount,payment_date,mode,receipt_no,note)
                          VALUES(?,?,?,?,?,?,?)""",
                          (sid,fee_month,amt,dt.strftime("%Y-%m-%d"),"Imported",receipt,note))
                    # No extra data guessed: contact and joining date remain editable.
            c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('initial_import_done','1')")
    except Exception as e:
        st.error(f"Could not import the Excel workbook: {e}")

def qdf(sql,args=()):
    with db() as c: return pd.read_sql_query(sql,c,params=args)

def money(v): return "₹{:,.2f}".format(float(v or 0)).replace(".00","")

def months():
    return [(f"{y}-{m:02d}",f"{calendar.month_name[m]} {y}") for y in range(date.today().year-3,date.today().year+3) for m in range(1,13)]

def receipt_html(payment_id):
    with db() as c:
        p=c.execute("""SELECT p.*,s.name,s.category,s.course,s.phone FROM payments p
                       JOIN students s ON s.id=p.student_id WHERE p.id=?""",(payment_id,)).fetchone()
    if not p: return ""
    logo=""
    if LOGO_PATH.exists():
        logo="data:image/png;base64,"+base64.b64encode(LOGO_PATH.read_bytes()).decode()
    month_label=datetime.strptime(p["fee_month"],"%Y-%m").strftime("%B %Y")
    return f"""
    <html><head><meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    @page{{size: A5 portrait;margin:8mm}} body{{font-family:Arial,sans-serif;color:#172033;margin:0}}
    .receipt{{width:100%;max-width:620px;margin:auto;border:1.5px solid #1e3a5f;padding:16px 18px;box-sizing:border-box}}
    .head{{display:flex;align-items:center;gap:12px;border-bottom:1px solid #9aa9bc;padding-bottom:10px}}
    .logo{{width:75px;height:65px;object-fit:contain}} h2{{font-size:17px;margin:0 0 4px;color:#173b68}}
    .sub{{font-size:10px;line-height:1.5;color:#475569}} h3{{text-align:center;font-size:15px;letter-spacing:1px;margin:13px 0}}
    .line{{display:flex;justify-content:space-between;gap:10px;padding:7px 0;border-bottom:1px dashed #cbd5e1;font-size:12px}}
    .amount{{font-size:20px;font-weight:bold;color:#12653b}} .paid{{display:inline-block;background:#e6f6ed;color:#12653b;padding:5px 10px;border-radius:20px;font-size:10px;font-weight:bold}}
    .foot{{font-size:10px;color:#475569;text-align:center;margin-top:16px}} .sign{{text-align:right;font-size:10px;margin-top:25px}}
    @media print{{body{{-webkit-print-color-adjust:exact;print-color-adjust:exact}} .no-print{{display:none}}}}
    </style></head><body><div class="receipt">
      <div class="head">{f'<img class="logo" src="{logo}">' if logo else ''}
      <div><h2>{CENTRE}</h2><div class="sub">{ADDRESS}<br>Phone: {PHONE}</div></div></div>
      <h3>FEE PAYMENT RECEIPT</h3>
      <div class="line"><span>Receipt No: <b>{p['receipt_no']}</b></span><span>Date: <b>{p['payment_date']}</b></span></div>
      <div class="line"><span>Student</span><b>{p['name']}</b></div>
      <div class="line"><span>Course / Category</span><b>{p['course'] or p['category']}</b></div>
      <div class="line"><span>Fee month</span><b>{month_label}</b></div>
      <div class="line"><span>Payment mode</span><b>{p['mode']}</b></div>
      <div class="line"><span>Amount received</span><span class="amount">{money(p['amount'])}</span></div>
      <p style="text-align:center"><span class="paid">PAYMENT RECEIVED</span></p>
      <div class="foot">Thank you for choosing {CENTRE}.</div>
      <div class="sign">Authorised Signature<br><br>______________________</div>
    </div><script>window.onload=()=>{{setTimeout(()=>window.print(),350)}};</script></body></html>"""

def next_receipt():
    with db() as c:
        n=c.execute("SELECT COUNT(*) FROM payments WHERE receipt_no LIKE 'TCSA-%'").fetchone()[0]+1
    return f"TCSA-{n:05d}"

def logo_sidebar():
    if LOGO_PATH.exists():
        st.sidebar.image(str(LOGO_PATH),width=150)
    st.sidebar.markdown("### Tuition Centre Manager")
    st.sidebar.caption("Student • Fees • Receipts • Reports")

init_db()
logo_sidebar()
page=st.sidebar.radio("MAIN MENU",["Dashboard","Students","Fee Collection","Monthly Fee Status","Income & Expenses","Reports","Backup & Data"])
st.sidebar.divider()
st.sidebar.caption("Local SQLite database • Keep regular backups")

if page=="Dashboard":
    st.title("Good day! 👋")
    st.caption("Your tuition centre at a glance")
    today=date.today()
    month=st.selectbox("Dashboard month",[x[0] for x in months()],index=[x[0] for x in months()].index(today.strftime("%Y-%m")),format_func=lambda x:datetime.strptime(x,"%Y-%m").strftime("%B %Y"))
    total=qdf("SELECT COUNT(*) n FROM students WHERE active=1").iloc[0,0]
    expected=qdf("SELECT COALESCE(SUM(monthly_fee),0) n FROM students WHERE active=1").iloc[0,0]
    paid=qdf("SELECT COALESCE(SUM(amount),0) n FROM payments WHERE fee_month=?",(month,)).iloc[0,0]
    paid_students=qdf("SELECT COUNT(DISTINCT student_id) n FROM payments WHERE fee_month=?",(month,)).iloc[0,0]
    expenses=qdf("SELECT COALESCE(SUM(amount),0) n FROM expenses WHERE substr(expense_date,1,7)=?",(month,)).iloc[0,0]
    a,b,c,d=st.columns(4)
    a.metric("Active students",total)
    b.metric("Fees received",money(paid))
    c.metric("Students paid",f"{paid_students} / {total}")
    d.metric("Expenses this month",money(expenses))
    st.subheader("Quick overview")
    e,f=st.columns(2)
    e.metric("Expected monthly fees",money(expected))
    f.metric("Net recorded income",money(paid-expenses))
    st.info("Tip: Open **Monthly Fee Status** to see paid, partial and unpaid students. Record each payment in Fee Collection to generate a receipt.")
    recent=qdf("""SELECT p.payment_date,s.name,s.category,p.fee_month,p.amount,p.mode,p.receipt_no
                  FROM payments p JOIN students s ON s.id=p.student_id ORDER BY p.id DESC LIMIT 8""")
    st.subheader("Recent payments")
    if len(recent): st.dataframe(recent, use_container_width=True, hide_index=True)
    else: st.info("No payments recorded yet.")

elif page=="Students":
    st.title("Student Management")
    st.caption("Add, update, search, deactivate or permanently remove student records.")
    with st.expander("➕ Add a new student",expanded=False):
        with st.form("add_student",clear_on_submit=True):
            c1,c2=st.columns(2)
            name=c1.text_input("Student name *")
            category=c2.selectbox("Tuition category",["BCA/MCA","ITI/CST","B.Tech","B.Sc","Other/XI/XII"])
            course=c1.text_input("Course / subject details")
            fee=c2.number_input("Monthly fee (₹)",min_value=0.0,step=100.0)
            phone=c1.text_input("Student phone")
            guardian=c2.text_input("Parent / guardian name")
            joined=c1.date_input("Joining date",value=date.today())
            ok=st.form_submit_button("Save student",type="primary",use_container_width=True)
            if ok:
                if not name.strip(): st.error("Please enter the student's name.")
                else:
                    with db() as con: con.execute("INSERT INTO students(name,category,course,monthly_fee,phone,guardian,join_date) VALUES(?,?,?,?,?,?,?)",(name.strip(),category,course,fee,phone,guardian,str(joined)))
                    st.success("Student added successfully."); st.rerun()
    search=st.text_input("🔎 Search student by name, course or phone")
    cat=st.selectbox("Filter category",["All","BCA/MCA","ITI/CST","B.Tech","B.Sc","Other/XI/XII"])
    active_filter=st.radio("Student status",["Active","Inactive","All"],horizontal=True)
    sql="SELECT * FROM students WHERE 1=1"; args=[]
    if search: sql+=" AND (name LIKE ? OR course LIKE ? OR phone LIKE ?)"; args += [f"%{search}%"]*3
    if cat!="All": sql+=" AND category=?"; args.append(cat)
    if active_filter=="Active": sql+=" AND active=1"
    elif active_filter=="Inactive": sql+=" AND active=0"
    sql+=" ORDER BY name"
    students=qdf(sql,args)
    st.write(f"**{len(students)} students found**")
    st.dataframe(students[["id","name","category","course","monthly_fee","phone","active"]].rename(columns={"id":"ID","name":"Name","category":"Category","course":"Course","monthly_fee":"Monthly fee","phone":"Phone","active":"Active (1=yes)"}),use_container_width=True,hide_index=True)
    if len(students):
        chosen=st.selectbox("Select student to edit / manage",students["id"].tolist(),format_func=lambda i: f"{students.loc[students.id==i,'name'].iloc[0]} (ID {i})")
        rec=students[students.id==chosen].iloc[0]
        with st.form("edit_student"):
            x1,x2=st.columns(2)
            ename=x1.text_input("Name",rec["name"])
            ecat=x2.selectbox("Category",["BCA/MCA","ITI/CST","B.Tech","B.Sc","Other/XI/XII"],index=["BCA/MCA","ITI/CST","B.Tech","B.Sc","Other/XI/XII"].index(rec["category"]) if rec["category"] in ["BCA/MCA","ITI/CST","B.Tech","B.Sc","Other/XI/XII"] else 0)
            ecourse=x1.text_input("Course / subjects",rec["course"] or "")
            efee=x2.number_input("Monthly fee (₹)",min_value=0.0,value=float(rec["monthly_fee"] or 0),step=100.0)
            ephone=x1.text_input("Phone",rec["phone"] or "")
            eguardian=x2.text_input("Guardian",rec["guardian"] or "")
            eactive=x1.checkbox("Active student",value=bool(rec["active"]))
            save=st.form_submit_button("Update student",type="primary")
            if save:
                with db() as con: con.execute("UPDATE students SET name=?,category=?,course=?,monthly_fee=?,phone=?,guardian=?,active=? WHERE id=?",(ename,ecat,ecourse,efee,ephone,eguardian,int(eactive),int(chosen)))
                st.success("Student updated."); st.rerun()
        with st.expander("⚠️ Student removal options"):
            st.warning("Recommended: deactivate a student to preserve all historical payment records. Permanent deletion also deletes that student's payment history.")
            if st.button("Deactivate this student"):
                with db() as con: con.execute("UPDATE students SET active=0 WHERE id=?",(int(chosen),))
                st.success("Student deactivated."); st.rerun()
            confirm=st.checkbox("I understand permanent deletion cannot be undone.")
            if st.button("Permanently delete student and payments",disabled=not confirm):
                with db() as con:
                    con.execute("DELETE FROM payments WHERE student_id=?",(int(chosen),))
                    con.execute("DELETE FROM students WHERE id=?",(int(chosen),))
                st.success("Student and payment history deleted."); st.rerun()

elif page=="Fee Collection":
    st.title("Record Fee Payment")
    st.caption("Record full or partial payments and create a printable receipt.")
    active=qdf("SELECT id,name,category,course,monthly_fee FROM students WHERE active=1 ORDER BY name")
    if active.empty: st.warning("Add students first.")
    else:
        selected=st.selectbox("Choose student",active.id.tolist(),format_func=lambda i:f"{active.loc[active.id==i,'name'].iloc[0]} — {active.loc[active.id==i,'category'].iloc[0]}")
        stu=active[active.id==selected].iloc[0]
        x,y=st.columns(2)
        fee_month=x.selectbox("Fee month",[z[0] for z in months()],index=[z[0] for z in months()].index(date.today().strftime("%Y-%m")),format_func=lambda v:datetime.strptime(v,"%Y-%m").strftime("%B %Y"))
        due=float(stu.monthly_fee or 0)
        already=qdf("SELECT COALESCE(SUM(amount),0) total FROM payments WHERE student_id=? AND fee_month=?",(int(selected),fee_month)).iloc[0,0]
        remaining=max(0,due-float(already))
        y.metric("Monthly fee",money(due)); x.metric("Already paid",money(already))
        st.metric("Calculated remaining balance",money(remaining))
        with st.form("payment_form"):
            a,b=st.columns(2)
            amount=a.number_input("Amount received (₹)",min_value=1.0,value=float(remaining if remaining>0 else max(due,1)),step=100.0)
            pdate=b.date_input("Payment date",value=date.today())
            mode=a.selectbox("Payment mode",["Cash","UPI","Bank Transfer","Other"])
            note=b.text_input("Note (optional)")
            submit=st.form_submit_button("💳 Save payment & generate receipt",type="primary",use_container_width=True)
            if submit:
                receipt=next_receipt()
                with db() as con:
                    con.execute("INSERT INTO payments(student_id,fee_month,amount,payment_date,mode,receipt_no,note) VALUES(?,?,?,?,?,?,?)",(int(selected),fee_month,float(amount),str(pdate),mode,receipt,note))
                    pid=con.execute("SELECT last_insert_rowid()").fetchone()[0]
                st.session_state["last_payment_id"]=pid
                st.success(f"Payment saved. Receipt: {receipt}")
                st.rerun()
        if st.session_state.get("last_payment_id"):
            pid=st.session_state["last_payment_id"]
            html=receipt_html(pid)
            st.markdown("### Receipt ready")
            st.components.v1.html(html,height=510,scrolling=True)

elif page=="Monthly Fee Status":
    st.title("Month-wise Payment Status")
    st.caption("Payment status is calculated from payments recorded against each student and month.")
    opts=[z[0] for z in months()]
    selected_month=st.selectbox("Select month",opts,index=opts.index(date.today().strftime("%Y-%m")),format_func=lambda v:datetime.strptime(v,"%Y-%m").strftime("%B %Y"))
    cat=st.selectbox("Course category",["All","BCA/MCA","ITI/CST","B.Tech","B.Sc","Other/XI/XII"],key="statuscat")
    stu=qdf("SELECT * FROM students WHERE active=1"+(" AND category=?" if cat!="All" else "")+" ORDER BY name",(cat,) if cat!="All" else ())
    paid=qdf("SELECT student_id,SUM(amount) paid,COUNT(*) n FROM payments WHERE fee_month=? GROUP BY student_id",(selected_month,))
    paid_map=dict(zip(paid.student_id,paid.paid))
    count_map=dict(zip(paid.student_id,paid.n))
    rows=[]
    for _,s in stu.iterrows():
        got=float(paid_map.get(s.id,0) or 0); fee=float(s.monthly_fee or 0)
        count=int(count_map.get(s.id,0) or 0)
        status=("Verify amount" if count>0 and got<=0 else ("Unpaid" if count==0 else ("Paid" if fee>0 and got>=fee else ("Paid" if fee==0 else "Partial"))))
        rows.append({"Student":s["name"],"Category":s["category"],"Course":s["course"],"Fee (₹)":fee,"Paid (₹)":got,"Balance (₹)":max(0,fee-got),"Status":status})
    result=pd.DataFrame(rows)
    if len(result):
        a,b,c=st.columns(3)
        a.metric("Paid",int((result.Status=="Paid").sum()))
        b.metric("Partial",int((result.Status=="Partial").sum()))
        c.metric("Unpaid",int((result.Status=="Unpaid").sum()))
        if (result.Status=="Verify amount").any(): st.warning("Some imported payments have a recorded date but no fee amount in the original Excel file. These are marked Verify amount, not Unpaid.")
        filt=st.multiselect("Show statuses",["Paid","Partial","Unpaid","Verify amount"],default=["Paid","Partial","Unpaid","Verify amount"])
        view=result[result.Status.isin(filt)]
        st.dataframe(view,use_container_width=True,hide_index=True)
        st.download_button("⬇️ Download this month's status (CSV)",view.to_csv(index=False).encode("utf-8-sig"),file_name=f"fee_status_{selected_month}.csv",mime="text/csv")
    else: st.info("No active students in this category.")

elif page=="Income & Expenses":
    st.title("Income & Expense Register")
    st.caption("Track all six income heads and your centre expenses.")
    with st.expander("➕ Add extra income"):
        with st.form("extra_income"):
            d=st.date_input("Income date",value=date.today(),key="incdate")
            typ=st.selectbox("Income category",["Extra Income","Tuition Fees (BCA/MCA)","Tuition Fees (ITI/CST)","Tuition Fees (B.Tech)","Tuition Fees (B.Sc)","Tuition Fees (Other/XI/XII)"])
            amt=st.number_input("Amount (₹)",min_value=1.0,step=100.0,key="incamt")
            desc=st.text_input("Description")
            if st.form_submit_button("Save income"):
                # Extra income is represented as an expense-register inverse? Use a dedicated settings-safe income table.
                with db() as con:
                    con.execute("""CREATE TABLE IF NOT EXISTS extra_income(id INTEGER PRIMARY KEY AUTOINCREMENT,income_date TEXT,category TEXT,amount REAL,description TEXT)""")
                    con.execute("INSERT INTO extra_income(income_date,category,amount,description) VALUES(?,?,?,?)",(str(d),typ,amt,desc))
                st.success("Income saved."); st.rerun()
    with st.expander("➖ Add expense"):
        with st.form("add_expense"):
            ed=st.date_input("Expense date",value=date.today())
            ec=st.selectbox("Expense category",["Electricity","Internet","Rent","Stationery","Equipment","Other"])
            ea=st.number_input("Expense amount (₹)",min_value=1.0,step=100.0)
            desc=st.text_input("Expense description")
            if st.form_submit_button("Save expense"):
                with db() as con: con.execute("INSERT INTO expenses(expense_date,category,amount,description) VALUES(?,?,?,?)",(str(ed),ec,ea,desc))
                st.success("Expense saved."); st.rerun()
    month=st.selectbox("Register month",[z[0] for z in months()],index=[z[0] for z in months()].index(date.today().strftime("%Y-%m")),format_func=lambda v:datetime.strptime(v,"%Y-%m").strftime("%B %Y"),key="regmonth")
    pay=qdf("SELECT payment_date date,'Tuition fee' category,amount,mode description,receipt_no reference FROM payments WHERE substr(payment_date,1,7)=?",(month,))
    exp=qdf("SELECT expense_date date,category,amount,description,CAST(id AS TEXT) reference FROM expenses WHERE substr(expense_date,1,7)=?",(month,))
    try: inc=qdf("SELECT income_date date,category,amount,description,CAST(id AS TEXT) reference FROM extra_income WHERE substr(income_date,1,7)=?",(month,))
    except: inc=pd.DataFrame(columns=["date","category","amount","description","reference"])
    alltx=pd.concat([pay,inc,exp],ignore_index=True)
    if len(alltx): st.dataframe(alltx.sort_values("date",ascending=False),use_container_width=True,hide_index=True)
    else: st.info("No transactions for this month.")
    st.metric("Expenses recorded",money(exp.amount.sum() if len(exp) else 0))

elif page=="Reports":
    st.title("Reports & History")
    st.caption("Export transaction history and review student payment records.")
    r1,r2=st.columns(2)
    start=r1.date_input("From date",value=date(date.today().year,1,1))
    end=r2.date_input("To date",value=date.today())
    payments=qdf("""SELECT p.receipt_no,p.payment_date,p.fee_month,s.name,s.category,s.course,p.amount,p.mode,p.note
                    FROM payments p JOIN students s ON s.id=p.student_id
                    WHERE p.payment_date BETWEEN ? AND ? ORDER BY p.payment_date DESC""",(str(start),str(end)))
    st.metric("Total collected in selected period",money(payments.amount.sum() if len(payments) else 0))
    st.dataframe(payments,use_container_width=True,hide_index=True)
    st.download_button("⬇️ Download payment history (CSV)",payments.to_csv(index=False).encode("utf-8-sig"),file_name="tuition_payment_history.csv",mime="text/csv")
    with st.expander("View an individual student's payment history"):
        allst=qdf("SELECT id,name FROM students ORDER BY name")
        if len(allst):
            sid=st.selectbox("Student",allst.id.tolist(),format_func=lambda i:allst.loc[allst.id==i,"name"].iloc[0])
            hist=qdf("SELECT fee_month,payment_date,amount,mode,receipt_no,note FROM payments WHERE student_id=? ORDER BY fee_month,payment_date",(int(sid),))
            st.dataframe(hist,use_container_width=True,hide_index=True)

elif page=="Backup & Data":
    st.title("Backup & Data")
    st.caption("Download a copy of your database and export your records.")
    if DB_PATH.exists():
        st.download_button("⬇️ Download full database backup",DB_PATH.read_bytes(),file_name="tuition_centre_backup.db",mime="application/octet-stream")
    for table in ["students","payments","expenses"]:
        data=qdf(f"SELECT * FROM {table}")
        st.download_button(f"Export {table.title()} (CSV)",data.to_csv(index=False).encode("utf-8-sig"),file_name=f"{table}.csv",mime="text/csv",key=f"export_{table}")
    st.info("Keep a backup regularly, especially before changing or moving the application. Your database is stored locally in tuition_centre.db.")
    st.subheader("Original import")
    st.write("The included Excel workbook is imported automatically the first time the application starts with an empty database.")
