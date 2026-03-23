# ============================================================
# LEAVE MANAGEMENT SYSTEM WITH CTO & EMAIL FIXES (FIXED)
# ============================================================

from flask import Flask, request, render_template_string, jsonify
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
import uuid
import os
import json

app = Flask(__name__)

# ==========================
# 1️⃣ CONFIGURATION
# ==========================
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
BASE_URL = os.environ.get("BASE_URL", "http://localhost:5000")
LEAVE_FILE = "Leave_Register.xlsx"
CTO_FILE = "CTO_Leave.xlsx"
MASTER_FILE = "Master_Data.xlsx"


# ============================================================
# INITIAL LEAVE CONFIG (DYNAMIC FROM MASTER FILE)
# ============================================================

def get_default_initial_leave():
    try:
        master = pd.read_excel(MASTER_FILE, sheet_name="Leave_Calc")

        # Use system default if master is empty
        if master.empty:
            return 14, 15

        # If master contains default company policy row
        return (
            int(master["initial sick leave"].iloc[0]),
            int(master["initial earn leave"].iloc[0])
        )

    except:
        # Ultimate safety fallback
        return 14, 15


# ============================================================
# SPECIAL STAFF CHECK (3-TIER WORKFLOW)
# ============================================================

def is_special_staff(designation):
    special_roles = [
        "M&E Assistant",
        "Help-Desk Assistant"
    ]

    return str(designation).strip() in special_roles


# ==========================
# 2️⃣ HELPERS
# ==========================

# ---- 2.1 Send Email ----
def send_email(to, cc, subject, html):
    import re

    # =============================
    # MAIN RECIPIENT EMAIL (With Buttons)
    # =============================
    msg_main = MIMEText(html, "html")
    msg_main["Subject"] = subject
    msg_main["From"] = EMAIL_ADDRESS
    msg_main["To"] = ", ".join(to)

    # =============================
    # CC INFORMATION ONLY VERSION
    # Remove action buttons
    # =============================
    info_html = re.sub(
        r'<a href="[^"]*".*?</a>',
        "",
        html,
        flags=re.DOTALL
    )

    msg_cc = MIMEText(info_html, "html")
    msg_cc["Subject"] = subject
    msg_cc["From"] = EMAIL_ADDRESS
    msg_cc["To"] = ", ".join(cc)

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

    # Send main recipients
    if to:
        server.sendmail(EMAIL_ADDRESS, to, msg_main.as_string())

    # Send CC recipients info-only
    if cc:
        server.sendmail(EMAIL_ADDRESS, cc, msg_cc.as_string())

    server.quit()


# ---- 2.2 Ensure Files Exist ----
def ensure_files():
    if not os.path.exists(LEAVE_FILE):
        df = pd.DataFrame(columns=[
            "request id", "staff id", "staff name", "staff email", "designation", "outlet",
            "available sick leave", "available earn leave", "leave type", "start date", "end date",
            "applied date", "# of days applied", "status", "remaining sick leave", "remaining earn leave",
            "recommended by", "recommended by cc", "approved by", "cto entitlement json", "cto enjoyment json"
        ])
        df.to_excel(LEAVE_FILE, sheet_name="Leave_Data", index=False)
    if not os.path.exists(CTO_FILE):
        df = pd.DataFrame(columns=[
            "staff name", "staff id", "designation", "outlet", "cto entitlement date", "details", "cto enjoyment date"
        ])
        df.to_excel(CTO_FILE, sheet_name="CTO_Leave", index=False)


# ---- 2.3 Safe Casting of Leave DF ----
def cast_leave_df_safe(df):
    str_cols = [
        "request id", "staff id", "staff name", "staff email", "designation", "outlet",
        "leave type", "start date", "end date", "applied date",
        "status", "recommended by", "approved by", "cto entitlement json", "cto enjoyment json"
    ]
    num_cols = ["available sick leave", "available earn leave", "# of days applied", "remaining sick leave",
                "remaining earn leave"]

    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)

    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    return df


# ---- 2.4 Safe Casting of CTO DF ----
def cast_cto_df_safe(cto_df):
    str_cols = ["staff name", "staff id", "designation", "outlet", "cto entitlement date", "details",
                "cto enjoyment date"]
    for col in str_cols:
        if col in cto_df.columns:
            cto_df[col] = cto_df[col].fillna("").astype(str)
    return cto_df


# ---- Get Initial Leave From Master (First Entry Only) ----
def get_initial_leave_from_master(staff_id):
    try:
        master = pd.read_excel(MASTER_FILE, sheet_name="Leave_Calc")

        master["staff id"] = master["staff id"].astype(str)

        row = master[master["staff id"] == str(staff_id)]

        if not row.empty:
            return (
                int(row.iloc[0]["initial sick leave"]),
                int(row.iloc[0]["initial earn leave"])
            )

    except:
        pass

    # Fallback safety values
    return get_default_initial_leave()


# ---- 2.5 Get Available Leave (used for initial leave submission) ----
def get_available_leave(df, staff_id):
    df_lower = df.copy()
    df_lower["staff id"] = df_lower["staff id"].astype(str)
    staff_rows = df_lower[(df_lower["staff id"] == str(staff_id)) & (df_lower["status"].isin(["Approved", "Rejected"]))]
    if staff_rows.empty:
        return get_initial_leave_from_master(staff_id)
    last_row = staff_rows.iloc[-1]
    return int(last_row["remaining sick leave"]), int(last_row["remaining earn leave"])


# ============================================================
# ROLE EMAIL VALIDATION + FETCH
# ============================================================

def get_role_based_emails(selected_roles):
    role_email_list = []

    try:
        role_df = pd.read_excel(MASTER_FILE, sheet_name="Role_Email")

        for role in selected_roles:

            match = role_df[
                role_df["role name"].str.lower() == role.lower()
                ]

            if match.empty:
                return False, f'Email address of "{role}" was not found'

            email = str(match.iloc[0]["email"]).strip()

            if email == "" or email.lower() == "nan":
                return False, f'Email address of "{role}" was not found'

            role_email_list.append(email)

        return True, role_email_list

    except:
        return False, "Role email configuration error"


# ==========================
# 3️⃣ DASHBOARD
# ==========================
@app.route("/")
def dashboard():
    return render_template_string("""

    <h2 style="font-family:Arial;">Staff Management Portal</h2>

    <div style="margin-top:40px;">

        <button onclick="window.location.href='/leave'"
        style="padding:20px 50px;font-size:20px;background:#007bff;
        color:white;border:none;border-radius:8px;margin-right:30px;">
        Leave Application
        </button>

        <button onclick="window.location.href='/cto'"
        style="padding:20px 50px;font-size:20px;background:#28a745;
        color:white;border:none;border-radius:8px;">
        CTO Management
        </button>

    </div>
    """)









@app.route("/validate_cto", methods=["POST"])
def validate_cto():
    staff_id = request.form.get("staff_id")
    entitlement_date = request.form.get("entitlement_date")

    if not staff_id or not entitlement_date:
        return jsonify({
            "status": "error",
            "message": "Staff ID and date required"
        })

    # =====================================================
    # Layer 1 — CTO Master Validation
    # =====================================================

    cto_df = pd.read_excel(CTO_FILE, sheet_name="CTO_Leave")
    cto_df = cast_cto_df_safe(cto_df)

    match = cto_df[
        (cto_df["staff id"] == str(staff_id)) &
        (cto_df["cto entitlement date"] == str(entitlement_date))
        ]

    if match.empty:
        return jsonify({
            "status": "error",
            "message": "Submit the CTO for the date first"
        })

    # =====================================================
    # Layer 2 — Enjoyment Lock
    # =====================================================

    if str(match.iloc[0]["cto enjoyment date"]).strip() not in ["", "nan", "None"]:
        return jsonify({
            "status": "error",
            "message": "You have already enjoyed CTO for this entitlement date"
        })

    # =====================================================
    # Layer 3 — Leave Register Usage History
    # =====================================================

    if os.path.exists(LEAVE_FILE):

        leave_df = pd.read_excel(LEAVE_FILE, sheet_name="Leave_Data")
        leave_df = cast_leave_df_safe(leave_df)

        for _, row in leave_df.iterrows():

            if str(row["staff id"]) != str(staff_id):
                continue

            try:
                entitlements = json.loads(str(row["cto entitlement json"]))

                if entitlement_date in entitlements:
                    return jsonify({
                        "status": "error",
                        "message": "This CTO date is already applied in another request"
                    })

            except:
                continue

    return jsonify({"status": "success"})


# ==========================
# STAFF AUTO FILL API
# ==========================
@app.route("/get_staff_info", methods=["POST"])
def get_staff_info():
    staff_id = request.form.get("staff_id")

    try:
        master = pd.read_excel(MASTER_FILE, sheet_name="staff_info", dtype=str)

        master["staff id"] = master["staff id"].astype(str)

        row = master[master["staff id"] == str(staff_id)]

        if not row.empty:
            return jsonify({
                "status": "success",
                "name": str(row.iloc[0]["staff name"]),
                "designation": str(row.iloc[0]["designation"]),
                "email": str(row.iloc[0]["email"])
            })

        return jsonify({
            "status": "error",
            "message": "Please contact admin to get access"
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": "System error"
        })


# ==========================
# 5️⃣ LEAVE FORM
# ==========================
@app.route("/leave", methods=["GET"])
def leave_form():
    master = pd.read_excel(MASTER_FILE)
    outlets = master["Outlet"].dropna().unique()

    return render_template_string("""

<style>

form{
    width:100%;
    max-width:1100px;
    font-family:Arial;
}

.form-group{
    display:flex;
    align-items:center;
    margin-bottom:12px;
}

label{
    width:160px;
    font-weight:bold;
}

input, select{
    flex:1;
    padding:6px;
    border-radius:4px;
    border:1px solid #ccc;
}

/* ===== CTO SECTION ===== */

.cto-row{
    display:grid;
    grid-template-columns:1fr 1fr auto;
    gap:10px;
    margin-bottom:10px;
}

</style>

<h2>Leave Form</h2>

<form method="POST" action="/submit_leave">
<div style="display:flex; gap:50px; align-items:flex-start;">
<div style="flex:2;">

<div class="form-group">
<label>Staff ID</label>
<input name="staff_id" id="staff_id">
</div>

<div class="form-group">
<label>Name</label>
<input name="staff_name" id="staff_name">
</div>

<div class="form-group">
<label>Email</label>
<input name="staff_email">
</div>

<div class="form-group">
<label>Designation</label>
<input name="designation" id="designation">
</div>

<div class="form-group">
<label>Outlet</label>
<select name="outlet">
{% for o in outlets %}
<option>{{o}}</option>
{% endfor %}
</select>
</div>

Leave Type:
<select name="leave_type">
<option>Sick Leave</option>
<option>Earn Leave</option>
</select>
<br><br>

Start Date:
<input type="date" name="start_date">

End Date:
<input type="date" name="end_date">

<h3 style="margin-top:25px;">CTO Section (Optional)</h3>

<div style="display:grid; grid-template-columns:1fr 1fr auto; font-weight:bold; margin-bottom:8px;">
    <div>CTO Entitlement Date</div>
    <div>CTO Enjoyment Date</div>
    <div></div>
</div>

<div id="cto_section"></div>

<br>
<button type="button" onclick="addCTO()">➕ Add CTO Row</button>

<input type="hidden" name="cto_entitlement_json" id="cto_entitlement_json">
<input type="hidden" name="cto_enjoyment_json" id="cto_enjoyment_json">

<br><br>

<div style="display:flex; gap:30px; align-items:center;">

    <!-- Office Resuming Date -->
    <div>
        <label><b>Office Resuming Date</b></label><br>
        <input type="date"
               name="office_resuming_date"
               id="office_resuming_date"
               required
               style="padding:8px; width:220px;">
    </div>

</div>
<br>

<div style="text-align:left; margin-top:25px;">

<button type="submit"
style="
padding:18px 55px;
font-size:20px;
font-weight:900;
background:#0056b3;
color:white;
border:none;
border-radius:8px;
cursor:pointer;
box-shadow:0px 4px 10px rgba(0,0,0,0.2);
">
✅ Submit Application
</button>

</div>
</div>
<div style="flex:1; border-left:1px solid #ddd; padding-left:25px;">
<h3 style="color:#2c3e50;">CC Email Options</h3>
<b>Role-Based CC</b><br><br>
<b>Role-Based CC</b><br><br>

<input type="checkbox" name="cc_roles" value="Admin & HR"> ADMIN & HR<br>
<input type="checkbox" name="cc_roles" value="MIS & M&E"> MIS & M&E<br>
<input type="checkbox" name="cc_roles" value="CFM Officer"> CFM OFFICER<br>

<b>Manual CC Emails</b><br><br>

<br>

<div style="font-weight:bold; margin-bottom:2px;">

</div>

<div id="cc_manual_section">

    <div class="cc-row" style="display:grid; grid-template-columns:1fr auto auto; gap:10px; margin-bottom:10px;">

        <input type="text"
               class="cc_manual_input"
               placeholder="Enter CC Email">

        <button type="button" onclick="addCC()">➕</button>

    </div>

</div>

<input type="hidden" name="cc_manual" id="cc_manual">

</div>
</div>

</form>

<script>

function addCTO(){

    let div = document.createElement("div");
    div.className = "cto-row";

    div.innerHTML = `
        <input type="date" class="entitlement">
        <input type="date" class="enjoyment" disabled>
        <button type="button" onclick="this.parentElement.remove()">✖</button>
    `;

    document.getElementById("cto_section").appendChild(div);
}

/* CTO Validation */
document.addEventListener("change", function(e){

    if(e.target.classList.contains("entitlement")){

        let staff_id = document.getElementById("staff_id").value;
        let date = e.target.value;
        let enjoyment = e.target.nextElementSibling;

        if(!date) return;

        fetch("/validate_cto",{
            method:"POST",
            headers:{"Content-Type":"application/x-www-form-urlencoded"},
            body:`staff_id=${staff_id}&entitlement_date=${date}`
        })
        .then(res=>res.json())
        .then(data=>{
            if(data.status=="error"){
                alert(data.message);
                e.target.value="";
                enjoyment.disabled=true;
            }else{
                enjoyment.disabled=false;
            }
        });
    }

});

/* Form Submit JSON Packing */
document.querySelector("form").addEventListener("submit", function(){

    let ent=[];
    let enj=[];

    document.querySelectorAll(".cto-row .entitlement").forEach(e=>ent.push(e.value));
    document.querySelectorAll(".cto-row .enjoyment").forEach(e=>enj.push(e.value));

    document.getElementById("cto_entitlement_json").value = JSON.stringify(ent);
    document.getElementById("cto_enjoyment_json").value = JSON.stringify(enj);

});
function addCC(){

    let div = document.createElement("div");

    div.style.display = "grid";
    div.style.gridTemplateColumns = "1fr auto auto";
    div.style.gap = "10px";
    div.style.marginBottom = "10px";

    div.innerHTML = `
        <input type="text" class="cc_manual_input" placeholder="Enter CC Email">
        <button type="button" onclick="removeCC(this)">✖</button>
    `;

    document.getElementById("cc_manual_section").appendChild(div);
}

function removeCC(btn){
    btn.parentElement.remove();
}

/* STAFF AUTO FILL */

document.getElementById("staff_id").addEventListener("change", function(){

    let staff_id = this.value;

    if(!staff_id) return;

    fetch("/get_staff_info",{
        method:"POST",
        headers:{
            "Content-Type":"application/x-www-form-urlencoded"
        },
        body:`staff_id=${staff_id}`
    })
    .then(res=>res.json())
    .then(data=>{

        if(data.status === "success"){

            document.getElementById("staff_name").value = data.name;
            document.getElementById("designation").value = data.designation;
            document.querySelector("input[name='staff_email']").value = data.email;

        }

        else{

            alert(data.message);

            document.getElementById("staff_name").value = "";
            document.getElementById("designation").value = "";
            document.querySelector("input[name='staff_email']").value = "";

        }

    });

});
document.querySelector("form").addEventListener("submit", function(e){

    let ccEmails = [];

    let emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    let isValid = true;

    document.querySelectorAll(".cc_manual_input").forEach(input=>{

        let val = input.value.trim();

        if(val !== ""){

            // Email format validation
            if(!emailRegex.test(val)){
                alert("Wrong email");
                isValid = false;
                return;
            }

            ccEmails.push(val);
        }

    });

    // Stop submission if invalid email found
    if(!isValid){
        e.preventDefault();
        return false;
    }

    document.getElementById("cc_manual").value = JSON.stringify(ccEmails);

});
</script>

""", outlets=outlets)


# ==========================
# 6️⃣ SUBMIT LEAVE
# ==========================
@app.route("/submit_leave", methods=["POST"])
def submit_leave():
    df = pd.read_excel(LEAVE_FILE, sheet_name="Leave_Data")
    df = cast_leave_df_safe(df)

    leave_start = request.form.get("start_date") or ""
    leave_end = request.form.get("end_date") or ""
    cto_ent_json = request.form.get("cto_entitlement_json") or ""
    cto_enj_json = request.form.get("cto_enjoyment_json") or ""

    # =====================================================
    # VALIDATION — ALLOW EITHER LEAVE OR CTO ONLY REQUEST
    # =====================================================

    normal_leave_valid = bool(leave_start and leave_end)

    cto_valid = False

    if cto_ent_json:
        try:
            entitlements = json.loads(cto_ent_json)
            enjoyments = json.loads(cto_enj_json) if cto_enj_json else []

            for i in range(len(entitlements)):
                ent = str(entitlements[i]).strip() if i < len(entitlements) else ""
                enj = str(enjoyments[i]).strip() if i < len(enjoyments) else ""

                if ent and enj:
                    cto_valid = True

        except:
            return jsonify({
                "status": "error",
                "message": "CTO validation error"
            })

    if not normal_leave_valid and not cto_valid:
        return "Please provide either Leave dates OR Complete CTO dates"

    # =====================================================
    # CTO Enjoyment Date Mandatory Validation
    # =====================================================

    if cto_ent_json:
        try:
            entitlements = json.loads(cto_ent_json)
            enjoyments = json.loads(cto_enj_json) if cto_enj_json else []

            for i in range(len(entitlements)):
                ent = str(entitlements[i]).strip() if i < len(entitlements) else ""
                enj = str(enjoyments[i]).strip() if i < len(enjoyments) else ""

                if ent and not enj:
                    return "Please provide enjoyment date"

        except:
            return jsonify({
                "status": "error",
                "message": "CTO validation error"
            })
    office_resuming_date = request.form.get("office_resuming_date") or ""

    # ============================================================
    # CC EMAIL VALIDATION (ROLE + MANUAL)
    # ============================================================

    # Role Based CC Selection
    selected_roles = request.form.getlist("cc_roles")

    role_status, role_emails = get_role_based_emails(selected_roles)

    if not role_status:
        return jsonify({
            "status": "error",
            "message": role_emails
        })

    # Manual CC Emails
    manual_cc_raw = request.form.get("cc_manual") or "[]"

    try:
        manual_cc_list = json.loads(manual_cc_raw)
    except:
        manual_cc_list = []

    # Combine CC lists
    final_cc_list = list(set(role_emails + manual_cc_list))

    # =====================================================
    # CTO ONLY APPLICATION VALIDATION
    # =====================================================

    staff_id = request.form.get("staff_id")

    if cto_ent_json:

        try:
            entitlements = json.loads(cto_ent_json)

            cto_df = pd.read_excel(CTO_FILE, sheet_name="CTO_Leave")
            cto_df = cast_cto_df_safe(cto_df)

            for ent in entitlements:

                match = cto_df[
                    (cto_df["staff id"] == str(staff_id)) &
                    (cto_df["cto entitlement date"] == str(ent))
                    ]

                # Layer 1 — CTO Master Check
                if match.empty:
                    return jsonify({
                        "status": "error",
                        "message": f"CTO entitlement {ent} not found"
                    })

                # Layer 2 — Enjoyment Lock
                if str(match.iloc[0]["cto enjoyment date"]).strip() not in ["", "nan", "None"]:
                    return jsonify({
                        "status": "error",
                        "message": f"CTO {ent} already enjoyed"
                    })

        except:
            return jsonify({
                "status": "error",
                "message": "CTO validation error"
            })

    # =====================================================
    # CTO HARD SERVER SIDE PROTECTION (VERY IMPORTANT)
    # =====================================================

    cto_ent_json = request.form.get("cto_entitlement_json") or ""

    if cto_ent_json:

        try:
            entitlements = json.loads(cto_ent_json)

            if os.path.exists(LEAVE_FILE):

                leave_df = pd.read_excel(LEAVE_FILE, sheet_name="Leave_Data")

                for _, row in leave_df.iterrows():

                    if str(row["staff id"]) != str(staff_id):
                        continue

                    try:
                        existing_entitlements = json.loads(str(row["cto entitlement json"]))

                        for ent in entitlements:
                            if ent in existing_entitlements:
                                return "Error: This CTO date is already applied in another request"

                    except:
                        continue

        except:
            pass
    # ---- Duplicate Leave Check ----
    if leave_start and leave_end:
        leave_start_dt = pd.to_datetime(leave_start)
        leave_end_dt = pd.to_datetime(leave_end)
        staff_rows = df[df["staff id"] == staff_id]
        for _, row in staff_rows.iterrows():
            # Skip pending leaves? You can decide; here we check only Pending/Approved
            if row["status"] not in ["Pending", "Approved"]:
                continue
            row_start = pd.to_datetime(row["start date"]) if row["start date"] else None
            row_end = pd.to_datetime(row["end date"]) if row["end date"] else None
            if row_start and row_end:
                # Check for overlapping
                if leave_start_dt <= row_end and leave_end_dt >= row_start:
                    return f"Error: You already have leave from {row_start.date()} to {row_end.date()}"

    days = (pd.to_datetime(leave_end) - pd.to_datetime(leave_start)).days + 1 if leave_start and leave_end else 0

    avail_sick, avail_earn = get_available_leave(df, staff_id)

    new_row = {
        "request id": str(uuid.uuid4())[:8],
        "staff id": staff_id,
        "staff name": request.form.get("staff_name"),
        "staff email": request.form.get("staff_email"),
        "designation": request.form.get("designation"),
        "outlet": request.form.get("outlet"),
        "available sick leave": avail_sick,
        "available earn leave": avail_earn,
        "leave type": request.form.get("leave_type", ""),
        "start date": leave_start,
        "end date": leave_end,
        "office resuming date": office_resuming_date,
        "applied date": datetime.now().strftime("%Y-%m-%d"),
        "# of days applied": days,
        "status": "Pending",
        "remaining sick leave": avail_sick,
        "remaining earn leave": avail_earn,
        "recommended by": "",
        "recommended by cc": "",
        "approved by": "",
        "cto entitlement json": cto_ent_json,
        "cto enjoyment json": cto_enj_json,
        "cc emails json": json.dumps(final_cc_list)
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(LEAVE_FILE, sheet_name="Leave_Data", index=False)

    # === SEND MAIL TO OS + OMTL AFTER LEAVE APPLICATION ===
    master = pd.read_excel(MASTER_FILE)
    master_row = master[master["Outlet"] == new_row["outlet"]].iloc[0]
    recipients = [master_row["OS Email"], master_row["OMTL Email"]]

    # ---------- Build CTO Section ----------
    cto_section = ""

    entitlements = json.loads(new_row["cto entitlement json"]) if new_row["cto entitlement json"] else []
    enjoyments = json.loads(new_row["cto enjoyment json"]) if new_row["cto enjoyment json"] else []

    cto_rows = ""

    for i in range(len(entitlements)):
        ent = entitlements[i] if i < len(entitlements) else ""
        enj = enjoyments[i] if i < len(enjoyments) else ""

        if str(ent).strip() != "":
            cto_rows += f"""
            <tr>
                <td>{ent}</td>
                <td>{enj}</td>
            </tr>
            """

    if cto_rows != "":
        cto_section = f"""
        <h4 style="margin-top:20px;color:#2c3e50;">CTO Information</h4>
        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:50%;">
            <tr style="background:#f2f2f2;">
                <th>CTO Entitlement Date</th>
                <th>CTO Enjoyment Date</th>
            </tr>
            {cto_rows}
        </table>
        """

    # ---------- MAIN EMAIL DESIGN ----------
    html = f"""
    <h3 style="color:#2c3e50;">New Leave Request (Recommendation Needed)</h3>

    <table border="1" cellpadding="6" cellspacing="0" 
    style="border-collapse:collapse;width:50%;">

        <tr><th align="left">Staff Name</th><td>{new_row['staff name']}</td></tr>
        <tr><th align="left">Staff ID</th><td>{new_row['staff id']}</td></tr>
        <tr><th align="left">Designation</th><td>{new_row['designation']}</td></tr>
        <tr><th align="left">Outlet</th><td>{new_row['outlet']}</td></tr>
        <tr><th align="left">Leave Type</th><td>{new_row['leave type']}</td></tr>

        <tr><th align="left">Leave Dates</th>
            <td>{new_row['start date']} → {new_row['end date']}</td>
        <tr>
<th align="left">Office Resuming Date</th>
<td>{office_resuming_date}</td>
</tr>

        <tr><th align="left">Number of Days</th><td>{new_row['# of days applied']}</td></tr>
        <tr><th align="left">Available Sick Leave</th><td>{avail_sick}</td></tr>
        <tr><th align="left">Available Earn Leave</th><td>{avail_earn}</td></tr>

    </table>

    {cto_section}

    <br><br>

    <br><br>

<div style="text-align:center;">

<a href="{BASE_URL}/leave/recommend/{new_row['request id']}"
style="background:#007bff;color:white;padding:12px 22px;
text-decoration:none;border-radius:6px;
font-weight:bold;margin-right:15px;display:inline-block;">
✅ RECOMMEND
</a>

<a href="{BASE_URL}/leave/reject/{new_row['request id']}"
style="background:#dc3545;color:white;padding:12px 22px;
text-decoration:none;border-radius:6px;
font-weight:bold;display:inline-block;">
❌ REJECT
</a>

</div>
    """
    designation = new_row["designation"]

    # FIRST STAGE MAIL → OS / OMTL ONLY
    send_email(recipients, [], "Leave Recommendation Needed", html)

    # NORMAL STAFF → SEND CC INFORMATION
    if not is_special_staff(designation):
        send_email([], final_cc_list, "Leave Application Information", html)
    return jsonify({
        "status": "success",
        "message": "Leave Submitted Successfully"
    })


# ==========================
# 7️⃣ RECOMMEND (FIXED VERSION)
# ==========================
@app.route("/leave/recommend/<req_id>")
def recommend(req_id):
    df = pd.read_excel(LEAVE_FILE, sheet_name="Leave_Data")
    df = cast_leave_df_safe(df)

    rows = df[df["request id"] == req_id]

    # ⭐ FIX 1 — Check empty BEFORE accessing row
    if rows.empty:
        return "Invalid Request ID"

    # ⭐ FIX 2 — Safe row extraction
    row = rows.iloc[0]
    office_resuming_date = row.get("office resuming date", "")
    idx = rows.index[0]

    if df.at[idx, "recommended by"].strip() != "":
        return "Already Recommended"

    df.at[idx, "recommended by"] = "OS/OMTL"
    df.to_excel(LEAVE_FILE, sheet_name="Leave_Data", index=False)

    staff_id = row["staff id"]
    avail_sick, avail_earn = get_available_leave(df, staff_id)

    # ---------- MASTER DATA ----------
    master = pd.read_excel(MASTER_FILE)
    master_row = master[master["Outlet"] == row["outlet"]].iloc[0]

    pm_email = [master_row["PM Email"]]

    # ---------- CTO SECTION ----------
    cto_section = ""

    # ⭐ FIX 3 — Safe JSON parsing
    entitlements = []
    enjoyments = []

    if str(row["cto entitlement json"]).strip():
        entitlements = json.loads(row["cto entitlement json"])

    if str(row["cto enjoyment json"]).strip():
        enjoyments = json.loads(row["cto enjoyment json"])

    cto_rows = ""

    for i in range(max(len(entitlements), len(enjoyments))):

        ent = entitlements[i] if i < len(entitlements) else ""
        enj = enjoyments[i] if i < len(enjoyments) else ""

        if str(ent).strip() != "":
            cto_rows += f"""
            <tr>
                <td>{ent}</td>
                <td>{enj}</td>
            </tr>
            """

    if cto_rows:
        cto_section = f"""
        <h4 style="margin-top:25px;color:#2c3e50;">CTO Information</h4>

        <table border="1" cellpadding="8" cellspacing="0"
        style="border-collapse:collapse;width:50%;font-family:Arial;">

            <tr style="background:#f2f2f2;">
                <th>CTO Entitlement Date</th>
                <th>CTO Enjoyment Date</th>
            </tr>

            {cto_rows}

        </table>
        """

    # ---------- EMAIL DESIGN ----------
    html = f"""
    <h3 style="color:#1e7e34;">Leave Approval Needed</h3>

    <table border="1" cellpadding="8" cellspacing="0"
    style="border-collapse:collapse;width:50%;font-family:Arial;">

    <tr><th align="left">Staff Name</th><td>{row['staff name']}</td></tr>
    <tr><th align="left">Staff ID</th><td>{row['staff id']}</td></tr>
    <tr><th align="left">Designation</th><td>{row['designation']}</td></tr>
    <tr><th align="left">Outlet</th><td>{row['outlet']}</td></tr>

    <tr><th align="left">Leave Type</th><td>{row['leave type']}</td></tr>

    <tr><th align="left">Leave Dates</th>
    <td>{row['start date']} → {row['end date']}</td></tr>

    <tr><th align="left">Number of Days</th>
    <td>{row['# of days applied']}</td></tr>

    <tr><th align="left">Available Sick Leave</th>
    <td>{avail_sick}</td></tr>

    <tr><th align="left">Available Earn Leave</th>
    <td>{avail_earn}</td></tr>
    <tr>
<th align="left">Office Resuming Date</th>
<td>{office_resuming_date}</td>
</tr>

    </table>

    {cto_section}

    <br><br>

    <a href="{BASE_URL}/leave/approve/{req_id}"
    style="background:green;color:white;padding:10px 18px;
    text-decoration:none;border-radius:6px;font-weight:bold;">
    APPROVE
    </a>

    &nbsp;

    <a href="{BASE_URL}/leave/reject/{req_id}"
    style="background:red;color:white;padding:10px 18px;
    text-decoration:none;border-radius:6px;font-weight:bold;">
    REJECT
    </a>
    """

    designation = row["designation"]

    # ============================================
    # NORMAL STAFF → SEND DIRECTLY TO PM
    # ============================================

    if not is_special_staff(designation):
        send_email(pm_email, [], "Leave Approval Needed", html)

        return "Recommended Successfully"

    # ============================================
    # SPECIAL STAFF → SEND TO CC FOR 2ND REVIEW
    # ============================================

    try:
        cc_list = json.loads(row.get("cc emails json", "[]"))
    except:
        cc_list = []

    if not cc_list:
        return "No CC reviewer configured"

    cc_html = html.replace(
        f"{BASE_URL}/leave/approve/{req_id}",
        f"{BASE_URL}/leave/recommend_cc/{req_id}"
    )

    send_email(cc_list, [], "Leave Recommendation Needed (CC Review)", cc_html)

    return "Recommendation Completed by OS/OMTL"


# ==========================
# SECOND RECOMMENDATION (CC)
# ==========================

@app.route("/leave/recommend_cc/<req_id>")
def recommend_cc(req_id):
    df = pd.read_excel(LEAVE_FILE, sheet_name="Leave_Data")
    df = cast_leave_df_safe(df)

    # Ensure column is string
    if "recommended by cc" not in df.columns:
        df["recommended by cc"] = ""
    else:
        df["recommended by cc"] = df["recommended by cc"].astype(str)

    rows = df[df["request id"] == req_id]
    if rows.empty:
        return "Invalid Request ID"
    idx = rows.index[0]
    val = df.at[idx, "recommended by cc"]

    if pd.notna(val) and str(val).strip() != "":
        return "Already Recommended By CC"

    df.at[idx, "recommended by cc"] = "CC"

    df.to_excel(LEAVE_FILE, sheet_name="Leave_Data", index=False)

    row = rows.iloc[0]

    master = pd.read_excel(MASTER_FILE)
    master_row = master[master["Outlet"] == row["outlet"]].iloc[0]

    pm_email = [master_row["PM Email"]]

    html = f"""
    <h3>Leave Approval Needed</h3>

    <p>Second recommendation completed.</p>

    <table border="1" cellpadding="6">

    <tr><td>Staff Name</td><td>{row['staff name']}</td></tr>
    <tr><td>Staff ID</td><td>{row['staff id']}</td></tr>
    <tr><td>Designation</td><td>{row['designation']}</td></tr>
    <tr><td>Outlet</td><td>{row['outlet']}</td></tr>
    <tr><td>Leave Type</td><td>{row['leave type']}</td></tr>
    <tr><td>Leave Dates</td><td>{row['start date']} → {row['end date']}</td></tr>
    <tr><td>Total Days</td><td>{row.get('total days', '')}</td></tr>
    <tr><td>Reason</td><td>{row.get('reason', '')}</td></tr>
    <tr><td>Recommended by Outlet</td><td>{row.get('recommended by outlet', '')}</td></tr>
    <tr><td>Recommended by Officer</td><td>{row.get('recommended by officer', '')}</td></tr>

    </table>



    <br><br>

    <a href="{BASE_URL}/leave/approve/{req_id}"
    style="background:green;color:white;padding:10px 18px;
    text-decoration:none;border-radius:6px;font-weight:bold;">
    APPROVE
    </a>

    &nbsp;

    <a href="{BASE_URL}/leave/reject/{req_id}"
    style="background:red;color:white;padding:10px 18px;
    text-decoration:none;border-radius:6px;font-weight:bold;">
    REJECT
    </a>
    """

    send_email(pm_email, [], "Leave Approval Needed", html)

    return "Recommended by Both the Outlet and Officer Level"


# ==========================
# 8️⃣ APPROVE (FIXED CUMULATIVE LEAVE)
# ==========================
@app.route("/leave/approve/<req_id>")
def approve(req_id):
    df = pd.read_excel(LEAVE_FILE, sheet_name="Leave_Data")
    df = cast_leave_df_safe(df)
    rows = df[df["request id"] == req_id]

    if rows.empty:
        return "Invalid Request ID"
    idx = rows.index[0]
    # Final Decision Lock Protection
    if df.at[idx, "status"] in ["Approved", "Rejected"]:
        return "This request is already finalized"

    df.at[idx, "status"] = "Approved"
    df.at[idx, "approved by"] = "PM"

    leave_type = df.at[idx, "leave type"]
    days = int(df.at[idx, "# of days applied"])
    staff_id = df.at[idx, "staff id"]

    # ---- FIX: Compute cumulative remaining leave for staff ----
    staff_leaves = df[(df["staff id"] == staff_id) & (df["status"] == "Approved")]
    total_sick_used = staff_leaves[staff_leaves["leave type"] == "Sick Leave"]["# of days applied"].sum()
    total_earn_used = staff_leaves[staff_leaves["leave type"] == "Earn Leave"]["# of days applied"].sum()

    # Get entitlement dynamically from master policy
    initial_sick, initial_earn = get_initial_leave_from_master(staff_id)

    remaining_sick = initial_sick - total_sick_used
    remaining_earn = initial_earn - total_earn_used

    # Update leave record
    df.at[idx, "remaining sick leave"] = remaining_sick
    df.at[idx, "remaining earn leave"] = remaining_earn
    df.at[idx, "available sick leave"] = remaining_sick + days if leave_type == "Sick Leave" else remaining_sick
    df.at[idx, "available earn leave"] = remaining_earn + days if leave_type == "Earn Leave" else remaining_earn

    # ---- Update CTO ----
    entitlements = json.loads(df.at[idx, "cto entitlement json"]) if df.at[idx, "cto entitlement json"] else []
    enjoyments = json.loads(df.at[idx, "cto enjoyment json"]) if df.at[idx, "cto enjoyment json"] else []
    cto_df = pd.read_excel(CTO_FILE, sheet_name="CTO_Leave")
    cto_df = cast_cto_df_safe(cto_df)
    for ent, enj in zip(entitlements, enjoyments):
        mask = (cto_df["staff id"] == str(staff_id)) & (cto_df["cto entitlement date"] == str(ent))
        cto_df.loc[mask, "cto enjoyment date"] = str(enj)
    cto_df.to_excel(CTO_FILE, sheet_name="CTO_Leave", index=False)

    # Sort by Staff ID + Remaining Balance (Descending Order ⭐)
    df = df.sort_values(
        by=["staff id", "remaining sick leave", "remaining earn leave"],
        ascending=[True, False, False]
    )

    df.to_excel(LEAVE_FILE, sheet_name="Leave_Data", index=False)

    master = pd.read_excel(MASTER_FILE)
    master_row = master[master["Outlet"] == rows.iloc[0]["outlet"]].iloc[0]
    cc_list = []

    try:
        cc_list = json.loads(rows.iloc[0].get("cc emails json", "[]"))
    except:
        cc_list = []

    recipients = [
        master_row["OS Email"],
        master_row["OMTL Email"],
        rows.iloc[0]["staff email"]
    ]

    # ---------- CTO SECTION ----------
    cto_section = ""

    entitlements = json.loads(rows.iloc[0]["cto entitlement json"]) if rows.iloc[0]["cto entitlement json"] else []
    enjoyments = json.loads(rows.iloc[0]["cto enjoyment json"]) if rows.iloc[0]["cto enjoyment json"] else []

    cto_rows = ""

    for i in range(len(entitlements)):
        ent = entitlements[i] if i < len(entitlements) else ""
        enj = enjoyments[i] if i < len(enjoyments) else ""

        if str(ent).strip() != "":
            cto_rows += f"""
            <tr>
                <td>{ent}</td>
                <td>{enj}</td>
            </tr>
            """

    if cto_rows != "":
        cto_section = f"""
        <h4 style="margin-top:25px;color:#2c3e50;">CTO Information</h4>

        <table border="1" cellpadding="8" cellspacing="0"
        style="border-collapse:collapse;width:100%;font-family:Arial;">

            <tr style="background:#f2f2f2;">
                <th>CTO Entitlement Date</th>
                <th>CTO Enjoyment Date</th>
            </tr>

            {cto_rows}

        </table>
        """

    # ---------- APPROVAL EMAIL ----------
    html = f"""
    <h3 style="color:#1e7e34;">Leave Approved</h3>

    <table border="1" cellpadding="8" cellspacing="0"
    style="border-collapse:collapse;width:70%;font-family:Arial;">

    <tr>
    <th align="left">Staff Name</th>
    <td>{rows.iloc[0]['staff name']}</td>
    </tr>

    <tr>
    <th align="left">Staff ID</th>
    <td>{rows.iloc[0]['staff id']}</td>
    </tr>

    <tr>
    <th align="left">Designation</th>
    <td>{rows.iloc[0]['designation']}</td>
    </tr>

    <tr>
    <th align="left">Outlet</th>
    <td>{rows.iloc[0]['outlet']}</td>
    </tr>

    <tr>
    <th align="left">Leave Type</th>
    <td>{rows.iloc[0]['leave type']}</td>
    </tr>

    <tr>
    <th align="left">Leave Dates</th>
    <td>{rows.iloc[0]['start date']} → {rows.iloc[0]['end date']}</td>
    </tr>

    <tr>
    <th align="left">Number of Days</th>
    <td>{rows.iloc[0]['# of days applied']}</td>
    </tr>
    <tr>
<th align="left">Office Resuming Date</th>
<td><b>{rows.iloc[0].get("office resuming date", "")}</b></td>
</tr>

    <tr>
    <th align="left">Status</th>
    <td>{df.at[idx, 'status']}</td>
    </tr>

    <tr>
    <th align="left">Remaining Sick Leave</th>
    <td>{remaining_sick}</td>
    </tr>

    <tr>
    <th align="left">Remaining Earn Leave</th>
    <td>{remaining_earn}</td>
    </tr>

    </table>

    {cto_section}
    """
    # Approval information should go to staff + OS + OMTL + CC list
    cc_list = []

    try:
        cc_list = json.loads(rows.iloc[0].get("cc emails json", "[]"))
    except:
        cc_list = []

    send_email(
        recipients,
        cc_list,
        "Leave Approved",
        html
    )

    return "Approved Successfully"


# ==========================
# 8️⃣ REJECT
# ==========================
@app.route("/leave/reject/<req_id>", methods=["GET", "POST"])
def reject(req_id):
    df = pd.read_excel(LEAVE_FILE, sheet_name="Leave_Data")
    df = cast_leave_df_safe(df)

    rows = df[df["request id"] == req_id]

    if rows.empty:
        return "Already Rejected"

    idx = rows.index[0]

    # 🔒 Block rejection if already approved
    if str(df.at[idx, "status"]).strip() == "Approved":
        return "This request has already been approved and cannot be rejected."

    if request.method == "GET":
        # Show message input form to PM
        return f"""
        <h3>Reject Leave Request</h3>
        <form method="POST">
        <label>Message to Staff (Optional)</label><br>
        <textarea name="reject_message"
        style="width:400px;height:120px;"></textarea>

        <br><br>

        <button type="submit"
        style="padding:12px 30px;font-size:18px;">
        Send Rejection
        </button>

        </form>
        """

    # POST = Send rejection
    reject_message = request.form.get("reject_message", "")
    # ==========================
    # 🔥 DETERMINE WHO REJECTED
    # ==========================

    recommended_by = str(rows.iloc[0].get("recommended by", "")).strip()

    recommended_by_cc = str(rows.iloc[0].get("recommended by cc", "")).strip()

    if recommended_by == "":
        rejected_by = "OS/OMTL"

    elif recommended_by != "" and recommended_by_cc == "":
        rejected_by = "CC"

    else:
        rejected_by = "PM"

    # Delete the rejected row completely
    df = df[df["request id"] != req_id]

    df.to_excel(LEAVE_FILE, sheet_name="Leave_Data", index=False)

    # Send Email
    master = pd.read_excel(MASTER_FILE)
    master_row = master[master["Outlet"] == rows.iloc[0]["outlet"]].iloc[0]

    recipients = [
        master_row["OS Email"],
        master_row["OMTL Email"],
        rows.iloc[0]["staff email"]
    ]

    # Rejection Email Body##############
    html = f"""
    <p><b>Message From {rejected_by}:</b></p>
    <p>{reject_message if reject_message else "No message provided"}</p>
    <table border="1" cellpadding="6">
    <tr><td>Staff Name</td><td>{rows.iloc[0]['staff name']}</td></tr>
    <tr><td>Leave Type</td><td>{rows.iloc[0]['leave type']}</td></tr>
    <tr><td>Status</td><td>Rejected</td></tr>
    </table>
    """

    # Retrieve CC list stored during submission
    cc_list = []

    try:
        cc_list = json.loads(rows.iloc[0].get("cc emails json", "[]"))
    except:
        cc_list = []

    send_email(
        recipients,
        cc_list,
        "Leave Rejected",
        html
    )

    return "Rejected Successfully"


###########################################################################################
###############################################$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
#$$$$$$$$$$$$######$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
##########################3333333333333333333##################33333333333333333333######333333
# ============================================================
# CTO MANAGEMENT SYSTEM - SEPARATE FILE VERSION
# ============================================================

from flask import Flask, request, render_template_string, redirect
import pandas as pd
import smtplib
from email.mime.text import MIMEText
import base64
import json
import os



# ============================================================
# FILE CONFIGURATION
# ============================================================

CTO_FILE = "CTO_Leave.xlsx"          # Separate file to store CTO
CTO_SHEET = "CTO_Leave"
OUTLET_FILE = "Master_Data.xlsx"     # Contains Outlet_Master

EMAIL_ADDRESS = "leavedata.system@gmail.com"
EMAIL_PASSWORD = "huyticknixqijeyv"

BASE_URL = "http://localhost:5000"

# ============================================================
# EMAIL FUNCTION
# ============================================================

def send_email(to_list, subject, html_body):
    msg = MIMEText(html_body, "html")
    msg["Subject"] = subject
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = ", ".join(to_list)

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    server.sendmail(EMAIL_ADDRESS, to_list, msg.as_string())
    server.quit()

# ============================================================
# ENSURE CTO FILE EXISTS
# ============================================================

def ensure_cto_file():
    if not os.path.exists(CTO_FILE):
        df = pd.DataFrame(columns=[
            "Staff Name", "Staff ID", "Designation", "Outlet",
            "CTO Entitlement Date", "Details"
        ])
        df.to_excel(CTO_FILE, sheet_name=CTO_SHEET, index=False)

# ============================================================
# ROOT
# ============================================================



# ============================================================
# CTO FORM
# ============================================================

@app.route("/cto", methods=["GET", "POST"])
def cto_form():

    ensure_cto_file()

    try:
        outlet_master = pd.read_excel(OUTLET_FILE, sheet_name="Outlet_Master")
        cto_sheet = pd.read_excel(CTO_FILE, sheet_name=CTO_SHEET)
    except Exception as e:
        return f"Excel read error: {e}"

    outlet_master.columns = outlet_master.columns.str.strip().str.lower()
    cto_sheet.columns = cto_sheet.columns.str.strip().str.lower()

    outlets = outlet_master["outlet"].dropna().unique()

    if request.method == "POST":
        outlet = request.form["outlet"]
        staff_name = request.form["staff_name"]
        staff_id = request.form["staff_id"]
        designation = request.form["designation"]
        cto_date = request.form["cto_date"]
        details = request.form["details"]

        # Prevent duplicate Staff ID + Date
        duplicate = cto_sheet[
            (cto_sheet["staff id"].astype(str) == str(staff_id)) &
            (cto_sheet["cto entitlement date"].astype(str) == str(cto_date))
        ]
        if not duplicate.empty:
            return "❌ This staff already has CTO for this date."

        # Get OS & OMTL emails from Outlet_Master
        outlet_row = outlet_master[outlet_master["outlet"] == outlet]
        if outlet_row.empty:
            return "Invalid outlet selected."

        os_email = outlet_row.iloc[0]["os email"]
        omtl_email = outlet_row.iloc[0]["omtl email"]

        # Encode data for approval link
        data = {
            "outlet": outlet,
            "staff name": staff_name,
            "staff id": staff_id,
            "designation": designation,
            "cto entitlement date": cto_date,
            "details": details
        }
        encoded = base64.urlsafe_b64encode(json.dumps(data).encode()).decode()
        approve_link = f"{BASE_URL}/cto/approve/{encoded}"

        # Send Approval Email
        html_content = f"""
        <h3>CTO Approval Required</h3>
        <p>
        Staff: {staff_name}<br>
        Staff ID: {staff_id}<br>
        Designation: {designation}<br>
        Outlet: {outlet}<br>
        CTO Date: {cto_date}<br>
        Details: {details}
        </p>
        <a href="{approve_link}" 
        style="padding:10px 20px;background:green;
        color:white;text-decoration:none;border-radius:5px;">
        APPROVE
        </a>
        """

        send_email([os_email, omtl_email], "CTO Approval Required", html_content)
        return "✅ CTO submitted. Approval email sent."

    return render_template_string("""
    <h2>CTO Request Form</h2>
    <form method="post">
        Outlet:
        <select name="outlet" required>
        {% for o in outlets %}
            <option value="{{o}}">{{o}}</option>
        {% endfor %}
        </select><br><br>

        Staff Name: <input name="staff_name" required><br><br>
        Staff ID: <input name="staff_id" required><br><br>
        Designation: <input name="designation" required><br><br>
        CTO Entitlement Date:
        <input type="date" name="cto_date" required><br><br>
        Details:
        <input name="details"><br><br>

        <button type="submit">Submit</button>
    </form>
    """, outlets=outlets)

# ============================================================
# APPROVAL ROUTE
# ============================================================

@app.route("/cto/approve/<encoded_data>")
def approve_cto (encoded_data):

    try:
        decoded = json.loads(base64.urlsafe_b64decode(encoded_data.encode()).decode())
    except:
        return "Invalid approval link."

    ensure_cto_file()

    try:
        cto_sheet = pd.read_excel(CTO_FILE, sheet_name=CTO_SHEET)
    except Exception as e:
        return f"Error reading CTO sheet: {e}"

    cto_sheet.columns = cto_sheet.columns.str.strip().str.lower()

    # Only first click writes
    duplicate = cto_sheet[
        (cto_sheet["staff id"].astype(str) == str(decoded["staff id"])) &
        (cto_sheet["cto entitlement date"].astype(str) == str(decoded["cto entitlement date"]))
    ]
    if not duplicate.empty:
        return "⚠ Already Approved"

    # Append approved data
    updated = pd.concat([cto_sheet, pd.DataFrame([decoded])], ignore_index=True)

    try:
        with pd.ExcelWriter(CTO_FILE, engine="openpyxl", mode='w') as writer:
            updated.to_excel(writer, sheet_name=CTO_SHEET, index=False)
    except Exception as e:
        return f"Error writing CTO_Leave: {e}"

    return "✅ CTO Approved Successfully"

# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    ensure_files()
    ensure_cto_file()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
