import pandas as pd
import uuid
from flask import Flask, render_template_string, request, redirect
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Excel files
LEAVE_FILE = "Leave_Register.xlsx"  # Existing file
LEAVE_SHEET = "Leave_Data"

MASTER_FILE = "Master_Data.xlsx"  # Contains Outlet, OS/OMTL emails, PM Name/Email

# Email config
EMAIL_ADDRESS = "your_email@gmail.com"
EMAIL_PASSWORD = "your_app_password"  # Gmail App Password

app = Flask(__name__)


# Helper to get BASE_URL dynamically
def get_base_url():
    return f"{request.scheme}://{request.host}"


# Helper to send email
def send_email(to, cc, subject, html):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = subject
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, to + cc, msg.as_string())


# Load Master Data
def load_master():
    df = pd.read_excel(MASTER_FILE)
    df.fillna("", inplace=True)
    return df


# Flask Routes
@app.route("/", methods=["GET", "POST"])
def form():
    master_df = load_master()
    outlets = master_df["Outlet"].tolist()
    leave_types = ["Sick Leave", "Earn Leave"]

    if request.method == "POST":
        # Collect form data
        staff_id = request.form["staff_id"]
        staff_name = request.form["staff_name"]
        staff_email = request.form["staff_email"]
        designation = request.form["designation"]
        outlet = request.form["outlet"]
        leave_type = request.form["leave_type"]
        start_date = request.form["start_date"]
        end_date = request.form["end_date"]

        # Generate unique Request ID
        request_id = str(uuid.uuid4())[:8]

        # Calculate leave days
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        leave_days = (end - start).days + 1

        # Load leave Excel
        leave_df = pd.read_excel(LEAVE_FILE, sheet_name=LEAVE_SHEET)

        # Remaining leaves
        remaining_earned = leave_df.loc[leave_df["Staff ID"] == staff_id, "Remaining Earned"].max()
        remaining_sick = leave_df.loc[leave_df["Staff ID"] == staff_id, "Remaining Sick"].max()

        if leave_type == "Earn Leave":
            remaining_earned = remaining_earned - leave_days
        else:
            remaining_sick = remaining_sick - leave_days

        # Add new row
        new_row = {
            "Request ID": request_id,
            "Staff ID": staff_id,
            "Staff Name": staff_name,
            "Staff Email": staff_email,
            "Designation": designation,
            "Outlet": outlet,
            "Leave Type": leave_type,
            "Start Date": start_date,
            "End Date": end_date,
            "Status": "Pending",
            "Leave Day": leave_days,
            "Remaining Earned": remaining_earned,
            "Remaining Sick": remaining_sick
        }
        leave_df = pd.concat([leave_df, pd.DataFrame([new_row])], ignore_index=True)
        leave_df.to_excel(LEAVE_FILE, sheet_name=LEAVE_SHEET, index=False)

        # Email to OS + OMTL
        row = master_df[master_df["Outlet"] == outlet].iloc[0]
        os_email = row["OS Email"]
        omtl_email = row["OMTL Email"]
        pm_name = row["PM Name"]
        pm_email = row["PM Email"]

        base_url = get_base_url()
        recommend_link = f"{base_url}/recommend/{request_id}"

        subject = f"Leave Request Recommendation: {staff_name}"
        html = f"""
        <h3>Leave Request from {staff_name}</h3>
        <p>Leave Type: {leave_type}<br>Start: {start_date}<br>End: {end_date}</p>
        <a href="{recommend_link}" style="padding:10px;background:orange;color:white;text-decoration:none;">Recommend</a>
        """

        send_email(to=[os_email, omtl_email], cc=[], subject=subject, html=html)

        return "Leave submitted successfully! Recommendation email sent to OS & OMTL."

    # Render HTML form
    form_html = """
    <h2>Leave Application Form</h2>
    <form method="post">
        Staff ID: <input name="staff_id" required><br>
        Staff Name: <input name="staff_name" required><br>
        Staff Email: <input name="staff_email" required><br>
        Designation: <input name="designation"><br>
        Outlet: <select name="outlet">
            {% for outlet in outlets %}
                <option value="{{ outlet }}">{{ outlet }}</option>
            {% endfor %}
        </select><br>
        Leave Type: <select name="leave_type">
            {% for lt in leave_types %}
                <option value="{{ lt }}">{{ lt }}</option>
            {% endfor %}
        </select><br>
        Start Date: <input type="date" name="start_date" required><br>
        End Date: <input type="date" name="end_date" required><br>
        <input type="submit" value="Apply">
    </form>
    """
    return render_template_string(form_html, outlets=outlets, leave_types=leave_types)


# OS/OMTL recommend
@app.route("/recommend/<request_id>")
def recommend(request_id):
    leave_df = pd.read_excel(LEAVE_FILE, sheet_name=LEAVE_SHEET)
    idx = leave_df[leave_df["Request ID"] == request_id].index[0]
    leave_df.at[idx, "Status"] = "Recommended"
    leave_df.to_excel(LEAVE_FILE, sheet_name=LEAVE_SHEET, index=False)

    # Send email to PM
    master_df = load_master()
    outlet = leave_df.at[idx, "Outlet"]
    staff_name = leave_df.at[idx, "Staff Name"]
    leave_type = leave_df.at[idx, "Leave Type"]
    start_date = leave_df.at[idx, "Start Date"]
    end_date = leave_df.at[idx, "End Date"]

    row = master_df[master_df["Outlet"] == outlet].iloc[0]
    pm_email = row["PM Email"]
    os_email = row["OS Email"]
    omtl_email = row["OMTL Email"]

    base_url = get_base_url()
    approve_link = f"{base_url}/approve/{request_id}"
    reject_link = f"{base_url}/reject/{request_id}"

    subject = f"Leave Request Approval Needed: {staff_name}"
    html = f"""
    <h3>Leave Request Recommended</h3>
    <p>{staff_name} requested {leave_type} from {start_date} to {end_date}</p>
    <a href="{approve_link}" style="padding:10px;background:green;color:white;text-decoration:none;">Approve</a>
    <a href="{reject_link}" style="padding:10px;background:red;color:white;text-decoration:none;">Reject</a>
    """

    send_email(to=[pm_email], cc=[os_email, omtl_email], subject=subject, html=html)
    return "Leave recommended! Email sent to PM."


# PM Approve
@app.route("/approve/<request_id>")
def approve(request_id):
    leave_df = pd.read_excel(LEAVE_FILE, sheet_name=LEAVE_SHEET)
    idx = leave_df[leave_df["Request ID"] == request_id].index[0]
    leave_df.at[idx, "Status"] = "Approved"
    leave_df.to_excel(LEAVE_FILE, sheet_name=LEAVE_SHEET, index=False)

    # Notify staff, OS, OMTL
    master_df = load_master()
    outlet = leave_df.at[idx, "Outlet"]
    staff_email = leave_df.at[idx, "Staff Email"]
    staff_name = leave_df.at[idx, "Staff Name"]
    remaining_earned = leave_df.at[idx, "Remaining Earned"]
    remaining_sick = leave_df.at[idx, "Remaining Sick"]

    row = master_df[master_df["Outlet"] == outlet].iloc[0]
    os_email = row["OS Email"]
    omtl_email = row["OMTL Email"]

    subject = f"Leave Approved: {staff_name}"
    html = f"""
    <p>Your leave has been approved.</p>
    <p>Remaining Earned Leave: {remaining_earned}<br>
    Remaining Sick Leave: {remaining_sick}</p>
    """
    send_email(to=[staff_email], cc=[os_email, omtl_email], subject=subject, html=html)
    return "Leave approved and notification sent."


# PM Reject
@app.route("/reject/<request_id>")
def reject(request_id):
    leave_df = pd.read_excel(LEAVE_FILE, sheet_name=LEAVE_SHEET)
    idx = leave_df[leave_df["Request ID"] == request_id].index[0]
    leave_df.at[idx, "Status"] = "Rejected"
    leave_df.to_excel(LEAVE_FILE, sheet_name=LEAVE_SHEET, index=False)

    # Notify staff, OS, OMTL
    master_df = load_master()
    outlet = leave_df.at[idx, "Outlet"]
    staff_email = leave_df.at[idx, "Staff Email"]
    staff_name = leave_df.at[idx, "Staff Name"]
    remaining_earned = leave_df.at[idx, "Remaining Earned"]
    remaining_sick = leave_df.at[idx, "Remaining Sick"]

    row = master_df[master_df["Outlet"] == outlet].iloc[0]
    os_email = row["OS Email"]
    omtl_email = row["OMTL Email"]

    subject = f"Leave Rejected: {staff_name}"
    html = f"""
    <p>Your leave has been rejected.</p>
    <p>Remaining Earned Leave: {remaining_earned}<br>
    Remaining Sick Leave: {remaining_sick}</p>
    """
    send_email(to=[staff_email], cc=[os_email, omtl_email], subject=subject, html=html)
    return "Leave rejected and notification sent."


# Run server
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
