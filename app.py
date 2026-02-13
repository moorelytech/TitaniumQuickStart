import os
import re
import socket
import threading
import webbrowser
from pathlib import Path
from datetime import datetime

from nicegui import ui

from ps_runner import run_ps_function, run_ps

APP_TITLE = "Titanium QuickStart (Web App)"
BASE_DIR = Path(__file__).resolve().parent
PS1_PATH = str(BASE_DIR / "TitaniumQuickStart_5.0.ps1")


# --------------------------
# Utilities
# --------------------------
def pick_free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    _, port = s.getsockname()
    s.close()
    return port


def ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


log_lock = threading.Lock()


def append_log(area: ui.textarea, msg: str) -> None:
    with log_lock:
        current = area.value or ""
        area.value = (current + ("\n" if current else "") + f"[{ts()}] {msg}").strip()
        area.update()


def ps_escape_single_quotes(s: str) -> str:
    # In PowerShell single-quoted strings: escape single quote by doubling it
    return (s or "").replace("'", "''")


def get_default_sdm_name() -> str:
    # Windows: USERNAME is typically present
    return os.environ.get("USERNAME") or os.environ.get("USER") or "Unknown User"


# --------------------------
# PowerShell bridge: build "WinForms-like" variables as PSCustomObjects
# --------------------------
def build_ps_prelude(
    *,
    sales_order: str,
    sdm_name: str,
    cust_name: str,
    short_desc: str,
    prod: bool,
    draas: bool,
    baas: bool,
    copy_sedt: bool,
    teams_emails: list[str],
    teams_roles: list[str],
    # Optional "advanced" values used by your OneNote page XML
    time_tracking_case: str = "",
    external_teams_link: str = "",
    primary_dc: str = "",
    draas_dc: str = "",
    target_completion: str = "",
    dr_rehearsal: str = "",
    migration_cutover: str = "",
    # Optional PI URL override
    pi_url_override: str = "",
    # Optional PI API key (if you want to enable PI lookups from web)
    pi_api_key: str = "",
) -> str:
    # Normalize to lengths you used in PS (customer name max ~45, short desc max 25)
    cust_name = cust_name.strip()
    short_desc = short_desc.strip()[:25]

    # Ensure the list is exactly 6 items
    emails = (teams_emails + [""] * 6)[:6]
    roles = (teams_roles + ["Member"] * 6)[:6]

    so = ps_escape_single_quotes(sales_order.strip())
    sdm = ps_escape_single_quotes(sdm_name.strip())
    cust = ps_escape_single_quotes(cust_name)
    desc = ps_escape_single_quotes(short_desc)

    # Advanced
    time_tracking_case = ps_escape_single_quotes(time_tracking_case.strip())
    external_teams_link = ps_escape_single_quotes(external_teams_link.strip())
    primary_dc = ps_escape_single_quotes(primary_dc.strip())
    draas_dc = ps_escape_single_quotes(draas_dc.strip())
    target_completion = ps_escape_single_quotes(target_completion.strip())
    dr_rehearsal = ps_escape_single_quotes(dr_rehearsal.strip())
    migration_cutover = ps_escape_single_quotes(migration_cutover.strip())
    pi_url_override = ps_escape_single_quotes(pi_url_override.strip())
    pi_api_key = ps_escape_single_quotes(pi_api_key.strip())

    emails_ps = [ps_escape_single_quotes(x.strip()) for x in emails]
    roles_ps = [ps_escape_single_quotes((x or "Member").strip()) for x in roles]

    # We “mock” the controls/labels/buttons that your PS code references.
    # If a function tries to set .Text or .Enabled, it won’t crash.
    prelude = f"""
# --------------------------
# WebUI Prelude: Mock WinForms controls + globals expected by the PS1 functions
# --------------------------

# Required variables
$global:SoNum = '{so}'
$global:CustName = '{cust}'
$global:SdmName = '{sdm}'

# Optional fields used in logic / UI
$global:ProjectShortDesc = '{desc}'

# Advanced OneNote fields (your XML references these)
$global:TimeTrackingCaseNumber = '{time_tracking_case}'
$global:ExternalTeamsLink       = '{external_teams_link}'
$global:PrimaryDatacenter        = '{primary_dc}'
$global:DRaaSDatacenter          = '{draas_dc}'
$global:TargetCompletionDate     = '{target_completion}'
$global:DRRehearsalDate          = '{dr_rehearsal}'
$global:MigrationCutoverDate     = '{migration_cutover}'

# If you want PI lookups in web mode:
$global:SdmQuickStartPiKey = '{pi_api_key}'

# ---- Mock "controls" (TextBox / MaskedTextBox / CheckBox / Button / ComboBox / Label) ----
function New-MockTextBox([string]$text) {{
    return [pscustomobject]@{{ Text = $text; ReadOnly = $false }}
}}
function New-MockCheckBox([bool]$checked, [bool]$enabled=$true, [string]$text='') {{
    return [pscustomobject]@{{ Checked = $checked; Enabled = $enabled; Text = $text }}
}}
function New-MockButton([string]$text='') {{
    return [pscustomobject]@{{ Text = $text; Enabled = $true; ForeColor = ''; Visible = $true }}
}}
function New-MockLabel([string]$text='') {{
    return [pscustomobject]@{{ Text = $text; ForeColor = ''; Visible = $true }}
}}
function New-MockCombo([string]$selected) {{
    return [pscustomobject]@{{ SelectedItem = $selected }}
}}

# Textboxes from your script
$SalesOrderTextBox    = New-MockTextBox('{so}')
$CustomerNameTextBox  = New-MockTextBox('{cust}')
$ProjectShortDescTextBox = New-MockTextBox('{desc}')
$SdmNameTextBox       = New-MockTextBox('{sdm}')
$PINumberTextBox      = New-MockTextBox('')
$PMNameTextBox        = New-MockTextBox('')
$TeamsChannelNameTextBox = New-MockTextBox("EXTERNAL - {cust} - {so}")

# Checkboxes (Workbook type / copy SEDT)
$ProdCheckBox  = New-MockCheckBox({str(prod).lower()}, $true)
$DRaaSCheckBox = New-MockCheckBox({str(draas).lower()}, $true)
$BaaSCheckBox  = New-MockCheckBox({str(baas).lower()}, $true)
$CopySEDTCheckBox = New-MockCheckBox({str(copy_sedt).lower()}, $true)

# Labels / status
$SOValid       = New-MockLabel('')
$PIStatusLabel = New-MockLabel('')

# Buttons referenced by logic
$OpenInstallLocationButton = New-MockButton('Browse Install Documents Now')
$CreateSCButton            = New-MockButton('Create Shortcut to KP')
$CreateONButton            = New-MockButton('Create OneNote Page')
$BrowsePIButton            = New-MockButton('Browse Project Insight')
$SubmitButton              = New-MockButton('Submit')

# Teams tab inputs (6)
$TeamsEmailTextBox1 = New-MockTextBox('{emails_ps[0]}')
$TeamsEmailTextBox2 = New-MockTextBox('{emails_ps[1]}')
$TeamsEmailTextBox3 = New-MockTextBox('{emails_ps[2]}')
$TeamsEmailTextBox4 = New-MockTextBox('{emails_ps[3]}')
$TeamsEmailTextBox5 = New-MockTextBox('{emails_ps[4]}')
$TeamsEmailTextBox6 = New-MockTextBox('{emails_ps[5]}')

$TeamsRoleComboBox1 = New-MockCombo('{roles_ps[0]}')
$TeamsRoleComboBox2 = New-MockCombo('{roles_ps[1]}')
$TeamsRoleComboBox3 = New-MockCombo('{roles_ps[2]}')
$TeamsRoleComboBox4 = New-MockCombo('{roles_ps[3]}')
$TeamsRoleComboBox5 = New-MockCombo('{roles_ps[4]}')
$TeamsRoleComboBox6 = New-MockCombo('{roles_ps[5]}')

# PI URL override (so Browse PI can work even without CheckProjectInsight)
if ('{pi_url_override}' -ne '') {{
    $global:ProjectInsightUrl = '{pi_url_override}'
}}

# Some functions reference TabControl1.SelectedTab; we provide a harmless stub
$TabControl1 = [pscustomobject]@{{ SelectedTab = $null }}
$MainTab = $null
"""
    return prelude.strip()


def run_async(area: ui.textarea, fn: str, prelude: str = "") -> None:
    append_log(area, f"> {fn} (prelude: {'yes' if prelude.strip() else 'no'})")

    def worker():
        try:
            code, out, err = run_ps_function(PS1_PATH, fn, prelude=prelude)
            if out.strip():
                append_log(area, out.strip())
            if err.strip():
                append_log(area, "STDERR:\n" + err.strip())
            append_log(area, f"> Exit code: {code}")
        except Exception as e:
            append_log(area, f"EXCEPTION: {e}")

    threading.Thread(target=worker, daemon=True).start()


def open_browser(port: int) -> None:
    url = f"http://127.0.0.1:{port}/"
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()


# --------------------------
# UI
# --------------------------
HELP_TEXT = """This tool allows the executing user to quickly prepare for an internal alignment and project introduction meeting with the customer shortly after receiving project assignment. Once executed, it will create the prerequisite documents, external customer facing Teams channel, and internal SDM Team shared OneNote section.

The user is to input two required fields:
A customer name: This textbox input will be used in the naming convention of a local Implementations folder on the user’s local OneDrive.
The sales order number (ex: T12345678). This textbox input must be a valid sales order number, it will be searched for in the TierPoint 'Knowledgepoint – Documents' folder and if it exists, a shortcut to the folder will be created in the user’s local OneDrive (\\Implementations\\<Customer Name Input>\\<Sales Order #> - <Short Description>\\)
The SDM field is populated based on the user who executed the script. Note, at this time, this should only be an SDM due to the nature in which the code is written.

There are three optional fields:
The Project Manager assigned to the project: Optional as it may not be known yet who is taking ownership of the project
The Project Insight number: Optional as the project may not yet be created in PI
A short description: Optional and will be incorporated into the user’s local folder naming schema in case there are multiple orders for the same customer name. Helps keep organization

Quick features:
With only the sales order number entered, the user can quickly open the Knowledge Point install documents folder via the 'Open Install Documents' button to review project scope and sales order details through all available documents found in KP including the actual sales order or the SSD. This is the same folder location that is linked to the project within Project Insight. Check if the SE Discovery Toolkit file exists. If the file exists, the script can pull the server list and associated server information into the workbook via the optional checkbox that is only visible if the file exists.
"""


@ui.page("/")
def main():

    # --- Authentication screen ---
    ui.label(APP_TITLE).classes("text-2xl font-bold")
    ui.label(f"Script: {PS1_PATH}").classes("text-sm opacity-70")

    if not os.path.exists(PS1_PATH):
        ui.markdown("⚠️ `TitaniumQuickStart_5.0.ps1` not found next to `app.py`.")
        return

    auth_status = {'done': False, 'error': None}
    # --- Authentication modal dialog ---
    credentials = {'username': '', 'password': ''}
    auth_status = {'done': False, 'error': None}

    def authenticate_user():
        def auth_worker():
            import subprocess
            try:
                username = username_input.value.strip()
                password = password_input.value.strip()
                credentials['username'] = username
                credentials['password'] = password
                
                if not username or not password:
                    auth_status['error'] = "Username and password are required."
                    error_label.set_text(auth_status['error'])
                    return
                
                # Simple connectivity check to CA (matching your PS1 logic)
                ps_code = """
$CAUrl = "https://ca.tierpoint.com"
try {
    $response = Invoke-WebRequest -Uri $CAUrl -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop
    if ($response.StatusCode -eq 200) {
        Write-Output "success"
    } else {
        Write-Output "Error: CA returned status $($response.StatusCode)"
    }
} catch {
    Write-Output "Error: $($_.Exception.Message)"
}
"""
                cmd = ["powershell", "-ExecutionPolicy", "Bypass", "-NonInteractive", "-NoProfile", "-Command", ps_code]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=15, stdin=subprocess.DEVNULL)
                output = result.stdout.strip()
                error = result.stderr.strip()
                
                # Parse output
                if "success" in output.lower():
                    auth_status['done'] = True
                    login_dialog.close()
                    error_label.set_text("")
                else:
                    auth_status['error'] = output or error or "Cannot connect to CyberArk. Please check your VPN connection."
                    error_label.set_text(auth_status['error'])
                    
            except subprocess.TimeoutExpired:
                auth_status['error'] = "Connection timeout. Check your network/VPN."
                error_label.set_text(auth_status['error'])
            except Exception as e:
                auth_status['error'] = f"Error: {str(e)}"
                error_label.set_text(auth_status['error'])
        # Run authentication in background thread to avoid blocking event loop
        threading.Thread(target=auth_worker, daemon=True).start()

    login_dialog = ui.dialog()
    with login_dialog:
        with ui.card().classes("w-96"):
            ui.label("Login Required").classes("text-lg font-bold")
            username_input = ui.input("Username").classes("w-full")
            password_input = ui.input("Password").props("type=password").classes("w-full")
            login_btn = ui.button("Login").classes("w-full mt-4")
            error_label = ui.label().classes("text-red-600 text-sm mt-2")

    # Explicitly set login button callback after dialog creation
    login_btn.on_click(authenticate_user)

    def poll_auth():
        if auth_status['done']:
            show_main_ui()
            poll_timer.deactivate()
        elif auth_status['error']:
            error_label.set_text(auth_status['error'])

    poll_timer = ui.timer(0.5, poll_auth)
    login_dialog.open()

    def check_network():
        import requests
        try:
            resp = requests.get("https://ca.tierpoint.com", timeout=5)
            if resp.status_code == 200:
                auth_status['done'] = True
            else:
                auth_status['error'] = f"Network authentication failed: Status {resp.status_code}"
        except Exception as e:
            auth_status['error'] = f"Network authentication error: {e}"
        ui.update()

    threading.Thread(target=check_network, daemon=True).start()

    def show_main_ui():
        # --- Define all UI variables at the top ---
        log_area = ui.textarea("Logs").props("readonly").classes("w-full h-60")
        sdm_default = get_default_sdm_name()
        so_input = ui.input("Sales Order #").props('placeholder="T12345678"').classes("w-56")
        sdm_input = ui.input("SDM (You)").props("readonly").classes("w-72")
        sdm_input.value = sdm_default
        prod_cb = ui.checkbox("Production")
        draas_cb = ui.checkbox("DRaaS")
        baas_cb = ui.checkbox("BaaS (Ded)")
        cust_input = ui.input("Customer Name").classes("w-full")
        short_desc_input = ui.input("Short Description").props("maxlength=25").classes("w-full")
        pi_num_input = ui.input("Project Insight #").props("readonly").classes("w-56")
        pm_input = ui.input("Project Manager").props("readonly").classes("w-72")
        copy_sedt_cb = ui.checkbox("Copy server list from SE Discovery Toolkit to workbook?")
        copy_sedt_cb.disable()
        pi_url_input = ui.input("Project Insight URL (optional)").props("clearable").classes("w-full")
        tt_case = ui.input("Time Tracking Case #").classes("w-full")
        ext_teams = ui.input("Link to External Teams").classes("w-full")
        primary_dc = ui.input("Primary Datacenter").classes("w-full")
        draas_dc = ui.input("DRaaS Datacenter").classes("w-full")
        target_completion = ui.input("Target Completion Date").classes("w-full")
        dr_rehearsal = ui.input("DR Rehearsal Date").classes("w-full")
        migration_cutover = ui.input("Migration Cutover Date").classes("w-full")
        pi_api_key = ui.input("PI API Key").classes("w-full")
        teams_emails = [ui.input(f"Email {i+1}").props('placeholder="user@tierpoint.com"').classes("w-full") for i in range(6)]
        teams_roles = [ui.select(["Owner", "Member"], value="Member").classes("w-40") for _ in range(6)]
        teams_channel_name = ui.input("Teams Channel Name").props("readonly").classes("w-full")
        so_regex = re.compile(r"^T\d{8}$", re.IGNORECASE)

        # --- Helper functions ---
        def ensure_sales_order_ok():
            so_val = so_input.value or ""
            if not re.match(r"^T\d{8}$", so_val):
                append_log(log_area, "Sales order must be 'T' followed by 8 digits (ex: T12345678).")
                return False
            return True

        def run_ps_action(action):
            def ps_worker():
                import subprocess
                import json
                import re
                ps_script = PS1_PATH
                args = [
                    '-SalesOrder', so_input.value or '',
                    '-SDMName', sdm_input.value or '',
                    '-CustomerName', cust_input.value or '',
                    '-ShortDesc', short_desc_input.value or '',
                    '-Prod', str(bool(prod_cb.value)),
                    '-DRaaS', str(bool(draas_cb.value)),
                    '-BaaS', str(bool(baas_cb.value)),
                    '-CopySEDT', str(bool(copy_sedt_cb.value)),
                    '-TeamsEmails', ','.join([t.value or '' for t in teams_emails]),
                    '-TeamsRoles', ','.join([r.value or 'Member' for r in teams_roles]),
                    '-Action', action
                ]
                cmd = ['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps_script] + args
                try:
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                    output = result.stdout
                    error = result.stderr.strip()
                    
                    # If ValidateSalesNumber, extract and parse JSON data
                    if action == 'ValidateSalesNumber' and '===BEGIN_PI_DATA===' in output:
                        # Extract JSON between markers
                        match = re.search(r'===BEGIN_PI_DATA===\s*(.*?)\s*===END_PI_DATA===', output, re.DOTALL)
                        if match:
                            json_str = match.group(1).strip()
                            try:
                                pi_data = json.loads(json_str)
                                # Populate UI fields with extracted data
                                if pi_data.get('CustomerName'):
                                    cust_input.set_value(pi_data['CustomerName'])
                                if pi_data.get('PINumber'):
                                    pi_num_input.set_value(pi_data['PINumber'])
                                if pi_data.get('PMName'):
                                    pm_input.set_value(pi_data['PMName'])
                                append_log(log_area, f"✓ Project Insight data loaded: {pi_data.get('CustomerName')} - {pi_data.get('PINumber')}")
                            except json.JSONDecodeError as je:
                                append_log(log_area, f'Error parsing Project Insight data: {je}')
                    
                    # Log non-marker output
                    log_output = re.sub(r'===BEGIN_PI_DATA===.*?===END_PI_DATA===', '[PI DATA LOADED]', output, flags=re.DOTALL).strip()
                    if log_output:
                        append_log(log_area, log_output)
                    
                    if error:
                        append_log(log_area, f'STDERR:\n{error}')
                except Exception as e:
                    append_log(log_area, f'Error executing {action}: {e}')
            threading.Thread(target=ps_worker, daemon=True).start()


        # --- Event handlers ---
        def validate_sales_order():
            if not ensure_sales_order_ok():
                return
            run_ps_action('ValidateSalesNumber')

        def submit():
            if not ensure_sales_order_ok():
                return
            if not (cust_input.value or '').strip():
                append_log(log_area, 'Customer name is required.')
                return
            if not (prod_cb.value or draas_cb.value or baas_cb.value):
                append_log(log_area, 'Select at least one Workbook Type (Production / DRaaS / BaaS).')
                return
            run_ps_action('SdmSubmit')

        def open_install_docs():
            if not ensure_sales_order_ok():
                return
            run_ps_action('OpenInstallLocation')

        def create_shortcut():
            if not ensure_sales_order_ok():
                return
            run_ps_action('CreateSCButton')

        def create_onenote():
            if not ensure_sales_order_ok():
                return
            if not (cust_input.value or '').strip():
                append_log(log_area, 'Customer name is required before creating OneNote content.')
                return
            run_ps_action('CreateOneNoteSection')

        def browse_pi():
            run_ps_action('BrowsePI')

        def create_team():
            run_ps_action('CreateTeam')

        # --- UI Layout (No Tabs) ---
        ui.label("Titanium QuickStart").classes("text-2xl font-bold")
        
        # --- Top Action Buttons ---
        with ui.card().classes("w-full"):
            ui.label("Actions").classes("font-semibold mb-2")
            with ui.row().classes("w-full gap-3 flex-wrap"):
                ui.button("Create OneNote Page", on_click=create_onenote).classes("flex-1")
                ui.button("Create Shortcut to KP", on_click=create_shortcut).classes("flex-1")
                ui.button("Submit", on_click=submit).classes("flex-1")
                ui.button("Browse Project Insight", on_click=browse_pi).classes("flex-1")
                ui.button("Browse Install Documents", on_click=open_install_docs).classes("flex-1")
                ui.button("Create Team", on_click=create_team).classes("flex-1")
                ui.button("Exit", color="negative", on_click=lambda: os._exit(0)).classes("flex-1")
        
        # --- Main Form Sections ---
        with ui.card().classes("w-full"):
            ui.label("Required Input").classes("font-semibold mb-3")
            
            # Sales Order row with Validate button and checkboxes
            with ui.row().classes("w-full items-center gap-4"):
                so_input.classes("w-56")
                prod_cb
                draas_cb
                baas_cb
                ui.button("Validate", on_click=validate_sales_order).props("outline")
            
            # SDM input
            with ui.row().classes("w-full items-center gap-4 mt-4"):
                sdm_input
        
        with ui.card().classes("w-full"):
            ui.label("Optional Fields").classes("font-semibold mb-3")
            cust_input
            short_desc_input
            with ui.row().classes("w-full gap-6"):
                pi_num_input
                pm_input
            copy_sedt_cb
            with ui.row().classes("w-full gap-4 mt-4"):
                pi_url_input.classes("flex-1")
        
        with ui.expansion("Advanced", icon="tune").classes("w-full"):
            tt_case
            ext_teams
            primary_dc
            draas_dc
            target_completion
            dr_rehearsal
            migration_cutover
            pi_api_key
        
        # --- Teams Section ---
        with ui.expansion("Teams", icon="people").classes("w-full"):
            ui.label('Enter email addresses and roles for team members').classes("text-sm mb-3")
            with ui.card().classes("w-full"):
                for i in range(6):
                    with ui.row().classes("w-full items-center gap-4"):
                        teams_emails[i].classes("flex-1")
                        teams_roles[i]
            ui.separator()
            teams_channel_name
        
        # --- Help Section ---
        with ui.expansion("Help", icon="help").classes("w-full"):
            with ui.card().classes("w-full"):
                ui.textarea("Help").props("readonly").classes("w-full h-96").set_value(HELP_TEXT)
        
        # --- Logs at Bottom ---
        ui.separator()
        with ui.row().classes("w-full justify-between items-center mb-2"):
            ui.button("Clear logs", on_click=lambda: (setattr(log_area, "value", ""), log_area.update())).props("outline")
            ui.label("Tip: Click Validate after entering Sales Order").classes("text-xs opacity-70")
        log_area



if __name__ in {"__main__", "__mp_main__"}:
    port = pick_free_port()
    open_browser(port)
    ui.run(host="127.0.0.1", port=port, reload=False, show=False)
