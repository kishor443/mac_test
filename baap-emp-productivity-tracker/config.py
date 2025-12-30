APP_NAME = "ProductivityTracker"
APP_VERSION = "2.0.0"


# Environments
ENVIRONMENTS = {
    "prod": "https://vjlbumzku2.execute-api.ap-south-1.amazonaws.com",
    "qa": "https://m8p03gbszk.execute-api.ap-south-1.amazonaws.com"
}
 
# Select environment: "prod" or "qa"

# CURRENT_ENV = "qa"
CURRENT_ENV = "prod"

BASE_URL = ENVIRONMENTS[CURRENT_ENV] + "/auth/api"


# ERP API Endpoints (update URLs as needed)
ERP_ATTENDANCE_URL = f"{BASE_URL}/attendance/"
ERP_CLIENTS_URL = f"{BASE_URL}/clients"
ERP_SHIFTS_BASE_URL = f"{BASE_URL}/shifts/client"
ERP_REFRESH_TOKEN_URL = f"{BASE_URL}/auth/refresh-token"
# Login endpoint: /auth/api/auth/login (note: /auth/api + /auth/login = /auth/api/auth/login)
ERP_CREDENTIAL_LOGIN_URL = f"{BASE_URL}/auth/login"
ERP_LOGIN_URL = f"{BASE_URL}/auth/verify-otp"
ERP_REQUEST_OTP_URL = f"{BASE_URL}/auth/request-otp"
ERP_VERIFY_OTP_URL = f"{BASE_URL}/auth/verify-otp"

# Notices
ERP_NOTICES_URL = f"{BASE_URL}/notices/client"

# Appointments
# If using a different base URL for appointments, set it here
# Otherwise, it will use BASE_URL
# Example: "https://dt1wp7hrm9.execute-api.ap-south-1.amazonaws.com/auth/api"
ERP_APPOINTMENT_BASE_URL = None  # Set to None to use BASE_URL, or set to custom URL


ERP_DEVICE_ID = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36"  # (Store user device)

# Default headers sometimes required by ERP (can be adjusted)
# Note: Origin/Referer should match the web app URL for CORS
ERP_DEFAULT_HEADERS = {
    "accept": "application/json, text/plain, */*",
    "origin": "https://pro-hance-productivity-tracker--ajinkyaw443.replit.app",
    "referer": "https://pro-hance-productivity-tracker--ajinkyaw443.replit.app/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36",
}

# Login API specific headers (matching platform.baap.company)
ERP_LOGIN_HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9,en-IN;q=0.8",
    "content-type": "application/json",
    "origin": "https://platform.baap.company",
    "referer": "https://platform.baap.company/",
    "sec-ch-ua": '"Chromium";v="142", "Microsoft Edge";v="142", "Not_A Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "cross-site",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36 Edg/142.0.0.0",
}

# Optional: if your ERP requires a shift id for clock-in, set it here
ERP_DEFAULT_SHIFT_ID = ""

# Activity/Idle Settings
IDLE_TIMEOUT_SECONDS = 5 * 60
IDLE_CHECK_INTERVAL_SECONDS = 10


# Activity analytics
ACTIVITY_SAMPLE_SECONDS = 5  # how often UI reports activity to session manager
ACTIVITY_SCORE_WINDOW_SECONDS = 60  # compute % active over last N seconds

# Screenshots - increased interval for better performance
SCREENSHOT_INTERVAL_SECONDS = 15 * 60 # take screenshot every 15 minutes while clocked in (reduced CPU usage)
SCREENSHOTS_DIR = "data/screenshots"
WEBCAM_PHOTOS_DIR = "data/webcam"
WEBCAM_DEVICE_NAME = "USB2.0 HD UVC WebCam"
MAX_BROWSER_TABS_CAPTURED = 50

# Data retention
DATA_RETENTION_DAYS = 30

# Worklog / telemetry uploads
ERP_WORKLOG_URL = f"{BASE_URL}/worklog"

# Networking
NETWORK_CHECK_INTERVAL_SECONDS = 30

# Data Storage
DATA_DIR = "data"
LOCAL_STORAGE_FILE = f"{DATA_DIR}/local_storage.json"
EXCEL_ACTIVITY_FILE = f"{DATA_DIR}/activity_log.xlsx"
EXCEL_ACTIVITY_SHEET = "ActivityLog"
EXCEL_LOCAL_STORAGE_SHEET = "LocalStorage"
LOG_FILE = f"{DATA_DIR}/app.log"