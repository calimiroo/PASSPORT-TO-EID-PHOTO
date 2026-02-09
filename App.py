import streamlit as st
import pandas as pd
import time
import json
import logging
from datetime import datetime, timedelta, date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import io
from PIL import Image, ImageDraw, ImageFont
import base64
import os
import tempfile

# --- Font Setup Function (Updated) ---
def get_font_path(font_name):
    """
    Returns the path to a font file, checking in order:
    1. Current directory (for local dev)
    2. ./fonts/ folder (for GitHub/Streamlit Cloud)
    3. System paths (fallback)
    4. Return None if nothing found -> PIL will use default
    """
    # 1. Check ./fonts/ folder
    fonts_dir = os.path.join(os.getcwd(), "fonts")
    font_path = os.path.join(fonts_dir, font_name)
    if os.path.exists(font_path):
        return font_path

    # 2. Check current directory
    if os.path.exists(font_name):
        return font_name

    # 3. Common system fallbacks (especially on Streamlit Cloud/Linux)
    system_fonts = {
        "Tajawal-Bold.ttf": [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        ],
        "Tajawal-Regular.ttf": [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ],
        # Generic fallbacks
        "arialbd.ttf": ["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"],
        "arial.ttf": ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"],
    }

    candidates = system_fonts.get(font_name, [])
    for path in candidates:
        if os.path.exists(path):
            return path

    # 4. If nothing found, return None -> PIL will use default
    return None

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- Page Setup ---
st.set_page_config(page_title="H-TRACING (ICP)", layout="wide")

# --- Password Protection ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
        <style>
        .big-title { font-size: 3.5rem; text-align: center; margin-top: 100px; color: #0d47a1; }
        .password-box { max-width: 400px; margin: 0 auto; text-align: center; margin-top: 50px; }
        </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="big-title">H-TRACING (ICP)</div>', unsafe_allow_html=True)
    st.markdown('<div style="text-align: center; font-size: 1.2rem; color: #555; margin-bottom: 40px;">Enter Password</div>', unsafe_allow_html=True)
    with st.container():
        st.markdown('<div class="password-box">', unsafe_allow_html=True)
        password = st.text_input("Password", type="password", label_visibility="collapsed")
        col1, col2, col3 = st.columns([1,1,1])
        with col2:
            if st.button("Enter", use_container_width=True):
                if password == "Hamada":
                    st.session_state.authenticated = True
                    st.success("Logged in successfully!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("Password Wrong")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# --- Main App ---
st.title("H-TRACING (ICP)")

# Table styling
st.markdown("""
    <style>
    .stTable td, .stTable th { white-space: nowrap !important; text-align: left !important; padding: 8px 15px !important; }
    .stTable { display: block !important; overflow-x: auto !important; }
    </style>
""", unsafe_allow_html=True)

# Session state
if 'run_state' not in st.session_state:
    st.session_state.run_state = 'stopped'
if 'batch_results' not in st.session_state:
    st.session_state.batch_results = []
if 'start_time_ref' not in st.session_state:
    st.session_state.start_time_ref = None
if 'single_result' not in st.session_state:
    st.session_state.single_result = None
if 'card_enlarged' not in st.session_state:
    st.session_state.card_enlarged = False

# Nationalities
countries_list = ["Select Nationality", "Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria", "Burkina Faso", "Burundi", "Cabo Verde", "Cambodia", "Cameroon", "Canada", "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo (Congo-Brazzaville)", "Costa Rica", "Côte d'Ivoire", "Croatia", "Cuba", "Cyprus", "Czechia (Czech Republic)", "Democratic Republic of the Congo", "Denmark", "Djibouti", "Dominica", "Dominican Republic", "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia", "Fiji", "Finland", "France", "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Holy See", "Honduras", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro", "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal", "Netherlands", "New Zealand", "Nicaragua", "Niger", "Nigeria", "North Korea", "North Macedonia", "Norway", "Oman", "Pakistan", "Palau", "Palestine State", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia", "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senin", "Serbia", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "South Korea", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", "Sweden", "Switzerland", "Syria", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo", "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States of America", "Uruguay", "Uzbekistan", "Vanuatu", "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe"]

def format_time(seconds):
    return str(timedelta(seconds=int(seconds)))

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def apply_styling(df):
    df.index = range(1, len(df)+1)
    def color_status(val):
        return 'background-color: #90EE90' if val == 'Found' else 'background-color: #FFCCCB'
    return df.style.applymap(color_status, subset=['Status'])

def reshape_arabic(text):
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        if text and any('\u0600' <= c <= '\u06FF' for c in text):
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        return text
    except ImportError:
        st.warning("Install: pip install arabic-reshaper python-bidi")
        return text

def format_date(date_str):
    if not date_str: return ''
    if 'T' in date_str: date_str = date_str.split('T')[0]
    for fmt in ['%Y-%m-%d', '%d/%m/%Y']:
        try:
            return datetime.strptime(date_str.strip(), fmt).strftime('%d/%m/%Y')
        except: pass
    return date_str

def wrap_text(draw, text, font, max_width):
    lines = []
    words = text.split(' ')
    current = ''
    for w in words:
        test = current + w + ' '
        if draw.textlength(test, font=font) <= max_width:
            current = test
        else:
            lines.append(current.strip())
            current = w + ' '
    if current: lines.append(current.strip())
    return lines

def create_card_image(data, size=(5760, 2700)):
    img = Image.new('RGB', size, (250,250,250))
    draw = ImageDraw.Draw(img)
    
    # Load Arabic-friendly fonts (with fallbacks)
    title_font_path = get_font_path("Tajawal-Bold.ttf")
    label_font_path = get_font_path("Tajawal-Regular.ttf")
    value_font_path = get_font_path("Tajawal-Regular.ttf")

    try:
        title_font = ImageFont.truetype(title_font_path or "DejaVuSans-Bold.ttf", 130)
    except:
        title_font = ImageFont.load_default()

    try:
        label_font = ImageFont.truetype(label_font_path or "DejaVuSans.ttf", 95)
    except:
        label_font = ImageFont.load_default()

    try:
        value_font = ImageFont.truetype(value_font_path or "DejaVuSans.ttf", 85)
    except:
        value_font = ImageFont.load_default()

    # Header
    draw.rectangle([(0,0), (size[0],150)], fill=(218,165,32))
    draw.text((120,40), "H-TRACING", fill=(0,0,139), font=title_font)

    # Photo area
    photo_x, photo_y = 180, 320
    photo_size = (950, 950)
    draw.rectangle([(photo_x, photo_y), (photo_x+photo_size[0], photo_y+photo_size[1])],
                   outline=(80,80,80), width=10, fill=(230,230,230))

    if 'Photo' in data and data['Photo']:
        try:
            photo_bytes = base64.b64decode(data['Photo'].split(',')[1])
            photo = Image.open(io.BytesIO(photo_bytes)).resize(photo_size, Image.LANCZOS)
            img.paste(photo, (photo_x, photo_y))
        except Exception as e:
            logger.warning(f"Photo load failed: {e}")
            draw.text((photo_x+120, photo_y+photo_size[1]//2-120), "YOUR\nPHOTO\nHERE",
                      fill=(120,120,120), font=title_font, align="center")
    else:
        draw.text((photo_x+120, photo_y+photo_size[1]//2-120), "YOUR\nPHOTO\nHERE",
                  fill=(120,120,120), font=title_font, align="center")

    # Labels & Values
    x_label, x_value = photo_x + photo_size[0] + 250, photo_x + photo_size[0] + 1850
    y = 350
    line_h = 135
    fields = [
        ("English Name:", "English Name"),
        ("Arabic Name:", "Arabic Name"),
        ("Unified Number:", "Unified Number"),
        ("EID Number:", "EID Number"),
        ("EID Expire Date:", "EID Expire Date"),
        ("Visa Issue Place:", "Visa Issue Place"),
        ("Profession:", "Profession"),
        ("English Sponsor Name:", "English Sponsor Name"),
        ("Arabic Sponsor Name:", "Arabic Sponsor Name"),
        ("Related Individuals:", "Related Individuals")
    ]

    max_w = size[0] - x_value - 200
    for lbl, key in fields:
        val = data.get(key, '')
        if key == "EID Expire Date": val = format_date(val)
        val_disp = reshape_arabic(str(val))
        draw.text((x_label, y), lbl, fill=(0,0,0), font=label_font)
        lines = wrap_text(draw, val_disp, value_font, max_w)
        for line in lines:
            draw.text((x_value, y), line, fill=(0,0,100), font=value_font)
            y += line_h / 1.8
        y += line_h - (len(lines)-1)*(line_h/1.8)

    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=98)
    buf.seek(0)
    return buf

class ICPScraper:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.url = "https://smartservices.icp.gov.ae/echannels/web/client/guest/index.html#/issueQrCode"

    def setup_driver(self):
        opts = webdriver.ChromeOptions()
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        opts.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)

        chrome_bin = "/usr/bin/chromium" if os.path.exists("/usr/bin/chromium") else None
        if chrome_bin:
            opts.binary_location = chrome_bin
            service = Service("/usr/bin/chromedriver") if os.path.exists("/usr/bin/chromedriver") else Service(ChromeDriverManager().install())
        else:
            service = Service(ChromeDriverManager().install())

        self.driver = webdriver.Chrome(service=service, options=opts)
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        self.wait = WebDriverWait(self.driver, 30)

    def safe_clear_and_fill(self, el, val):
        el.send_keys(Keys.CONTROL + "a")
        el.send_keys(Keys.BACKSPACE)
        time.sleep(0.3)
        el.send_keys(str(val))

    def select_from_dropdown(self, label, text):
        try:
            xpath = f"//label[contains(text(),'{label}')]/following::div[contains(@class,'ui-select-container')][1]"
            cont = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cont)
            cont.click()
            time.sleep(1)
            inp = self.wait.until(EC.visibility_of_element_located(
                (By.XPATH, f"//label[contains(text(),'{label}')]/following::input[not(@type='hidden')][1]")
            ))
            self.safe_clear_and_fill(inp, text)
            time.sleep(2)
            item = self.wait.until(EC.element_to_be_clickable((
                By.XPATH, f"//div[contains(@class,'ui-select-choices')]//span[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'{text.lower()}')]"
            )))
            self.driver.execute_script("arguments[0].click();", item)
            time.sleep(1)
        except Exception as e:
            logger.warning(f"Dropdown {label}: {e}")

    def capture_network_data(self):
        time.sleep(20)
        try:
            logs = self.driver.get_log('performance')
            for entry in reversed(logs):
                msg = json.loads(entry['message'])['message']
                if 'Network.responseReceived' in msg['method']:
                    req_id = msg['params'].get('requestId')
                    try:
                        body = self.driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': req_id})['body']
                        if 'isValid' in body:
                            data = json.loads(body)
                            if data.get('isValid'):
                                info = data.get('personalInfo', [{}])[0]
                                return {
                                    'English Name': info.get('englishFullName'),
                                    'Arabic Name': info.get('arabicFullName'),
                                    'Unified Number': info.get('unifiedNumber'),
                                    'EID Number': info.get('identityNumber'),
                                    'EID Expire Date': info.get('identityExpireDate'),
                                    'Visa Issue Place': info.get('englishIdentityIssuePlace'),
                                    'Profession': info.get('englishProfession'),
                                    'English Sponsor Name': info.get('englishSponsorName'),
                                    'Arabic Sponsor Name': info.get('arabicSponsorName'),
                                    'Status': 'Found'
                                }
                            else:
                                return {'Status': 'Not Found'}
                    except: pass
        except Exception as e:
            logger.error(f"Net err: {e}")
        return {'Status': 'Not Found'}

    def extract_qr_url(self):
        self.driver.execute_script("""
            if (!window.jsQR) {
                const s = document.createElement('script');
                s.src = 'https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js';
                document.head.appendChild(s);
            }
        """)
        time.sleep(3)
        return self.driver.execute_async_script("""
            const cb = arguments[arguments.length-1];
            const getQR = () => {
                let c = document.querySelector('canvas');
                if (c) return c;
                let i = document.querySelectorAll('img');
                for (let img of i) {
                    if (img.src && (img.src.includes('data:image') || img.src.includes('blob'))) return img;
                }
                return null;
            };
            const el = getQR();
            if (!el) return cb(null);
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            const img = new Image();
            img.crossOrigin = 'anonymous';
            img.src = el.toDataURL ? el.toDataURL() : el.src;
            img.onload = () => {
                canvas.width = img.width;
                canvas.height = img.height;
                ctx.drawImage(img, 0, 0);
                const d = ctx.getImageData(0,0,img.width,img.height);
                const code = window.jsQR(d.data, d.width, d.height);
                cb(code ? code.data : null);
            };
            img.onerror = () => cb(null);
        """)

    def perform_single_search(self, passport, nat, dob, gender):
        self.setup_driver()
        try:
            self.driver.get(self.url)
            logger.info(f"Searching: {passport}")
            time.sleep(3)

            self.driver.execute_script("""
                document.querySelector('input[value="personalInfo"]')?.click();
            """)
            time.sleep(2)

            self.select_from_dropdown('Current Nationality', nat)
            self.select_from_dropdown('Passport Type', 'ORDINARY PASSPORT')

            ppt = self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Passport Number')]/following::input[1]")))
            self.safe_clear_and_fill(ppt, passport)

            dob_f = pd.to_datetime(dob, dayfirst=True).strftime('%d/%m/%Y')
            dob_el = self.driver.find_element(By.XPATH, "//label[contains(text(),'Date of Birth')]/following::input[1]")
            self.safe_clear_and_fill(dob_el, dob_f)
            dob_el.send_keys(Keys.TAB)

            gen_el = self.driver.find_element(By.XPATH, "//label[contains(text(),'Gender')]/following::input[1]")
            self.safe_clear_and_fill(gen_el, gender)
            gen_el.send_keys(Keys.TAB)

            rel_el = self.driver.find_element(By.XPATH, "//label[contains(text(),'related to your file')]/following::input[1]")
            result = {'Status': 'Not Found'}
            rc = 0

            for attempt in [0,1,2,3,4,5]:
                rel_el.clear()
                rel_el.send_keys(str(attempt))
                rel_el.send_keys(Keys.TAB)
                time.sleep(1)
                btn = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[ng-click='search()']")))
                self.driver.execute_script("arguments[0].removeAttribute('disabled'); arguments[0].click();", btn)
                time.sleep(7)

                net_res = self.capture_network_data()
                if net_res.get('Status') == 'Found':
                    result = net_res
                    rc = attempt
                    break

            if result.get('Status') == 'Found':
                result.update({
                    'Passport Number': passport,
                    'Nationality': nat,
                    'Gender': gender,
                    'Related Individuals': str(rc),
                    'EID Expire Date': format_date(result.get('EID Expire Date', '')),
                })

                # --- استخراج QR ثم الصورة من بطاقة الهوية ---
                qr = self.extract_qr_url()
                if qr:
                    logger.info("Opening QR page...")
                    self.driver.get(qr)
                    time.sleep(8)

                    # استخراج من canvas أولاً
                    photo_b64 = self.driver.execute_script("""
                        let src = null;
                        const canvas = document.querySelector('canvas');
                        if (canvas) src = canvas.toDataURL('image/jpeg', 0.9);
                        if (!src) {
                            const imgs = document.querySelectorAll('img');
                            for (let img of imgs) {
                                const s = img.src;
                                if (s && s.startsWith('data:image') && s.includes('base64')) {
                                    src = s;
                                    break;
                                }
                            }
                        }
                        return src;
                    """)
                    if photo_b64 and photo_b64.startswith('data:image'):
                        result['Photo'] = photo_b64
                        logger.info("✅ Photo extracted from QR page.")

            return result

        except Exception as e:
            logger.error(f"Search error: {e}")
            return {'Passport Number': passport, 'Status': 'Error'}
        finally:
            if self.driver:
                self.driver.quit()

def toggle_card():
    st.session_state.card_enlarged = not st.session_state.card_enlarged

tab1, tab2 = st.tabs(["Single Search", "Batch"])

with tab1:
    st.subheader("بحث فردي")
    c1,c2,c3 = st.columns(3)
    p = c1.text_input("رقم الجواز", key="sp")
    n = c2.selectbox("الجنسية", countries_list, key="sn")
    # --- التحديث: إضافة max_value وتنبيه ---
    d = c3.date_input(
        "تاريخ الميلاد",
        value=None,
        min_value=datetime(1900, 1, 1),
        max_value=date.today(),  # <-- هذا هو التحديث المهم
        format="DD/MM/YYYY",
        key="sd"
    )
    if d and d > date.today():
        st.warning("⚠️ لا يمكن أن يكون تاريخ الميلاد في المستقبل!")

    g = st.radio("الجنس", ["Male", "Female"], key="sg")

    b1,b2,b3 = st.columns(3)
    with b1:
        if st.button("بحث الآن", key="go"):
            if p and n != "Select Nationality" and d:
                with st.spinner("جاري البحث..."):
                    scraper = ICPScraper()
                    res = scraper.perform_single_search(p, n, d, "1" if g=="Male" else "0")
                    st.session_state.single_result = res

    with b2:
        if st.button("⏹ توقف"):
            st.session_state.single_result = None
            st.rerun()

    with b3:
        if st.button("🗑 مسح"):
            st.session_state.single_result = None
            st.rerun()

    if st.session_state.single_result:
        cols = ['English Name','Arabic Name','Unified Number','EID Number','EID Expire Date',
                'Visa Issue Place','Profession','English Sponsor Name','Arabic Sponsor Name',
                'Related Individuals','Status']
        df = pd.DataFrame([{k:v for k,v in st.session_state.single_result.items() if k in cols}])
        st.table(apply_styling(df))

        if st.session_state.single_result.get('Status') == 'Found':
            card_buf = create_card_image(st.session_state.single_result)
            w = 1400 if st.session_state.card_enlarged else 700
            st.image(card_buf, caption="البطاقة", width=w)
            st.button("تكبير" if not st.session_state.card_enlarged else "تصغير", on_click=toggle_card)
            st.download_button("📥 حفظ البطاقة", card_buf, f"card_{st.session_state.single_result.get('Unified Number')}.jpg", "image/jpeg")

with tab2:
    st.subheader("معالجة جماعية")
    up = st.file_uploader("رفع ملف Excel", type=["xlsx"])
    if up:
        df = pd.read_excel(up)
        st.dataframe(df.head(), height=150)
        c1,c2,c3 = st.columns(3)
        if c1.button("▶ ابدأ"):
            st.session_state.run_state = 'running'
            st.session_state.start_time_ref = time.time()
        if c2.button("⏸ توقف مؤقت"):
            st.session_state.run_state = 'paused'
        if c3.button("⏹ إلغاء"):
            st.session_state = {k:v for k,v in st.session_state.items() if k in ['authenticated']}
            st.rerun()

        prog = st.progress(0)
        stat = st.empty()
        live = st.empty()
        success = 0

        for i, row in df.iterrows():
            while st.session_state.run_state == 'paused':
                stat.warning("متوقف مؤقتًا...")
                time.sleep(1)
            if st.session_state.run_state == 'stopped': break

            p = str(row.get('Passport Number','')).strip()
            n = str(row.get('Nationality','Egypt')).strip()
            try:
                dob_row = row['Date of Birth']
                # --- التحديث: التأكد من أن التاريخ لا يكون في المستقبل ---
                if pd.isna(dob_row):
                    continue
                dob_dt = pd.to_datetime(dob_row)
                if dob_dt.date() > date.today():
                    st.warning(f"⚠️ تاريخ الميلاد في السطر {i+1} في المستقبل. تخطي.")
                    continue
                d_formatted = dob_dt.strftime('%d/%m/%Y')
            except:
                d_formatted = ''

            g = str(row.get('Gender', '1')).strip()

            stat.info(f"المعالجة: {i+1}/{len(df)} | {p}")
            scraper = ICPScraper()
            res = scraper.perform_single_search(p, n, d_formatted, g)
            st.session_state.batch_results.append(res)
            if res.get('Status') == 'Found': success += 1

            elapsed = time.time() - st.session_state.start_time_ref
            stat.markdown(f"✅ نجاح: {success} | ⏱ زمن: {format_time(elapsed)}")

            # عرض النتائج الحية
            filtered = [{k:v for k,v in r.items() if k in cols} for r in st.session_state.batch_results]
            live.table(apply_styling(pd.DataFrame(filtered)))
            prog.progress((i+1)/len(df))

        if len(st.session_state.batch_results) == len(df):
            st.success("انتهى البحث!")
            final_df = pd.DataFrame([{k:v for k,v in r.items() if k in cols} for r in st.session_state.batch_results])
            excel = to_excel(final_df)
            st.download_button("📥 تنزيل النتائج", excel, "نتائج_البحث.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
