import streamlit as st
import pandas as pd
import time
import json
import logging
from datetime import datetime, timedelta
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

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- Page Setup ---
st.set_page_config(page_title="ICP Data Search", layout="wide")

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
        if st.button("Enter", use_container_width=True):
            if password == "Hamada":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Password Wrong")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

st.title("H-TRACING (ICP)")

# --- Styling ---
st.markdown("""
    <style>
    .stTable td, .stTable th { white-space: nowrap !important; text-align: left !important; padding: 8px 15px !important; }
    .stTable { display: block !important; overflow-x: auto !important; }
    </style>
    """, unsafe_allow_html=True)

# --- Session State ---
if 'run_state' not in st.session_state: st.session_state.run_state = 'stopped'
if 'batch_results' not in st.session_state: st.session_state.batch_results = []
if 'start_time_ref' not in st.session_state: st.session_state.start_time_ref = None
if 'single_result' not in st.session_state: st.session_state.single_result = None
if 'card_enlarged' not in st.session_state: st.session_state.card_enlarged = False

countries_list = ["Select Nationality", "Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria", "Burkina Faso", "Burundi", "Cabo Verde", "Cambodia", "Cameroon", "Canada", "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo (Congo-Brazzaville)", "Costa Rica", "Côte d'Ivoire", "Croatia", "Cuba", "Cyprus", "Czechia (Czech Republic)", "Democratic Republic of the Congo", "Denmark", "Djibouti", "Dominica", "Dominican Republic", "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia", "Fiji", "Finland", "France", "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Holy See", "Honduras", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro", "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal", "Netherlands", "New Zealand", "Nicaragua", "Niger", "Nigeria", "North Korea", "North Macedonia", "Norway", "Oman", "Pakistan", "Palau", "Palestine State", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia", "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "South Korea", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", "Sweden", "Switzerland", "Syria", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo", "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States of America", "Uruguay", "Uzbekistan", "Vanuatu", "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe"]

def format_time(seconds): return str(timedelta(seconds=int(seconds)))
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()
def apply_styling(df):
    df.index = range(1, len(df) + 1)
    def color_status(val):
        color = '#90EE90' if val == 'Found' else '#FFCCCB'
        return f'background-color: {color}'
    return df.style.applymap(color_status, subset=['Status'])

def reshape_arabic(text):
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        if text and any('\u0600' <= c <= '\u06FF' for c in text):
            return get_display(arabic_reshaper.reshape(text))
        return text
    except: return text

def format_date(date_str):
    if not date_str: return ''
    if 'T' in date_str: date_str = date_str.split('T')[0]
    try: return datetime.strptime(date_str.strip(), '%Y-%m-%d').strftime('%d/%m/%Y')
    except: return date_str

def wrap_text(draw, text, font, max_width):
    lines = []
    words = text.split(' ')
    current_line = ''
    for word in words:
        test_line = current_line + word + ' '
        if draw.textlength(test_line, font=font) <= max_width: current_line = test_line
        else:
            lines.append(current_line.strip())
            current_line = word + ' '
    if current_line: lines.append(current_line.strip())
    return lines

def create_card_image(data, size=(5760, 2700)):
    img = Image.new('RGB', size, color=(250, 250, 250))
    draw = ImageDraw.Draw(img)
    title_font_size = 180
    label_font_size = 140
    value_font_size = 130
    
    def get_font(size, bold=False):
        font_names = ["arialbd.ttf", "DejaVuSans-Bold.ttf", "FreeSansBold.ttf", "arial.ttf", "DejaVuSans.ttf", "FreeSans.ttf"]
        if not bold: font_names = ["arial.ttf", "DejaVuSans.ttf", "FreeSans.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf", "FreeSansBold.ttf"]
        for font_name in font_names:
            try: return ImageFont.truetype(font_name, size)
            except: continue
        return ImageFont.load_default()

    title_font = get_font(title_font_size, bold=True)
    label_font = get_font(label_font_size)
    value_font = get_font(value_font_size)

    draw.rectangle([(0, 0), (size[0], 150)], fill=(218, 165, 32))
    draw.text((120, 40), "H-TRACING", fill=(0, 0, 139), font=title_font)

    photo_x, photo_y = 180, 320
    photo_size = (950, 950)
    draw.rectangle([(photo_x, photo_y), (photo_x + photo_size[0], photo_y + photo_size[1])], outline=(80, 80, 80), width=10, fill=(230, 230, 230))

    if 'Photo' in data and data['Photo']:
        try:
            photo_bytes = base64.b64decode(data['Photo'].split(',')[1])
            personal_photo = Image.open(io.BytesIO(photo_bytes)).resize(photo_size, Image.LANCZOS)
            img.paste(personal_photo, (photo_x, photo_y))
        except:
            draw.text((photo_x + 120, photo_y + photo_size[1] // 2 - 120), "YOUR\nPHOTO\nHERE", fill=(120, 120, 120), font=title_font, align="center")
    else:
        draw.text((photo_x + 120, photo_y + photo_size[1] // 2 - 120), "YOUR\nPHOTO\nHERE", fill=(120, 120, 120), font=title_font, align="center")

    x_label = photo_x + photo_size[0] + 250
    x_value = x_label + 1100
    y_start = 350
    line_height = 190
    fields = [("English Name:", 'English Name'), ("Arabic Name:", 'Arabic Name'), ("Unified Number:", 'Unified Number'), ("EID Number:", 'EID Number'), ("EID Expire Date:", 'EID Expire Date'), ("Visa Issue Place:", 'Visa Issue Place'), ("Profession:", 'Profession'), ("English Sponsor Name:", 'English Sponsor Name'), ("Arabic Sponsor Name:", 'Arabic Sponsor Name'), ("Related Individuals:", 'Related Individuals')]

    y = y_start
    max_value_width = size[0] - x_value - 200
    for label_text, key in fields:
        value = data.get(key, '')
        if key in ['EID Expire Date']: value = format_date(value)
        value_display = reshape_arabic(str(value))
        draw.text((x_label, y), label_text, fill=(0, 0, 0), font=label_font)
        wrapped_lines = wrap_text(draw, value_display, value_font, max_value_width)
        for line in wrapped_lines:
            draw.text((x_value, y), line, fill=(0, 0, 100), font=value_font)
            y += line_height // 1.5
        y += line_height - (len(wrapped_lines) - 1) * (line_height // 1.5)

    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=98)
    buffer.seek(0)
    return buffer

class ICPScraper:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.url = "https://smartservices.icp.gov.ae/echannels/web/client/guest/index.html#/issueQrCode"

    def setup_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option("useAutomationExtension", False)
        
        chrome_bin = "/usr/bin/chromium"
        if not os.path.exists(chrome_bin): chrome_bin = "/usr/bin/chromium-browser"
        if os.path.exists(chrome_bin):
            options.binary_location = chrome_bin
            service = Service("/usr/bin/chromedriver") if os.path.exists("/usr/bin/chromedriver") else Service(ChromeDriverManager().install())
        else:
            service = Service(ChromeDriverManager().install())
        
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"})
        self.wait = WebDriverWait(self.driver, 30)

    def safe_clear_and_fill(self, element, value):
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.BACKSPACE)
        time.sleep(0.5)
        element.send_keys(str(value))

    def select_from_dropdown(self, label_name, search_value):
        try:
            dropdown_xpath = f"//label[contains(text(),'{label_name}')]/following::div[contains(@class,'ui-select-container')][1]"
            container = self.wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
            container.click()
            time.sleep(1)
            input_el = container.find_element(By.XPATH, ".//input[@type='search' or @type='text']")
            input_el.send_keys(search_value)
            time.sleep(1)
            option_xpath = f"//div[contains(@class,'ui-select-choices-row')]//span[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{search_value.lower()}')]"
            self.wait.until(EC.element_to_be_clickable((By.XPATH, option_xpath))).click()
            time.sleep(0.5)
        except Exception as e: logger.error(f"Dropdown error ({label_name}): {e}")

    def perform_single_search(self, passport, nationality, dob, gender):
        result = {'Passport Number': passport, 'Nationality': nationality, 'Status': 'Not Found'}
        try:
            self.setup_driver()
            self.driver.get(self.url)
            self.select_from_dropdown("Nationality", nationality)
            dob_input = self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Date of Birth')]/following::input[1]")))
            self.safe_clear_and_fill(dob_input, dob)
            dob_input.send_keys(Keys.ENTER)
            passport_input = self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Passport Number')]/following::input[1]")))
            self.safe_clear_and_fill(passport_input, passport)
            gender_xpath = f"//label[contains(text(),'Gender')]/following::input[@value='{gender}']"
            self.driver.execute_script("arguments[0].click();", self.wait.until(EC.presence_of_element_located((By.XPATH, gender_xpath))))
            self.driver.execute_script("arguments[0].click();", self.wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'btn-primary') and contains(.,'Search')]"))))
            
            time.sleep(5)
            logs = self.driver.get_log('performance')
            for entry in logs:
                msg = json.loads(entry['message'])['message']
                if msg.get('method') == 'Network.responseReceived':
                    resp = msg['params']['response']
                    if 'issueQrCode' in resp['url'] and resp['status'] == 200:
                        body = self.driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': msg['params']['requestId']})
                        data = json.loads(body['body'])
                        if data.get('success') and data.get('data'):
                            d = data['data']
                            result.update({
                                'Status': 'Found',
                                'English Name': d.get('enName'),
                                'Arabic Name': d.get('arName'),
                                'Unified Number': d.get('unifiedNumber'),
                                'EID Number': d.get('idNumber'),
                                'EID Expire Date': d.get('expiryDate'),
                                'Visa Issue Place': d.get('issuePlaceEn'),
                                'Profession': d.get('professionEn'),
                                'English Sponsor Name': d.get('sponsorEnName'),
                                'Arabic Sponsor Name': d.get('sponsorArName'),
                                'Related Individuals': d.get('relatedIndividualsCount'),
                                'Photo': d.get('photo')
                            })
                            break
        except Exception as e: logger.error(f"Search error: {e}")
        finally:
            if self.driver: self.driver.quit()
        return result

# --- UI ---
tab1, tab2 = st.tabs(["Single Search", "Batch Processing"])
with tab1:
    with st.form("single_form"):
        col1, col2 = st.columns(2)
        p_in = col1.text_input("Passport Number")
        n_in = col2.selectbox("Nationality", countries_list)
        d_in = col1.date_input("Date of Birth", min_value=datetime(1900,1,1))
        g_in = col2.radio("Gender", ["Male", "Female"], horizontal=True)
        submitted = st.form_submit_button("Search")
    if submitted:
        gender_val = "1" if g_in == "Male" else "2"
        with st.spinner("Searching..."):
            res = ICPScraper().perform_single_search(p_in, n_in, d_in.strftime("%d/%m/%Y"), gender_val)
            st.session_state.single_result = res
    if st.session_state.single_result:
        res = st.session_state.single_result
        if res.get('Status') == 'Found':
            st.success("Data Found!")
            card_buf = create_card_image(res)
            st.image(card_buf, use_container_width=True)
            st.download_button("📥 Download Card", card_buf, f"card_{res['EID Number']}.jpg", "image/jpeg")
        else: st.error("No Data Found")

with tab2:
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
    if uploaded_file:
        df_orig = pd.read_excel(uploaded_file)
        st.write(f"Total: {len(df_orig)}")
        col_c1, col_c2, col_c3 = st.columns(3)
        if col_c1.button("▶️ Start"): st.session_state.run_state = 'running'; st.session_state.start_time_ref = time.time()
        if col_c2.button("⏸️ Pause"): st.session_state.run_state = 'paused'
        if col_c3.button("⏹️ Reset"): st.session_state.run_state = 'stopped'; st.session_state.batch_results = []; st.rerun()
        
        prog = st.progress(0); stat = st.empty(); table_area = st.empty()
        for i, row in df_orig.iterrows():
            if st.session_state.run_state == 'stopped': break
            while st.session_state.run_state == 'paused': time.sleep(1)
            if i < len(st.session_state.batch_results): continue
            
            p_num = str(row.get('Passport Number', '')).strip()
            nat = str(row.get('Nationality', 'Egypt')).strip()
            try: dob = pd.to_datetime(row.get('Date of Birth')).strftime('%d/%m/%Y')
            except: dob = str(row.get('Date of Birth', ''))
            gen = str(row.get('Gender', '1')).strip()
            
            stat.info(f"Processing {i+1}/{len(df_orig)}: {p_num}")
            res = ICPScraper().perform_single_search(p_num, nat, dob, gen)
            st.session_state.batch_results.append(res)
            
            df_res = pd.DataFrame(st.session_state.batch_results)
            table_area.table(apply_styling(df_res))
            prog.progress((i + 1) / len(df_orig))
        
        if len(st.session_state.batch_results) == len(df_orig) and len(df_orig) > 0:
            st.success("Finished!")
            st.download_button("📥 Download Excel", to_excel(pd.DataFrame(st.session_state.batch_results)), "results.xlsx")
