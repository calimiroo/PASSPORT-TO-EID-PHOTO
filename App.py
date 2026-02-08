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
import arabic_reshaper
from bidi.algorithm import get_display

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- Page Setup ---
st.set_page_config(page_title="ICP Data Search", layout="wide")

# --- Password Protection (Simple Start Page) ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
        <style>
        .big-title {
            font-size: 3.5rem;
            text-align: center;
            margin-top: 100px;
            color: #0d47a1;
        }
        .password-box {
            max-width: 400px;
            margin: 0 auto;
            text-align: center;
            margin-top: 50px;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="big-title">H-TRACING (ICP)</div>', unsafe_allow_html=True)
    st.markdown('<div style="text-align: center; font-size: 1.2rem; color: #555; margin-bottom: 40px;">Enter Password</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown('<div class="password-box">', unsafe_allow_html=True)
        password = st.text_input("Password", type="password", label_visibility="collapsed")
        col1, col2, col3 = st.columns([1, 1, 1])
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

# --- Improve table appearance ---
st.markdown("""
    <style>
    .stTable td, .stTable th {
        white-space: nowrap !important;
        text-align: left !important;
        padding: 8px 15px !important;
    }
    .stTable {
        display: block !important;
        overflow-x: auto !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Session State ---
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

# List of nationalities
countries_list = ["Select Nationality", "Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria", "Burkina Faso", "Burundi", "Cabo Verde", "Cambodia", "Cameroon", "Canada", "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo (Congo-Brazzaville)", "Costa Rica", "Côte d'Ivoire", "Croatia", "Cuba", "Cyprus", "Czechia (Czech Republic)", "Democratic Republic of the Congo", "Denmark", "Djibouti", "Dominica", "Dominican Republic", "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia", "Fiji", "Finland", "France", "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Holy See", "Honduras", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro", "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal", "Netherlands", "New Zealand", "Nicaragua", "Niger", "Nigeria", "North Korea", "North Macedonia", "Norway", "Oman", "Pakistan", "Palau", "Palestine State", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia", "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "South Korea", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", "Sweden", "Switzerland", "Syria", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo", "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States of America", "Uruguay", "Uzbekistan", "Vanuatu", "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe"]

def format_time(seconds):
    return str(timedelta(seconds=int(seconds)))

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
        if text and any('\u0600' <= c <= '\u06FF' for c in text):
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        return text
    except:
        return text

def format_date(date_str):
    if not date_str:
        return ''
    if 'T' in date_str:
        date_str = date_str.split('T')[0]
    try:
        parsed = datetime.strptime(date_str.strip(), '%Y-%m-%d')
        return parsed.strftime('%d/%m/%Y')
    except:
        try:
            parsed = datetime.strptime(date_str.strip(), '%d/%m/%Y')
            return date_str.strip()
        except:
            return date_str
    return date_str

def wrap_text(draw, text, font, max_width):
    lines = []
    words = text.split(' ')
    current_line = ''
    for word in words:
        test_line = current_line + word + ' '
        if draw.textlength(test_line, font=font) <= max_width:
            current_line = test_line
        else:
            lines.append(current_line.strip())
            current_line = word + ' '
    if current_line:
        lines.append(current_line.strip())
    return lines

def create_card_image(data, size=(5760, 2700)):
    img = Image.new('RGB', size, color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    # أحجام خطوط أكبر وأوضح
    title_font_size = 220
    label_font_size = 140
    value_font_size = 130

    try:
        title_font = ImageFont.truetype("arialbd.ttf", title_font_size)
        label_font = ImageFont.truetype("arialbd.ttf", label_font_size)
        value_font = ImageFont.truetype("arial.ttf", value_font_size)
    except:
        title_font = ImageFont.load_default()
        label_font = ImageFont.load_default()
        value_font = ImageFont.load_default()

    # شريط العنوان الذهبي
    draw.rectangle([(0, 0), (size[0], 280)], fill=(218, 165, 32))
    draw.text((200, 60), "H-TRACING", fill=(0, 0, 139), font=title_font)

    # الصورة الشخصية
    photo_x, photo_y = 300, 400
    photo_size = (1400, 1400)
    draw.rectangle([(photo_x, photo_y), (photo_x + photo_size[0], photo_y + photo_size[1])],
                   outline=(50, 50, 50), width=15, fill=(240, 240, 240))

    if 'Photo' in data and data['Photo']:
        try:
            photo_bytes = base64.b64decode(data['Photo'].split(',')[1])
            personal_photo = Image.open(io.BytesIO(photo_bytes))
            personal_photo = personal_photo.resize(photo_size, Image.LANCZOS)
            img.paste(personal_photo, (photo_x, photo_y))
        except Exception as e:
            logger.warning(f"Failed to load photo: {e}")
            draw.text((photo_x + 200, photo_y + 500), "NO PHOTO", fill=(100, 100, 100), font=title_font)
    else:
        draw.text((photo_x + 200, photo_y + 500), "NO PHOTO", fill=(100, 100, 100), font=title_font)

    # مواضع النصوص
    x_label = photo_x + photo_size[0] + 400
    x_value = x_label + 1800
    y_start = 450
    line_height = 220

    fields = [
        ("English Name:", 'English Name'),
        ("Arabic Name:", 'Arabic Name'),
        ("Unified Number:", 'Unified Number'),
        ("EID Number:", 'EID Number'),
        ("EID Expire Date:", 'EID Expire Date'),
        ("Visa Issue Place:", 'Visa Issue Place'),
        ("Profession:", 'Profession'),
        ("English Sponsor Name:", 'English Sponsor Name'),
        ("Arabic Sponsor Name:", 'Arabic Sponsor Name'),
        ("Related Individuals:", 'Related Individuals')
    ]

    y = y_start
    max_value_width = size[0] - x_value - 300

    for label_text, key in fields:
        value = data.get(key, '')
        if key in ['EID Expire Date']:
            value = format_date(value)

        value_display = reshape_arabic(str(value))
        label_display = reshape_arabic(label_text)

        draw.text((x_label, y), label_display, fill=(0, 0, 0), font=label_font)

        wrapped_lines = wrap_text(draw, value_display, value_font, max_value_width)
        for line in wrapped_lines:
            draw.text((x_value, y), line, fill=(0, 0, 100), font=value_font)
            y += line_height
        y += 80  # مسافة إضافية بين الحقول

    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=95)
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
        options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option("useAutomationExtension", False)
        
        import os
        chrome_bin = "/usr/bin/chromium"
        if not os.path.exists(chrome_bin):
            chrome_bin = "/usr/bin/chromium-browser"
            
        if os.path.exists(chrome_bin):
            options.binary_location = chrome_bin
            service = Service("/usr/bin/chromedriver") if os.path.exists("/usr/bin/chromedriver") else Service(ChromeDriverManager().install())
        else:
            service = Service(ChromeDriverManager().install())
        
        self.driver = webdriver.Chrome(service=service, options=options)
        
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        
        self.wait = WebDriverWait(self.driver, 30)

    def safe_clear_and_fill(self, element, value):
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        element.send_keys(value)

    def select_from_dropdown(self, label_text, value):
        try:
            dropdown = self.wait.until(EC.element_to_be_clickable((By.XPATH, f"//label[contains(text(),'{label_text}')]/following::select[1]")))
            dropdown.click()
            option = self.wait.until(EC.element_to_be_clickable((By.XPATH, f"//option[contains(text(),'{value}') or @value='{value}']")))
            option.click()
        except Exception as e:
            logger.warning(f"Dropdown select failed for {label_text}: {e}")

    def capture_network_data(self):
        logs = self.driver.get_log('performance')
        for entry in logs:
            message = json.loads(entry['message'])['message']
            if message['method'] == 'Network.responseReceived':
                url = message['params']['response']['url']
                if 'searchPersonalInfo' in url:
                    request_id = message['params']['requestId']
                    try:
                        response_body = self.driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': request_id})
                        data = json.loads(response_body['body'])
                        if data.get('status') == 'success' and data.get('data'):
                            info = data['data'][0]
                            result = {
                                'Status': 'Found',
                                'English Name': info.get('fullNameEnglish', ''),
                                'Arabic Name': info.get('fullNameArabic', ''),
                                'Unified Number': info.get('unifiedNumber', ''),
                                'EID Number': info.get('emiratesIdNumber', ''),
                                'EID Expire Date': info.get('emiratesIdExpiryDate', ''),
                                'Visa Issue Place': info.get('visaIssuePlace', ''),
                                'Profession': info.get('professionEnglish', ''),
                                'English Sponsor Name': info.get('sponsorNameEnglish', ''),
                                'Arabic Sponsor Name': info.get('sponsorNameArabic', '')
                            }
                            return result
                    except:
                        pass
        return {'Status': 'Not Found'}

    def extract_qr_url(self):
        qr_url = self.driver.execute_async_script("""
            const callback = arguments[arguments.length - 1];
            const extractQR = async () => {
                const link = document.querySelector('a[ng-click="generateQrCode()"]');
                if (link) {
                    link.click();
                    await new Promise(r => setTimeout(r, 3000));
                    return document.querySelector('a[ng-href^="blob:"]')?.getAttribute('ng-href') || 
                           document.querySelector('a[href^="blob:"]')?.href || null;
                }
                return null;
            };
            extractQR().then(callback);
        """)
        return qr_url

    def perform_single_search(self, passport_number, nationality, date_of_birth, gender):
        self.setup_driver()
        try:
            self.driver.get(self.url)
            logger.info(f"[*] Processing Passport: {passport_number}")
            time.sleep(3)
            self.driver.execute_script("""
                var radio = document.querySelector('input[value="personalInfo"]') || document.querySelector('input[ng-value="0"]');
                if(radio) {
                    radio.click();
                    radio.dispatchEvent(new Event('change', { bubbles: true }));
                }
            """)
            time.sleep(2)
            self.select_from_dropdown('Current Nationality', nationality)
            self.select_from_dropdown('Passport Type', 'ORDINARY PASSPORT')
            ppt_field = self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Passport Number')]/following::input[1]")))
            self.safe_clear_and_fill(ppt_field, passport_number)
            dob_formatted = pd.to_datetime(date_of_birth, dayfirst=True).strftime('%d/%m/%Y')
            dob_field = self.driver.find_element(By.XPATH, "//label[contains(text(),'Date of Birth')]/following::input[1]")
            self.safe_clear_and_fill(dob_field, dob_formatted)
            dob_field.send_keys(Keys.TAB)
            gender_field = self.driver.find_element(By.XPATH, "//label[contains(text(),'Gender')]/following::input[1]")
            self.safe_clear_and_fill(gender_field, gender)
            gender_field.send_keys(Keys.TAB)
            related_field = self.driver.find_element(By.XPATH, "//label[contains(text(),'related to your file')]/following::input[1]")
            result = {'Status': 'Not Found'}
            related_count = 0
            logger.info("Trying related count: 0")
            related_field = self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'related to your file')]/following::input[1]")))
            related_field.clear()
            related_field.send_keys("0")
            related_field.send_keys(Keys.TAB)
            time.sleep(1)
            search_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[ng-click='search()']")))
            self.driver.execute_script("arguments[0].removeAttribute('disabled'); arguments[0].classList.remove('disabled'); arguments[0].click();", search_button)
            time.sleep(5)
            temp_result = self.capture_network_data()
            if temp_result.get('Status') == 'Found':
                result = temp_result
                related_count = 0
            else:
                for rc in range(1, 6):
                    logger.info(f"Trying related count: {rc}")
                    related_field.clear()
                    related_field.send_keys(str(rc))
                    related_field.send_keys(Keys.TAB)
                    time.sleep(1)
                    search_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[ng-click='search()']")))
                    self.driver.execute_script("arguments[0].removeAttribute('disabled'); arguments[0].classList.remove('disabled'); arguments[0].click();", search_button)
                    time.sleep(5)
                    temp_result = self.capture_network_data()
                    if temp_result.get('Status') == 'Found':
                        result = temp_result
                        related_count = rc
                        break
            if result.get('Status') == 'Found':
                result['Related Individuals'] = str(related_count)
                if 'EID Expire Date' in result:
                    result['EID Expire Date'] = format_date(result['EID Expire Date'])
                if 'Date of Birth' in result:
                    result['Date of Birth'] = format_date(result['Date of Birth'])
                result['Passport Number'] = passport_number
                result['Nationality'] = nationality
                result['Gender'] = gender
                qr_url = self.extract_qr_url()
                if qr_url:
                    logger.info(f"Extracted QR URL: {qr_url}")
                    self.driver.get(qr_url)
                    time.sleep(15)
                    try:
                        photo_elements = self.driver.find_elements(By.CSS_SELECTOR, 'img[src^="data:image"]')
                        if photo_elements:
                            photo_element = max(photo_elements, key=lambda el: len(el.get_attribute('src') or ''))
                            photo_src = photo_element.get_attribute('src')
                            if photo_src and 'base64' in photo_src:
                                result['Photo'] = photo_src
                                logger.info("Personal photo extracted successfully.")
                    except Exception as e:
                        logger.warning(f"Failed to extract photo: {e}")
            return result
        except Exception as e:
            logger.error(f"Error during search: {e}")
            return {'Passport Number': passport_number, 'Nationality': nationality, 'Date of Birth': date_of_birth, 'Gender': gender, 'Status': 'Error'}
        finally:
            if self.driver:
                self.driver.quit()

def toggle_card():
    st.session_state.card_enlarged = not st.session_state.card_enlarged

tab1, tab2 = st.tabs(["Single Search", "Upload Excel File"])

with tab1:
    st.subheader("Single Person Search")
    c1, c2, c3 = st.columns(3)
    p_in = c1.text_input("Passport Number", key="s_p")
    n_in = c2.selectbox("Nationality", countries_list, key="s_n")
    d_in = c3.date_input("Date of Birth", value=None, min_value=datetime(1900,1,1), format="DD/MM/YYYY", key="s_d")
    g_in = st.radio("Gender", options=["Male", "Female"], index=0, key="s_g")
   
    col_btn1, col_btn_stop, col_btn2 = st.columns(3)
    with col_btn1:
        if st.button("Search Now", key="single_search_button"):
            if p_in and n_in != "Select Nationality" and d_in:
                with st.spinner("Searching..."):
                    scraper = ICPScraper()
                    gender_value = "1" if g_in == "Male" else "0"
                    res = scraper.perform_single_search(p_in, n_in, d_in.strftime("%d/%m/%Y"), gender_value)
                    st.session_state.single_result = res or None
   
    with col_btn_stop:
        if st.button("🛑 Stop", key="stop_single_search"):
            st.session_state.single_result = None
            st.rerun()
   
    with col_btn2:
        if st.button("Clear", key="clear_button"):
            st.session_state.single_result = None
            st.rerun()
   
    single_table_area = st.empty()
    card_image_area = st.empty()
    
    if st.session_state.single_result:
        displayed_fields = ['English Name', 'Arabic Name', 'Unified Number', 'EID Number',
                            'EID Expire Date', 'Visa Issue Place', 'Profession',
                            'English Sponsor Name', 'Arabic Sponsor Name', 'Related Individuals', 'Status']
        filtered_df = pd.DataFrame([{k: v for k, v in st.session_state.single_result.items() if k in displayed_fields}])
        single_table_area.table(apply_styling(filtered_df))
        
        if st.session_state.single_result.get('Status') == 'Found':
            card_buffer = create_card_image(st.session_state.single_result)
            card_image_area.image(card_buffer, caption="Generated Card (Preview)", use_column_width=True)
            
            st.button("🔍 Toggle Full Width" if not st.session_state.card_enlarged else "🔍 Normal View", on_click=toggle_card)
            
            st.download_button(
                label="📥 Download Full Resolution Card",
                data=card_buffer,
                file_name=f"card_{st.session_state.single_result.get('Unified Number', 'unknown')}.jpg",
                mime="image/jpeg"
            )

with tab2:
    st.subheader("Batch Processing Control")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
    if uploaded_file:
        df_original = pd.read_excel(uploaded_file)
        df_show = df_original.copy()
        df_show.index = range(1, len(df_show) + 1)
        st.write(f"Total records: {len(df_original)}")
        st.dataframe(df_show, height=150, use_container_width=True)
        
        col_ctrl1, col_ctrl2, col_ctrl3 = st.columns(3)
        if col_ctrl1.button("▶️ Start / Resume"):
            st.session_state.run_state = 'running'
            if st.session_state.start_time_ref is None:
                st.session_state.start_time_ref = time.time()
        if col_ctrl2.button("⏸️ Pause"):
            st.session_state.run_state = 'paused'
        if col_ctrl3.button("⏹️ Stop & Reset"):
            st.session_state.run_state = 'stopped'
            st.session_state.batch_results = []
            st.session_state.start_time_ref = None
            st.rerun()
            
        progress_bar = st.progress(0)
        status_text = st.empty()
        stats_area = st.empty()
        live_table_area = st.empty()
        actual_success = 0
        
        for i, row in df_original.iterrows():
            while st.session_state.run_state == 'paused':
                status_text.warning("Paused...")
                time.sleep(1)
            if st.session_state.run_state == 'stopped':
                break
            if i < len(st.session_state.batch_results):
                if st.session_state.batch_results[i].get("Status") == "Found":
                    actual_success += 1
                displayed_fields = ['English Name', 'Arabic Name', 'Unified Number', 'EID Number',
                                    'EID Expire Date', 'Visa Issue Place', 'Profession',
                                    'English Sponsor Name', 'Arabic Sponsor Name', 'Related Individuals', 'Status']
                filtered_batch_df = pd.DataFrame([{k: v for k, v in item.items() if k in displayed_fields}
                                                  for item in st.session_state.batch_results])
                live_table_area.table(apply_styling(filtered_batch_df))
                progress_bar.progress((i + 1) / len(df_original))
                continue
                
            p_num = str(row.get('Passport Number', '')).strip()
            nat = str(row.get('Nationality', 'Egypt')).strip()
            try:
                dob = pd.to_datetime(row.get('Date of Birth')).strftime('%d/%m/%Y')
            except:
                dob = str(row.get('Date of Birth', ''))
            gender = str(row.get('Gender', '1')).strip()
            
            status_text.info(f"Processing {i+1}/{len(df_original)}: {p_num}")
            scraper = ICPScraper()
            res = scraper.perform_single_search(p_num, nat, dob, gender)
            if res.get('Status') == 'Found':
                actual_success += 1
            st.session_state.batch_results.append(res)
            
            elapsed = time.time() - (st.session_state.start_time_ref or time.time())
            stats_area.markdown(f"✅ **Success:** {actual_success} | ⏱️ **Time:** {format_time(elapsed)}")
            
            displayed_fields = ['English Name', 'Arabic Name', 'Unified Number', 'EID Number',
                                'EID Expire Date', 'Visa Issue Place', 'Profession',
                                'English Sponsor Name', 'Arabic Sponsor Name', 'Related Individuals', 'Status']
            filtered_batch_df = pd.DataFrame([{k: v for k, v in item.items() if k in displayed_fields}
                                              for item in st.session_state.batch_results])
            live_table_area.table(apply_styling(filtered_batch_df))
            progress_bar.progress((i + 1) / len(df_original))
            
        if len(st.session_state.batch_results) == len(df_original) and len(df_original) > 0:
            st.success("Search Finished!")
            final_df = pd.DataFrame([{k: v for k, v in item.items() if k in displayed_fields}
                                     for item in st.session_state.batch_results])
            excel_data = to_excel(final_df)
            st.download_button(
                label="📥 Download Results",
                data=excel_data,
                file_name=f"search_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
