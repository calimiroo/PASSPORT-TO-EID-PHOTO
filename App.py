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

# --- Main Application ---
st.title("H-TRACING (ICP)")

st.markdown("""
    <style>
    .stTable td, .stTable th { white-space: nowrap !important; text-align: left !important; padding: 8px 15px !important; }
    .stTable { display: block !important; overflow-x: auto !important; }
    </style>
    """, unsafe_allow_html=True)

# --- Session State Management ---
for key in ['run_state', 'batch_results', 'start_time_ref', 'single_result', 'card_enlarged']:
    if key not in st.session_state:
        st.session_state[key] = 'stopped' if key == 'run_state' else [] if key == 'batch_results' else None if key in ['start_time_ref', 'single_result'] else False

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
    if text and any('\u0600' <= c <= '\u06FF' for c in str(text)):
        reshaped_text = arabic_reshaper.reshape(str(text))
        return get_display(reshaped_text)
    return str(text)

def format_date(date_str):
    if not date_str: return ''
    if 'T' in date_str: date_str = date_str.split('T')[0]
    try:
        return datetime.strptime(date_str.strip(), '%Y-%m-%d').strftime('%d/%m/%Y')
    except:
        return date_str

def wrap_text(draw, text, font, max_width):
    lines = []
    if not text: return lines
    words = text.split(' ')
    current_line = ''
    for word in words:
        test_line = f"{current_line}{word} "
        if draw.textlength(test_line, font=font) <= max_width:
            current_line = test_line
        else:
            lines.append(current_line.strip())
            current_line = f"{word} "
    if current_line:
        lines.append(current_line.strip())
    return lines

def create_card_image(data, size=(5760, 2700)):
    img = Image.new('RGB', size, color=(250, 250, 250))
    draw = ImageDraw.Draw(img)
    title_font_size, label_font_size, value_font_size = 130, 95, 85
    
    try:
        title_font = ImageFont.truetype("arialbd.ttf", title_font_size)
        label_font = ImageFont.truetype("arial.ttf", label_font_size)
        value_font = ImageFont.truetype("arial.ttf", value_font_size)
    except IOError:
        logger.warning("Arial font not found. Using default font. Arabic text might not render correctly.")
        title_font, label_font, value_font = [ImageFont.load_default(size) for size in [title_font_size, label_font_size, value_font_size]]

    draw.rectangle([(0, 0), (size[0], 150)], fill=(218, 165, 32))
    draw.text((120, 40), "H-TRACING", fill=(0, 0, 139), font=title_font)

    photo_x, photo_y, photo_size = 180, 320, (950, 950)
    draw.rectangle([(photo_x, photo_y), (photo_x + photo_size[0], photo_y + photo_size[1])], outline=(80, 80, 80), width=10, fill=(230, 230, 230))

    if data.get('Photo'):
        try:
            photo_bytes = base64.b64decode(data['Photo'].split(',')[1])
            personal_photo = Image.open(io.BytesIO(photo_bytes)).resize(photo_size, Image.LANCZOS)
            img.paste(personal_photo, (photo_x, photo_y))
        except Exception as e:
            logger.warning(f"Failed to load personal photo: {e}")
            draw.text((photo_x + 120, photo_y + photo_size[1] // 2 - 120), "PHOTO\nNOT\nFOUND", fill=(120, 120, 120), font=title_font, align="center")
    else:
        draw.text((photo_x + 120, photo_y + photo_size[1] // 2 - 120), "PHOTO\nNOT\nFOUND", fill=(120, 120, 120), font=title_font, align="center")

    x_label, x_value, y, line_height = photo_x + photo_size[0] + 250, photo_x + photo_size[0] + 1850, 350, 135
    max_value_width = size[0] - x_value - 200
    
    fields = [
        ("English Name:", 'English Name'), ("Arabic Name:", 'Arabic Name'), ("Unified Number:", 'Unified Number'),
        ("EID Number:", 'EID Number'), ("EID Expire Date:", 'EID Expire Date'), ("Visa Issue Place:", 'Visa Issue Place'),
        ("Profession:", 'Profession'), ("English Sponsor Name:", 'English Sponsor Name'), ("Arabic Sponsor Name:", 'Arabic Sponsor Name'),
        ("Related Individuals:", 'Related Individuals')
    ]

    for label_text, key in fields:
        value = data.get(key, '')
        if key == 'EID Expire Date': value = format_date(value)
        
        value_display = reshape_arabic(value)
        draw.text((x_label, y), label_text, fill=(0, 0, 0), font=label_font)
        
        is_arabic = any('\u0600' <= char <= '\u06FF' for char in str(value))
        text_anchor = "ra" if is_arabic else "la"
        
        wrapped_lines = wrap_text(draw, value_display, value_font, max_value_width)
        line_y = y
        for line in wrapped_lines:
            draw.text((x_value + max_value_width if is_arabic else x_value, line_y), line, fill=(0, 0, 100), font=value_font, anchor=text_anchor)
            line_y += line_height // 1.8
        y += line_height * max(1, len(wrapped_lines))

    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=98)
    return buffer.getvalue()

class ICPScraper:
    def __init__(self):
        self.driver, self.wait, self.url = None, None, "https://smartservices.icp.gov.ae/echannels/web/client/guest/index.html#/issueQrCode"

    def setup_driver(self ):
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
        
        if os.path.exists("/usr/bin/chromium-browser"):
            options.binary_location = "/usr/bin/chromium-browser"
            service = Service("/usr/bin/chromedriver")
        else:
            service = Service(ChromeDriverManager().install())
        
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"})
        self.wait = WebDriverWait(self.driver, 30)

    def safe_clear_and_fill(self, element, value):
        element.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
        time.sleep(0.5)
        element.send_keys(str(value))

    def select_from_dropdown(self, label_name, search_value):
        try:
            container = self.wait.until(EC.element_to_be_clickable((By.XPATH, f"//label[contains(text(),'{label_name}')]/following::div[contains(@class,'ui-select-container')][1]")))
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", container)
            container.click()
            time.sleep(1)
            search_input = self.wait.until(EC.visibility_of_element_located((By.XPATH, f"//label[contains(text(),'{label_name}')]/following::input[not(@type='hidden')][1]")))
            self.safe_clear_and_fill(search_input, search_value)
            time.sleep(2)
            result_item = self.wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'ui-select-choices')]//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{str(search_value).lower()}')]")))
            self.driver.execute_script("arguments[0].click();", result_item)
            time.sleep(1)
        except Exception as e:
            logger.warning(f"Dropdown selection failed for {label_name}: {e}")

    def capture_network_data(self):
        logger.info(" [>] Analyzing Network logs...")
        time.sleep(20)
        try:
            for entry in reversed(self.driver.get_log('performance')):
                message = json.loads(entry['message'])['message']
                if 'Network.responseReceived' in message['method']:
                    request_id = message.get('params', {}).get('requestId')
                    try:
                        body = self.driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': request_id})['body']
                        data = json.loads(body)
                        if data.get('isValid'):
                            info = data.get('personalInfo', [{}])[0]
                            return {
                                'English Name': info.get('englishFullName'), 'Arabic Name': info.get('arabicFullName'),
                                'Unified Number': info.get('unifiedNumber'), 'EID Number': info.get('identityNumber'),
                                'EID Expire Date': info.get('identityExpireDate'), 'Visa Issue Place': info.get('englishIdentityIssuePlace'),
                                'Profession': info.get('englishProfession'), 'English Sponsor Name': info.get('englishSponsorName'),
                                'Arabic Sponsor Name': info.get('arabicSponsorName'), 'Status': 'Found'
                            }
                        elif data.get('isValid') is False:
                            return {'Status': 'Not Found'}
                    except: continue
        except Exception as e:
            logger.error(f"Capture Error: {e}")
        return {'Status': 'Not Found'}

    def extract_qr_url(self):
        self.driver.execute_script("if (typeof jsQR === 'undefined') { const script = document.createElement('script'); script.src = 'https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js'; document.head.appendChild(script ); }")
        time.sleep(3)
        return self.driver.execute_async_script("""
            const callback = arguments[arguments.length - 1];
            const el = document.querySelector('canvas') || document.querySelector('img[src*="data:image"]');
            if (!el) return callback(null);
            const canvas = document.createElement('canvas'), context = canvas.getContext('2d'), img = new Image();
            img.crossOrigin = "Anonymous";
            img.src = el.tagName === 'CANVAS' ? el.toDataURL() : el.src;
            img.onload = () => {
                canvas.width = img.width; canvas.height = img.height;
                context.drawImage(img, 0, 0);
                const code = jsQR(context.getImageData(0, 0, img.width, img.height).data, img.width, img.height);
                callback(code ? code.data : null);
            };
            img.onerror = () => callback(null);
        """)

    def perform_single_search(self, passport_number, nationality, date_of_birth, gender):
        self.setup_driver()
        try:
            self.driver.get(self.url)
            logger.info(f"[*] Processing Passport: {passport_number}")
            time.sleep(3)
            self.driver.execute_script("document.querySelector('input[value=\"personalInfo\"]').click();")
            time.sleep(2)
            self.select_from_dropdown('Current Nationality', nationality)
            self.select_from_dropdown('Passport Type', 'ORDINARY PASSPORT')
            self.safe_clear_and_fill(self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Passport Number')]/following::input[1]"))), passport_number)
            dob_field = self.driver.find_element(By.XPATH, "//label[contains(text(),'Date of Birth')]/following::input[1]")
            self.safe_clear_and_fill(dob_field, pd.to_datetime(date_of_birth, dayfirst=True).strftime('%d/%m/%Y'))
            dob_field.send_keys(Keys.TAB)
            gender_field = self.driver.find_element(By.XPATH, "//label[contains(text(),'Gender')]/following::input[1]")
            self.safe_clear_and_fill(gender_field, gender)
            gender_field.send_keys(Keys.TAB)
            
            result, related_count = {'Status': 'Not Found'}, 0
            for rc in range(6):
                logger.info(f"Trying related count: {rc}")
                related_field = self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'related to your file')]/following::input[1]")))
                self.safe_clear_and_fill(related_field, str(rc))
                related_field.send_keys(Keys.TAB)
                time.sleep(1)
                search_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[ng-click='search()']")))
                self.driver.execute_script("arguments[0].click();", search_button)
                time.sleep(5)
                temp_result = self.capture_network_data()
                if temp_result.get('Status') == 'Found':
                    result, related_count = temp_result, rc
                    break
            
            if result.get('Status') == 'Found':
                result.update({'Related Individuals': str(related_count), 'Passport Number': passport_number, 'Nationality': nationality, 'Gender': gender})
                qr_url = self.extract_qr_url()
                if qr_url:
                    logger.info(f"Extracted QR URL: {qr_url}")
                    self.driver.get(qr_url)
                    time.sleep(15)
                    try:
                        photo_elements = self.driver.find_elements(By.CSS_SELECTOR, 'img[src^="data:image"]')
                        if photo_elements:
                            result['Photo'] = max(photo_elements, key=lambda el: len(el.get_attribute('src') or '')).get_attribute('src')
                            logger.info("Personal photo extracted.")
                    except Exception as e:
                        logger.warning(f"Failed to extract personal photo: {e}")
            return result
        except Exception as e:
            logger.error(f"Error during search: {e}")
            return {'Passport Number': passport_number, 'Nationality': nationality, 'Date of Birth': date_of_birth, 'Gender': gender, 'Status': 'Error'}
        finally:
            if self.driver: self.driver.quit()

def toggle_card():
    st.session_state.card_enlarged = not st.session_state.card_enlarged

# --- UI Tabs ---
tab1, tab2 = st.tabs(["Single Search", "Upload Excel File"])

with tab1:
    st.subheader("Single Person Search")
    c1, c2, c3 = st.columns(3)
    p_in = c1.text_input("Passport Number", key="s_p")
    n_in = c2.selectbox("Nationality", countries_list, key="s_n")
    d_in = c3.date_input("Date of Birth", value=None, min_value=datetime(1900,1,1), format="DD/MM/YYYY", key="s_d")
    g_in = st.radio("Gender", options=["Male", "Female"], index=0, key="s_g")
   
    col_btn1, col_btn_stop, col_btn2 = st.columns(3)
    if col_btn1.button("Search Now", key="single_search_button"):
        if p_in and n_in != "Select Nationality" and d_in:
            with st.spinner("Searching..."):
                st.session_state.single_result = ICPScraper().perform_single_search(p_in, n_in, d_in.strftime("%d/%m/%Y"), "1" if g_in == "Male" else "0")
   
    if col_btn_stop.button("🛑 Stop", key="stop_single_search") or col_btn2.button("Clear", key="clear_button"):
        st.session_state.single_result = None
        st.rerun()
   
    if st.session_state.single_result:
        df = pd.DataFrame([st.session_state.single_result])
        st.table(apply_styling(df[['English Name', 'Arabic Name', 'Unified Number', 'EID Number', 'EID Expire Date', 'Visa Issue Place', 'Profession', 'English Sponsor Name', 'Arabic Sponsor Name', 'Related Individuals', 'Status']]))
        if st.session_state.single_result.get('Status') == 'Found':
            card_buffer = create_card_image(st.session_state.single_result)
            st.image(card_buffer, caption="Generated Card", width=1400 if st.session_state.card_enlarged else 700)
            st.button("Enlarge/Shrink Card", on_click=toggle_card)
            st.download_button("📥 Download Card", card_buffer, f"card_{st.session_state.single_result.get('Unified Number', 'unknown')}.jpg", "image/jpeg")

with tab2:
    st.subheader("Batch Processing Control")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
    if uploaded_file:
        df_original = pd.read_excel(uploaded_file)
        st.write(f"Total records: {len(df_original)}")
        st.dataframe(df_original, height=150, use_container_width=True)
        
        col_ctrl1, col_ctrl2, col_ctrl3 = st.columns(3)
        if col_ctrl1.button("▶️ Start / Resume"):
            st.session_state.run_state = 'running'
            if st.session_state.start_time_ref is None: st.session_state.start_time_ref = time.time()
        if col_ctrl2.button("⏸️ Pause"): st.session_state.run_state = 'paused'
        if col_ctrl3.button("⏹️ Stop & Reset"):
            st.session_state.run_state, st.session_state.batch_results, st.session_state.start_time_ref = 'stopped', [], None
            st.rerun()

        if st.session_state.run_state == 'running':
            progress_bar = st.progress(0)
            status_text, stats_area, live_table_area = st.empty(), st.empty(), st.empty()
            
            for i, row in df_original.iterrows():
                if st.session_state.run_state == 'stopped': break
                while st.session_state.run_state == 'paused':
                    status_text.warning("Paused...")
                    time.sleep(1)
                
                if i < len(st.session_state.batch_results): continue

                p_num, nat, dob, gender = str(row.get('Passport Number', '')).strip(), str(row.get('Nationality', 'Egypt')).strip(), str(row.get('Date of Birth', '')), str(row.get('Gender', '1')).strip()
                status_text.info(f"Processing {i+1}/{len(df_original)}: {p_num}")
                
                res = ICPScraper().perform_single_search(p_num, nat, dob, gender)
                st.session_state.batch_results.append(res)
                
                elapsed = time.time() - (st.session_state.start_time_ref or time.time())
                success_count = sum(1 for item in st.session_state.batch_results if item.get("Status") == "Found")
                stats_area.markdown(f"✅ **Success:** {success_count} | ⏱️ **Time:** {format_time(elapsed)}")
                
                df_results = pd.DataFrame(st.session_state.batch_results)
                live_table_area.table(apply_styling(df_results[['English Name', 'Arabic Name', 'Unified Number', 'EID Number', 'EID Expire Date', 'Visa Issue Place', 'Profession', 'English Sponsor Name', 'Arabic Sponsor Name', 'Related Individuals', 'Status']]))
                progress_bar.progress((i + 1) / len(df_original))

            if len(st.session_state.batch_results) == len(df_original):
                st.success("Search Finished!")
                st.download_button("📥 Download Results", to_excel(pd.DataFrame(st.session_state.batch_results)), f"results_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
