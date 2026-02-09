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
import sys

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

# --- إذا تم التحقق بنجاح، يستمر التطبيق الرئيسي ---
st.title("H-TRACING (ICP)")

# --- Improve table appearance and make it single line (No Wrap) ---
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

# --- Session State Management ---
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
if 'batch_df' not in st.session_state:
    st.session_state.batch_df = None
if 'current_index' not in st.session_state:
    st.session_state.current_index = 0

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
        import arabic_reshaper
        from bidi.algorithm import get_display
        if text and any('\u0600' <= c <= '\u06FF' for c in text):
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        return text
    except ImportError:
        st.warning("Libraries 'arabic-reshaper' and 'python-bidi' are not installed. Arabic texts may appear unformatted.")
        return text
    except:
        return text

def format_date(date_str):
    if not date_str:
        return ''
    if 'T' in date_str:
        date_str = date_str.split('T')[0]
    try:
        # محاولة تحليل كـ YYYY-MM-DD
        parsed = datetime.strptime(date_str.strip(), '%Y-%m-%d')
        return parsed.strftime('%d/%m/%Y')
    except ValueError:
        try:
            # محاولة تحليل كـ DD/MM/YYYY
            parsed = datetime.strptime(date_str.strip(), '%d/%m/%Y')
            return parsed.strftime('%d/%m/%Y')  # ضمان سنة كاملة
        except ValueError:
            try:
                # إذا كانت سنة فقط، افترض 31/12/YYYY مع سنة كاملة
                if len(date_str.strip()) == 4 and date_str.strip().isdigit():
                    year = int(date_str.strip())
                    return f"31/12/{year:04d}"
                # إذا كانت DD/MM/YY، حول إلى DD/MM/YYYY
                elif '/' in date_str and len(date_str.split('/')[-1]) == 2:
                    parts = date_str.split('/')
                    year = int(parts[-1])
                    full_year = 1900 + year if year > 50 else 2000 + year  # افتراض للقرن
                    return f"{parts[0]}/{parts[1]}/{full_year:04d}"
                else:
                    return date_str
            except:
                return date_str
    return date_str

def wrap_text(draw, text, font, max_width):
    lines = []
    if any('\u0600' <= c <= '\u06FF' for c in text):
        words = text.split(' ')
    else:
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

def create_card_image(data, size=(3500, 2000)):
    img = Image.new('RGB', size, color=(250, 250, 250))
    draw = ImageDraw.Draw(img)
    
    title_font_size = 140
    label_font_size = 100
    value_font_size = 90
    
    fonts_tried = []
    arabic_font_path = "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf"
    default_font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    
    try:
        title_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", title_font_size)
        label_font = ImageFont.truetype(default_font_path, label_font_size)
    except Exception as e:
        logger.warning(f"Failed to load default fonts: {e}")
        title_font = ImageFont.load_default()
        label_font = ImageFont.load_default()
    
    try:
        value_font = ImageFont.truetype(arabic_font_path, value_font_size)
    except Exception as e:
        logger.warning(f"Failed to load Arabic font: {e}. Falling back to default.")
        value_font = ImageFont.truetype(default_font_path, value_font_size)
    
    header_height = 180
    draw.rectangle([(0, 0), (size[0], header_height)], fill=(218, 165, 32))
    
    title_text = "H-TRACING ICP CARD"
    try:
        title_width = draw.textlength(title_text, font=title_font)
    except:
        title_width = len(title_text) * 70
    
    title_x = (size[0] - title_width) // 2
    draw.text((title_x, 50), title_text, fill=(0, 0, 139), font=title_font)
    
    info_x = 100
    info_y = header_height + 50
    
    if data.get('Passport Number'):
        draw.text((info_x, info_y), f"Passport: {data.get('Passport Number', '')}", 
                 fill=(0, 0, 0), font=label_font)
        info_y += 110
    
    if data.get('Nationality'):
        draw.text((info_x, info_y), f"Nationality: {data.get('Nationality', '')}", 
                 fill=(0, 0, 0), font=label_font)
    
    info_y += 110
    if data.get('Date of Birth'):
        dob_formatted = format_date(data.get('Date of Birth', ''))
        draw.text((info_x, info_y), f"Date of Birth: {dob_formatted}", 
                 fill=(0, 0, 0), font=label_font)

    photo_size = (700, 700)
    photo_x = size[0] - photo_size[0] - 100
    photo_y = header_height + 50
    
    draw.rectangle([(photo_x, photo_y), (photo_x + photo_size[0], photo_y + photo_size[1])],
                   outline=(80, 80, 80), width=8, fill=(230, 230, 230))

    if 'Photo' in data and data['Photo']:
        try:
            photo_bytes = base64.b64decode(data['Photo'].split(',')[1])
            personal_photo = Image.open(io.BytesIO(photo_bytes))
            personal_photo = personal_photo.resize(photo_size, Image.LANCZOS)
            img.paste(personal_photo, (photo_x, photo_y))
        except Exception as e:
            logger.warning(f"Failed to load personal photo: {e}")
            draw.text((photo_x + 200, photo_y + photo_size[1] // 2 - 50), 
                     "PHOTO", fill=(120, 120, 120), font=title_font, align="center")
    else:
        draw.text((photo_x + 200, photo_y + photo_size[1] // 2 - 50), 
                 "PHOTO", fill=(120, 120, 120), font=title_font, align="center")

    text_start_x = 100
    text_start_y = header_height + 300
    line_height = 120
    
    fields = [
        ("English Name:", 'English Name'),
        ("Arabic Name:", 'Arabic Name'),
        ("Unified Number:", 'Unified Number'),
        ("EID Number:", 'EID Number'),
        ("EID Expiry Date:", 'EID Expire Date'),
        ("Visa Issue Place:", 'Visa Issue Place'),
        ("Profession:", 'Profession'),
        ("Sponsor Name:", 'English Sponsor Name'),
        ("Arabic Sponsor:", 'Arabic Sponsor Name'),
        ("Related Persons:", 'Related Individuals')
    ]

    y = text_start_y
    for label_text, key in fields:
        value = data.get(key, '')
        if key in ['EID Expire Date', 'Date of Birth']:
            value = format_date(value)
        
        value_display = reshape_arabic(str(value)) if key in ['Arabic Name', 'Arabic Sponsor Name', 'Related Individuals'] else str(value)
        
        draw.text((text_start_x, y), label_text, fill=(0, 0, 0), font=label_font)
        
        value_x = text_start_x + 700
        max_value_width = size[0] - value_x - 150
        
        wrapped_lines = wrap_text(draw, value_display, value_font, max_value_width)
        for i, line in enumerate(wrapped_lines):
            draw.text((value_x, y + (i * 85)), line, fill=(0, 0, 100), font=value_font)
        
        y += line_height + (len(wrapped_lines) - 1) * 85

    footer_y = size[1] - 100
    draw.text((100, footer_y), "Generated by H-TRACING System", 
              fill=(100, 100, 100), font=label_font)
    
    current_date = datetime.now().strftime("%d/%m/%Y %H:%M")
    date_text = f"Date: {current_date}"
    try:
        date_width = draw.textlength(date_text, font=label_font)
    except:
        date_width = len(date_text) * 50
    
    draw.text((size[0] - date_width - 100, footer_y), date_text, 
              fill=(100, 100, 100), font=label_font)

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
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-setuid-sandbox")
        options.add_argument("--no-sandbox")
        
        try:
            options.binary_location = "/usr/bin/chromium"
            service = Service("/usr/bin/chromedriver")
            self.driver = webdriver.Chrome(service=service, options=options)
        except Exception as e:
            logger.warning(f"Failed to use system chromedriver: {e}")
            try:
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service, options=options)
            except Exception as e2:
                logger.error(f"Failed to setup driver: {e2}")
                raise
        
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
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", container)
            container.click()
            time.sleep(1)
            search_input = self.wait.until(EC.visibility_of_element_located((By.XPATH, f"//label[contains(text(),'{label_name}')]/following::input[not(@type='hidden')][1]")))
            self.safe_clear_and_fill(search_input, search_value)
            time.sleep(2)
            result_xpath = f"//div[contains(@class,'ui-select-choices')]//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{str(search_value).lower()}')]"
            result_item = self.wait.until(EC.element_to_be_clickable((By.XPATH, result_xpath)))
            self.driver.execute_script("arguments[0].click();", result_item)
            time.sleep(1)
        except Exception as e:
            logger.warning(f"Dropdown selection failed for {label_name}: {e}")

    def capture_network_data(self):
        logger.info(" [>] Analyzing Network logs...")
        time.sleep(20)
        try:
            logs = self.driver.get_log('performance')
            for entry in reversed(logs):
                message = json.loads(entry['message'])['message']
                if 'Network.responseReceived' in message['method']:
                    params = message.get('params', {})
                    request_id = params.get('requestId')
                    try:
                        resp_obj = self.driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': request_id})
                        body = resp_obj['body']
                        if 'isValid' in body:
                            data = json.loads(body)
                            if 'isValid' in data:
                                if data['isValid']:
                                    personal_info = data.get('personalInfo', {})
                                    info = personal_info[0] if isinstance(personal_info, list) and personal_info else personal_info
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
                                elif data['isValid'] is False:
                                    return {'Status': 'Not Found'}
                    except:
                        continue
        except Exception as e:
            logger.error(f"Capture Error: {e}")
        return {'Status': 'Not Found'}

    def extract_qr_url(self):
        self.driver.execute_script("""
            if (typeof jsQR === 'undefined') {
                const script = document.createElement('script');
                script.src = 'https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js';
                document.head.appendChild(script);
            }
        """)
        time.sleep(3)
        qr_url = self.driver.execute_async_script("""
            const callback = arguments[arguments.length - 1];
            const extractQR = async () => {
                const getQR = () => {
                    let c = document.querySelector('canvas');
                    if (c) return c;
                    let i = document.querySelectorAll('img');
                    for (let img of i) {
                        if (img.src && (img.src.includes('data:image') || img.src.includes('blob') || img.src.includes('qr'))) return img;
                    }
                    return null;
                };
                const el = getQR();
                if (!el) return null;
                const canvas = document.createElement('canvas');
                const context = canvas.getContext('2d');
                const img = new Image();
                img.crossOrigin = "Anonymous";
                img.src = el.toDataURL ? el.toDataURL() : el.src;
                return new Promise((resolve) => {
                    img.onload = () => {
                        canvas.width = img.width;
                        canvas.height = img.height;
                        context.drawImage(img, 0, 0);
                        const imageData = context.getImageData(0, 0, img.width, img.height);
                        const code = jsQR(imageData.data, imageData.width, imageData.height);
                        resolve(code ? code.data : null);
                    };
                    img.onerror = () => resolve(null);
                });
            };
            extractQR().then(callback);
        """)
        return qr_url

    def perform_single_search(self, passport_number, nationality, date_of_birth, gender):
        try:
            self.setup_driver()
            self.driver.get(self.url)
            logger.info(f"[*] Processing Passport: {passport_number}")
            time.sleep(3)
            
            self.driver.execute_script("""
                var radios = document.querySelectorAll('input[type="radio"]');
                for (var i = 0; i < radios.length; i++) {
                    if (radios[i].value === "personalInfo" || radios[i].getAttribute('ng-value') === "0") {
                        radios[i].click();
                        break;
                    }
                }
            """)
            time.sleep(2)
            
            self.select_from_dropdown('Current Nationality', nationality)
            self.select_from_dropdown('Passport Type', 'ORDINARY PASSPORT')
            
            ppt_field = self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'Passport Number')]/following::input[1]")))
            self.safe_clear_and_fill(ppt_field, passport_number)
            
            dob_formatted = format_date(date_of_birth)  # ضمان تنسيق كامل
            dob_field = self.driver.find_element(By.XPATH, "//label[contains(text(),'Date of Birth')]/following::input[1]")
            self.safe_clear_and_fill(dob_field, dob_formatted)
            dob_field.send_keys(Keys.TAB)
            
            gender_field = self.driver.find_element(By.XPATH, "//label[contains(text(),'Gender')]/following::input[1]")
            self.safe_clear_and_fill(gender_field, gender)
            gender_field.send_keys(Keys.TAB)
            
            result = {'Status': 'Not Found'}
            related_count = 0
            
            for rc in range(0, 6):
                logger.info(f"Trying related count: {rc}")
                related_field = self.wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),'related to your file')]/following::input[1]")))
                related_field.clear()
                related_field.send_keys(str(rc))
                related_field.send_keys(Keys.TAB)
                time.sleep(1)
                
                search_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[ng-click='search()']")))
                self.driver.execute_script("arguments[0].click();", search_button)
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
                
                result['Passport Number'] = passport_number
                result['Nationality'] = nationality
                result['Date of Birth'] = dob_formatted
                result['Gender'] = gender
                
                qr_url = self.extract_qr_url()
                if qr_url:
                    logger.info(f"Extracted QR URL: {qr_url}")
                    self.driver.get(qr_url)
                    time.sleep(10)
                    
                    try:
                        images = self.driver.find_elements(By.TAG_NAME, 'img')
                        for img in images:
                            src = img.get_attribute('src')
                            if src and src.startswith('data:image'):
                                result['Photo'] = src
                                logger.info("Personal photo extracted successfully.")
                                break
                    except Exception as e:
                        logger.warning(f"Failed to extract personal photo: {e}")
            
            return result
            
        except Exception as e:
            logger.error(f"Error during search: {e}")
            return {'Passport Number': passport_number, 'Nationality': nationality, 
                    'Date of Birth': date_of_birth, 'Gender': gender, 'Status': 'Error'}
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass

def toggle_card():
    st.session_state.card_enlarged = not st.session_state.card_enlarged

# تبويبات التطبيق
tab1, tab2 = st.tabs(["Single Search", "Batch Search"])

with tab1:
    st.subheader("Single Person Search")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        passport = st.text_input("Passport Number", key="passport")
    with col2:
        nationality = st.selectbox("Nationality", countries_list, key="nationality")
    with col3:
        min_dob = datetime(1900, 1, 1)
        max_dob = datetime.now().date()
        dob = st.date_input("Date of Birth", value=None, format="DD/MM/YYYY", key="dob", min_value=min_dob, max_value=max_dob)
    
    gender = st.radio("Gender", ["Male", "Female"], horizontal=True, key="gender")
    
    col_search, col_clear = st.columns(2)
    with col_search:
        if st.button("🔍 Search Now", type="primary", use_container_width=True):
            if passport and nationality != "Select Nationality" and dob:
                with st.spinner("Searching... This may take a moment."):
                    scraper = ICPScraper()
                    gender_code = "1" if gender == "Male" else "0"
                    result = scraper.perform_single_search(
                        passport, 
                        nationality, 
                        dob.strftime("%d/%m/%Y"), 
                        gender_code
                    )
                    st.session_state.single_result = result
                    st.rerun()
            else:
                st.warning("Please fill all fields")
    
    with col_clear:
        if st.button("Clear Results", use_container_width=True):
            st.session_state.single_result = None
            st.rerun()
    
    if st.session_state.single_result:
        result = st.session_state.single_result
        
        if result.get('Status') == 'Found':
            st.success("✅ Record Found!")
            
            display_data = {
                'Field': ['English Name', 'Arabic Name', 'Unified Number', 'EID Number',
                         'EID Expiry Date', 'Visa Issue Place', 'Profession',
                         'English Sponsor', 'Arabic Sponsor', 'Related Individuals', 'Date of Birth'],
                'Value': [result.get('English Name', ''), result.get('Arabic Name', ''),
                         result.get('Unified Number', ''), result.get('EID Number', ''),
                         result.get('EID Expire Date', ''), result.get('Visa Issue Place', ''),
                         result.get('Profession', ''), result.get('English Sponsor Name', ''),
                         result.get('Arabic Sponsor Name', ''), result.get('Related Individuals', ''), result.get('Date of Birth', '')]
            }
            
            df = pd.DataFrame(display_data)
            st.dataframe(df, use_container_width=True, hide_index=True)
            
            st.subheader("Generated Card")
            
            card_buffer = create_card_image(result)
            
            col_view, col_download = st.columns([1, 2])
            with col_view:
                if st.button("🔍 Enlarge" if not st.session_state.card_enlarged else "🔍 Shrink"):
                    toggle_card()
                    st.rerun()
            
            card_width = 1000 if st.session_state.card_enlarged else 600
            st.image(card_buffer, width=card_width, caption="ICP Digital Card")
            
            with col_download:
                st.download_button(
                    label="📥 Download Card",
                    data=card_buffer,
                    file_name=f"icp_card_{result.get('Unified Number', 'unknown')}.jpg",
                    mime="image/jpeg",
                    use_container_width=True
                )
            
        elif result.get('Status') == 'Not Found':
            st.error("❌ Record Not Found")
            st.info("The provided information does not match any records in the ICP system.")
        else:
            st.error("⚠️ Search Error")
            st.info("An error occurred during the search. Please try again.")

with tab2:
    st.subheader("Batch Processing")
    
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'], 
                                     help="Upload an Excel file with columns: Passport Number, Nationality, Date of Birth, Gender")
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.session_state.batch_df = df
            st.write(f"📊 Found {len(df)} records in the file")
            
            st.dataframe(df.head(), use_container_width=True)
            
            col_start, col_pause, col_stop = st.columns(3)
            
            if st.session_state.run_state == 'stopped':
                pause_btn_disabled = True
            else:
                pause_btn_disabled = False if st.session_state.run_state == 'running' else True
            
            with col_start:
                start_btn = st.button("▶️ Start Processing", use_container_width=True, disabled=(st.session_state.run_state != 'stopped'))
            
            with col_pause:
                pause_btn = st.button("⏸️ Pause / Resume", use_container_width=True, disabled=pause_btn_disabled)
            
            with col_stop:
                stop_btn = st.button("⏹️ Stop & Reset", use_container_width=True, disabled=(st.session_state.run_state == 'stopped'))
            
            if start_btn and st.session_state.run_state == 'stopped':
                st.session_state.run_state = 'running'
                st.session_state.start_time_ref = time.time()
                st.session_state.batch_results = []
                st.session_state.current_index = 0
                st.rerun()
            
            if pause_btn:
                if st.session_state.run_state == 'running':
                    st.session_state.run_state = 'paused'
                elif st.session_state.run_state == 'paused':
                    st.session_state.run_state = 'running'
                st.rerun()
            
            if stop_btn:
                st.session_state.run_state = 'stopped'
                st.session_state.batch_results = []
                st.session_state.current_index = 0
                st.session_state.start_time_ref = None
                st.rerun()
            
            if st.session_state.run_state in ['running', 'paused'] and st.session_state.batch_df is not None:
                progress_bar = st.progress(0)
                status_text = st.empty()
                results_placeholder = st.empty()
                
                if st.session_state.run_state == 'running' and st.session_state.current_index < len(st.session_state.batch_df):
                    row = st.session_state.batch_df.iloc[st.session_state.current_index]
                    passport = row['Passport Number']
                    nationality = row['Nationality']
                    dob = row['Date of Birth']
                    gender = row['Gender']
                    gender_code = "1" if gender.lower() == "male" else "0"
                    
                    scraper = ICPScraper()
                    result = scraper.perform_single_search(passport, nationality, str(dob), gender_code)
                    st.session_state.batch_results.append(result)
                    st.session_state.current_index += 1
                    
                    if st.session_state.current_index < len(st.session_state.batch_df):
                        time.sleep(1)  # تأخير صغير لتجنب الحظر
                        st.rerun()
                    else:
                        st.session_state.run_state = 'stopped'
                        st.rerun()
                
                # عرض التقدم
                progress = st.session_state.current_index / len(st.session_state.batch_df)
                progress_bar.progress(progress)
                
                elapsed_time = time.time() - st.session_state.start_time_ref if st.session_state.start_time_ref else 0
                status_text.text(f"Status: {st.session_state.run_state.capitalize()} | Processed: {st.session_state.current_index}/{len(st.session_state.batch_df)} | Elapsed: {format_time(elapsed_time)}")
                
                if st.session_state.batch_results:
                    results_df = pd.DataFrame(st.session_state.batch_results)
                    styled_df = apply_styling(results_df)
                    results_placeholder.write(styled_df)
                    
                    excel_data = to_excel(results_df)
                    st.download_button(
                        label="📥 Download Results",
                        data=excel_data,
                        file_name="batch_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
        except Exception as e:
            st.error(f"Error reading file: {e}")

# تذييل الصفحة
st.markdown("---")
st.markdown("### 💡 Tips")
st.markdown("""
- Make sure passport number is entered correctly
- Select the correct nationality from the dropdown
- Date format should be DD/MM/YYYY
- For batch processing, ensure your Excel file has the correct column names
""")
