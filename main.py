from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import os
import threading
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from openpyxl import Workbook
import time
import logging
import queue
import shutil
import zipfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SCREENSHOT_FOLDER'] = 'screenshots'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure required directories exist
for folder in [app.config['UPLOAD_FOLDER'], app.config['SCREENSHOT_FOLDER'], 'static']:
    os.makedirs(folder, exist_ok=True)

# Global variables for tracking progress
progress_queue = queue.Queue()
current_status = {"message": "", "progress": 0}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'txt', 'xlsx', 'csv'}

def convert_to_txt(file_path):
    """Convert XLSX or CSV file to TXT format"""
    try:
        # Get file extension
        file_ext = file_path.rsplit('.', 1)[1].lower()
        
        if file_ext == 'txt':
            return file_path
            
        # Generate output txt file path
        txt_path = os.path.splitext(file_path)[0] + '.txt'
        
        # Read file based on extension
        if file_ext == 'xlsx':
            df = pd.read_excel(file_path)
        elif file_ext == 'csv':
            df = pd.read_csv(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_ext}")
            
        # Ensure the DataFrame has the required columns
        required_columns = ['Register Number', 'DOB']
        
        # Try to find columns case-insensitively
        column_mapping = {}
        for req_col in required_columns:
            found = False
            for col in df.columns:
                if col.lower().replace(' ', '') == req_col.lower().replace(' ', ''):
                    column_mapping[req_col] = col
                    found = True
                    break
            if not found:
                raise ValueError(f"Required column not found: {req_col}")
        
        # Select and rename columns
        df = df[[column_mapping['Register Number'], column_mapping['DOB']]]
        df.columns = ['Register Number', 'DOB']
        
        # Format DOB to match required format (assuming it's in a standard format)
        df['DOB'] = pd.to_datetime(df['DOB']).dt.strftime('%d-%m-%Y')
        
        # Save as tab-delimited txt file
        df.to_csv(txt_path, sep='\t', index=False)
        
        # Remove original file
        os.remove(file_path)
        
        return txt_path
        
    except Exception as e:
        if os.path.exists(file_path):
            os.remove(file_path)
        raise Exception(f"Error converting file: {str(e)}")

def create_zip_archive():
    try:
        zip_path = 'static/results_package.zip'
        
        # Delete existing zip if it exists
        if os.path.exists(zip_path):
            os.remove(zip_path)
            
        # Ensure the Excel file exists
        if not os.path.exists('static/results.xlsx'):
            return None
            
        # Create a new zip file
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add Excel file
            zipf.write('static/results.xlsx', arcname='results.xlsx')
            
            # Add screenshots
            screenshot_dir = app.config['SCREENSHOT_FOLDER']
            if os.path.exists(screenshot_dir):
                for filename in os.listdir(screenshot_dir):
                    file_path = os.path.join(screenshot_dir, filename)
                    if os.path.isfile(file_path):
                        zipf.write(file_path, arcname=os.path.join('screenshots', filename))
        
        # Verify the zip file was created and is not empty
        if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
            return zip_path
        return None
        
    except Exception as e:
        print(f"Error creating zip: {e}")
        return None

class ResultScraper:
    def __init__(self, txt_file, website_url):
        self.txt_file = txt_file
        self.website_url = website_url
        self.total_students = 0
        self.processed_students = 0
        self.subjects = []

    def update_progress(self, message, progress=None):
        if progress is not None:
            current_status["progress"] = progress
        current_status["message"] = message
        progress_queue.put(current_status.copy())

    def read_student_data(self):
        students = []
        with open(self.txt_file, 'r') as file:
            next(file)  # Skip header
            for line in file:
                data = line.strip().split('\t')
                if len(data) >= 2:
                    regno, dob = data[:2]
                    students.append((regno.strip(), dob.strip()))
        self.total_students = len(students)
        return students

    def get_page_dimensions(self, driver):
        """Get total height and width of the page"""
        total_height = driver.execute_script("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);")
        total_width = driver.execute_script("return Math.max(document.body.scrollWidth, document.documentElement.scrollWidth);")
        return total_width, total_height

    def take_full_screenshot(self, driver, save_path):
        """Take a full page screenshot"""
        try:
            # Get page dimensions
            width, height = self.get_page_dimensions(driver)
            
            # Add padding to ensure everything is visible
            width += 100
            height += 100
            
            # Set window size to capture everything
            original_size = driver.get_window_size()
            driver.set_window_size(width, height)
            
            # Wait for any resize animations to complete
            time.sleep(1)
            
            # Take screenshot
            element = driver.find_element(By.ID, "printdataResults")
            element.screenshot(save_path)
            
            # Return to original window size
            driver.set_window_size(original_size['width'], original_size['height'])
        except Exception as e:
            print(f"Error taking screenshot: {e}")

    def process_results(self):
        try:
            # Clear previous screenshots
            shutil.rmtree(app.config['SCREENSHOT_FOLDER'], ignore_errors=True)
            os.makedirs(app.config['SCREENSHOT_FOLDER'], exist_ok=True)

            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--start-maximized')
            options.add_argument('--window-size=1920,1080')
            driver = webdriver.Chrome(options=options)

            result_wb = Workbook()
            result_ws = result_wb.active
            
            students = self.read_student_data()
            self.update_progress(f"Found {len(students)} students to process", 0)

            if students:
                success = self.process_single_student(driver, students[0][0], students[0][1], result_ws, True)
                if not success:
                    raise Exception("Failed to process first student and get subject list")

            for idx, (regno, dob) in enumerate(students[1:], 2):
                self.process_single_student(driver, regno, dob, result_ws, False)
                self.update_progress(f"Processing student {idx}/{self.total_students}: {regno}", 
                                   (idx / self.total_students) * 100)

                if idx % 5 == 0 or idx == len(students):
                    result_wb.save('static/results.xlsx')

            driver.quit()
            result_wb.save('static/results.xlsx')
            
            # Create zip archive
            create_zip_archive()
            
            self.update_progress("Processing completed", 100)
            return True

        except Exception as e:
            self.update_progress(f"Error: {str(e)}", 100)
            return False

    def process_single_student(self, driver, regno, dob, worksheet, is_first):
        try:
            driver.get(self.website_url)
            wait = WebDriverWait(driver, 10)
            
            # Fill form
            regno_field = wait.until(EC.element_to_be_clickable((By.ID, "txtRollNo")))
            dob_field = driver.find_element(By.ID, "txtDoB")
            captcha_text = driver.find_element(By.ID, "mainCaptcha")
            captcha_field = driver.find_element(By.ID, "txtInput")

            regno_field.clear()
            dob_field.clear()
            captcha_field.clear()

            regno_field.send_keys(regno)
            dob_field.send_keys(dob)
            captcha_field.send_keys(captcha_text.text)
            
            try:
                submit_button = driver.find_element(By.XPATH, "//button[@type='submit']")
                submit_button.click()
            except:
                captcha_field.send_keys(Keys.RETURN)

            time.sleep(2)

            try:
                result_div = wait.until(EC.presence_of_element_located((By.ID, "printdataResults")))
                
                # Take full page screenshot
                screenshot_path = os.path.join(app.config['SCREENSHOT_FOLDER'], f"{regno}_result.png")
                self.take_full_screenshot(driver, screenshot_path)
                
                # Extract student details
                name = "Unknown"
                sgpa = "N/A"

                # Get all divs inside printdataResults and check for name
                name_elements = driver.find_elements(By.XPATH, "//div[@id='printdataResults']/div")
                for elem in name_elements:
                    if "Name" in elem.text:
                        name = elem.text.split(": ")[1] if ': ' in elem.text else "Unknown"
                        break

                # Get SGPA
                for elem in name_elements:
                    if "SGPA" in elem.text:
                        sgpa = elem.text.split("SGPA")[1].replace(':', '').strip()
                        break

                # Extract subject data
                subjects = []
                grades = []
                table_elements = driver.find_elements(By.XPATH, "//div[@id='printdataResults']//table")
                if table_elements:
                    rows = table_elements[0].find_elements(By.TAG_NAME, "tr")[1:]
                    for row in rows:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        if len(cells) >= 5:
                            subject_code = cells[1].text.strip()
                            subject_name = cells[2].text.strip()
                            grade = cells[4].text.strip()
                            subjects.append(f"{subject_code}-{subject_name}")
                            grades.append(grade)
                
                if is_first:
                    headers = ["Register Number", "Name", "SGPA"] + subjects
                    worksheet.append(headers)
                    self.subjects = subjects
                
                row_data = [regno, name, sgpa] + grades
                worksheet.append(row_data)
                
                return True

            except TimeoutException:
                if is_first:
                    raise Exception("Failed to process first student")
                worksheet.append([regno, "N/A", "N/A"])
                return False

        except Exception as e:
            if is_first:
                raise e
            worksheet.append([regno, "Error", str(e)])
            return False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    website_url = request.form.get('website_url')
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not website_url:
        return jsonify({'error': 'Website URL is required'}), 400
        
    if file and allowed_file(file.filename):
        try:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Convert file to txt if needed
            txt_filepath = convert_to_txt(filepath)
            
            scraper = ResultScraper(txt_filepath, website_url)
            thread = threading.Thread(target=scraper.process_results)
            thread.daemon = True
            thread.start()
            
            return jsonify({'message': 'Processing started'})
            
        except Exception as e:
            return jsonify({'error': str(e)}), 400
    
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/progress')
def get_progress():
    try:
        while True:
            try:
                latest_status = progress_queue.get_nowait()
                return jsonify(latest_status)
            except queue.Empty:
                return jsonify(current_status)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download')
def download_results():
    try:
        # Create the zip file
        zip_path = create_zip_archive()
        
        if zip_path and os.path.exists(zip_path):
            try:
                return send_file(
                    zip_path,
                    mimetype='application/zip',
                    as_attachment=True,
                    download_name='results_package.zip'
                )
            except Exception as e:
                return jsonify({'error': f'Error sending file: {str(e)}'}), 500
        else:
            return jsonify({'error': 'No results found or error creating zip file'}), 404
            
    except Exception as e:
        return jsonify({'error': f'Error processing download: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True)