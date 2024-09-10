import asyncio
import os
import nest_asyncio
import io
import time
import pandas as pd
from telegram import Update, InputFile
from telegram.ext import CallbackContext, Application, CommandHandler, MessageHandler, filters
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# Define the bot token
nest_asyncio.apply()

TOKEN = '7249104078:AAF-ZqmVzWCY0bwXyOtSc8HFpRSQE0VPB78'

# Global state variables
running_tasks = {}
stop_event = asyncio.Event()
async def start(update: Update, context: CallbackContext) -> None:
    user = update.message.from_user
    username = user.username if user.username else user.first_name
    user_id = user.id
    
    # Initialize task tracking for this user
    running_tasks[user_id] = None
    
    # Reply to the user
    await update.message.reply_text(f"Halo! {username} Kirimkan file Excel untuk memulai pengecekan.")
async def stop(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    if user_id in running_tasks and running_tasks[user_id] is not None:
        stop_event.set()
        running_tasks[user_id] = None
        await update.message.reply_text('Proses dihentikan. Mengirimkan hasil yang tersedia...')
        
        # Check if results file exists and send it
        results_file_path = 'results_file.txt'
        if os.path.exists(results_file_path):
            with open(results_file_path, 'rb') as result_file:
                await update.message.reply_document(InputFile(result_file, filename='results_file.txt'))
        else:
            await update.message.reply_text('Tidak ada hasil yang tersedia untuk dikirim.')
    else:
        await update.message.reply_text('Tidak ada proses yang sedang berjalan untuk dihentikan.')
async def process_excel(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    if update.message.document:
        print("process_excel function called")

        # Debugging: Log the type of message received
        print(f"Message Type: {update.message}")

        try:
            # Retrieve the file object
            document = update.message.document
            file_name = document.file_name  # Get the original file name from the document
            file = await document.get_file()  # Ensure to await this call
            file_path = f'user_files/{file_name}'  # Define the path to save the file
            print(f"Downloading file to {file_path}")
            
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            # Download the file to drive
            await file.download_to_drive(file_path)
            print("Download complete")
        except Exception as e:
            await update.message.reply_text(f'Error during file download: {str(e)}')
            print(f"Error during file download: {str(e)}")
            return

        await update.message.reply_text('Memulai proses pemrosesan file Excel...')

        # Set up Selenium
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # Optional: run in headless mode
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        chrome_service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

        def reset_page():
            driver.get("https://myim3.indosatooredoo.com/ceknomor")
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "NIK")))

        reset_page()

        try:
            df = pd.read_excel(file_path, dtype={'NIK': str, 'NO KK': str})
            print("Excel file loaded successfully")
        except Exception as e:
            await update.message.reply_text(f'Gagal memuat file Excel: {str(e)}')
            driver.quit()
            return

        processed_file_path = f'user_files/{file_name}'
        
        results_file_path = f'results_{file_name}.txt'
        with open(results_file_path, 'a') as result_file:
            for index, row in df.iterrows():
                if stop_event.is_set():
                    await update.message.reply_text('Proses dihentikan. Mengirimkan hasil yang tersedia...')
                    driver.quit()
                    if os.path.exists(processed_file_path):
                        with open(processed_file_path, 'rb') as processed_file:
                            await update.message.reply_document(InputFile(processed_file, filename=processed_file_path))
                    else:
                        await update.message.reply_text('Tidak ada hasil yang tersedia untuk dikirim.')
                    return

                try:
                    nik = row['NIK']
                    kk = row['NO KK']
                    
                    nik_field = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, "NIK")))
                    kk_field = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, "KK")))

                    nik_field.clear()
                    nik_field.send_keys(nik)
                    
                    kk_field.clear()
                    kk_field.send_keys(kk)
                    
                    captcha_box = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.CLASS_NAME, "captchaBox")))
                    captcha_box.click()
                    
                    driver.execute_script("document.getElementById('checkSubmitButton').disabled = false;")
                    
                    submit_button = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, "checkSubmitButton")))
                    submit_button.click()
                    
                    time.sleep(2)
                    
                    try:
                        modal = WebDriverWait(driver, 2).until(
                            EC.visibility_of_element_located((By.ID, "myModal"))
                        )
                        if modal.is_displayed():
                            result_text = f"Submission failed for NIK: {nik}, KK: {kk}. The modal appeared."
                            print(result_text)
                            result_file.write(result_text + '\n')
                            df.at[index, 'Result'] = "Invalid data"
                            close_button = driver.find_element(By.XPATH, "//button[@data-dismiss='modal']")
                            close_button.click()
                            reset_page()
                            continue
                    except TimeoutException:
                        pass
                    
                    phone_numbers = []
                    phone_elements = driver.find_elements(By.XPATH, "//ul[@class='list-unstyled margin-5-top']//li")
                    for element in phone_elements:
                        phone_numbers.append(element.text.strip())
                    
                    result_text = ", ".join(phone_numbers) if phone_numbers else "No numbers found"
                    
                    print(f"Result for NIK: {nik}, KK: {kk} -> {result_text}")
                    result_file.write(f"Result for NIK: {nik}, KK: {kk} -> {result_text}\n")
                    
                    df.at[index, 'Result'] = result_text
                    df['NIK'] = df['NIK'].astype(str)
                    df['NO KK'] = df['NO KK'].astype(str)
                    df.to_excel(file_path, index=False)
                    
                    reset_page()
                
                except (NoSuchElementException, TimeoutException, WebDriverException) as e:
                    error_message = f"An error occurred with NIK: {nik}, KK: {kk}. Error: {str(e)}"
                    print(error_message)
                    result_file.write(error_message + '\n')
                    df.at[index, 'Result'] = f"Error: {str(e)}"
                    reset_page()
                    continue

        driver.quit()

        # Send results file to user
        if os.path.exists(processed_file_path):
            with open(processed_file_path, 'rb') as processed_file:
                await update.message.reply_document(InputFile(processed_file, filename=f'{file_name}'))
        else:
            await update.message.reply_text('Tidak ada hasil yang tersedia untuk dikirim.')

    else:
        await update.message.reply_text('Harap kirimkan file Excel dengan format yang benar.')
        print("No document provided")
async def main() -> None:
    # Create Application and dispatcher
    application = Application.builder().token(TOKEN).build()

    # Register handlers for commands
    application.add_handler(CommandHandler('start', start))
    application.add_handler(CommandHandler('stop', stop))
    application.add_handler(MessageHandler(filters.Document.ALL, process_excel))  # Handler for documents

    # Start the bot
    await application.run_polling()

if __name__ == '__main__':
    try:
        asyncio.run(main())
    except RuntimeError as e:
        if str(e) == "This event loop is already running":
            print("Event loop sudah berjalan. Mencoba menggunakan loop yang ada.")
            loop = asyncio.get_event_loop()
            loop.run_until_complete(main())
        else:
            raise
