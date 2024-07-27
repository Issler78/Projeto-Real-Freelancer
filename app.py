import openpyxl
from time import sleep
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common import alert

# Load table
products_table = openpyxl.load_workbook('produtos_ficticios.xlsx')
products_page = products_table['Produtos']

# Enter to the system
options = webdriver.ChromeOptions()
options.add_experimental_option('detach', True)
driver = webdriver.Chrome(options=options)
driver.get('https://cadastro-produtos-devaprender.netlify.app/')
sleep(0.5)

# Read the products table
for row in products_page.iter_rows(values_only=True, min_row=2):
    # Get all values
    name, description, category, product_code, weight, lenght_width_height, price, stock, expiration_date, color, size, material, maker, country, ob, barcode, location_stock = row
    
    
    
    # Get product name input
    product_name_input = driver.find_element(By.XPATH, "//input[@id='product_name']")
    product_name_input.send_keys(name)
    
    # Get description input
    desc_input = driver.find_element(By.XPATH, "//textarea[@id='description']")
    desc_input.send_keys(description)
    
    # Get category input
    category_input = driver.find_element(By.XPATH, "//input[@id='category']")
    category_input.send_keys(category)

    # Get product code input
    product_code_input = driver.find_element(By.XPATH, "//input[@id='product_code']")
    product_code_input.send_keys(product_code)

    # Get weight input
    weight_input = driver.find_element(By.XPATH, "//input[@id='weight']")
    weight_input.send_keys(weight)
    
    # Get L x W x H input
    lenght_width_height_input = driver.find_element(By.XPATH, "//input[@id='dimensions']")
    lenght_width_height_input.send_keys(lenght_width_height)
    
    # Get next step button
    sleep(0.5)
    next_button = driver.find_element(By.XPATH, "//button[@class='btn btn-primary me-2']")
    next_button.click()
    
    
    
    # Get price input
    price_input = driver.find_element(By.XPATH, "//input[@id='price']")
    price_input.send_keys(price)
    
    # Get stock input
    stock_input = driver.find_element(By.XPATH, "//input[@id='stock']")
    stock_input.send_keys(stock)
    
    # Get expiry date input
    expiry_date_input = driver.find_element(By.XPATH, "//input[@id='expiry_date']")
    # Format date
    expiration_date = datetime.strptime(expiration_date, "%Y-%m-%d")
    expiration_date_format = expiration_date.strftime("%d/%m/%Y")
    expiry_date_input.send_keys(expiration_date_format)

    # Get color input
    color_input = driver.find_element(By.XPATH, "//input[@id='color']")
    color_input.send_keys(color)

    # Get size option
    if size == 'Pequeno':
        select_option = driver.find_element(By.XPATH, "//option[@value='Pequeno']")
    elif size == 'Médio':
        select_option = driver.find_element(By.XPATH, "//option[@value='Médio']")
    elif size == 'Grande':
        select_option = driver.find_element(By.XPATH, "//option[@value='Grande']")
    select_option.click()

    # Get material input
    material_input = driver.find_element(By.XPATH, "//input[@id='material']")
    material_input.send_keys(material)
    
    # Get next step button
    sleep(0.5)
    next_button = driver.find_element(By.XPATH, "//button[@class='btn btn-primary me-2']")
    next_button.click()
    
    
    
    # Get manufacturer button
    manufacturer_input = driver.find_element(By.XPATH, "//input[@id='manufacturer']")
    manufacturer_input.send_keys(maker)
    
    # Get country of origin input
    country_input = driver.find_element(By.XPATH, "//input[@id='country']")
    country_input.send_keys(country)
    
    # Get ob input
    ob_input = driver.find_element(By.XPATH, "//textarea[@id='remarks']")
    ob_input.send_keys(ob)
    
    # Get barcode input
    barcode_input = driver.find_element(By.XPATH, "//input[@id='barcode']")
    barcode_input.send_keys(barcode)
    
    # Get location in the stock input
    location_input = driver.find_element(By.XPATH, "//input[@id='warehouse_location']")
    location_input.send_keys(location_stock)
    
    # Get save product button
    sleep(0.5)
    save_button = driver.find_element(By.XPATH, "//button[@class='btn btn-primary me-2']")
    save_button.click()
    sleep(0.5)
    
    
    
    # Click in the alert
    driver.switch_to.alert.accept()

    # Click in the button to add new product
    add_new_product_button = driver.find_element(By.XPATH, "//button[@class='btn btn-primary']")
    add_new_product_button.click()
print('All products have been added successfuly!')
driver.quit()