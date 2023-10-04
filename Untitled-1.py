from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from openpyxl import load_workbook
import time

wb = load_workbook(
    filename="D:/python/aspks/RSPO_TEST.xlsx")

sheetRange = wb['Profil']

driver = webdriver.Chrome()
url = 'https://greenpoint.fortasbi.org/desktop/panel/poktan'
driver.get(url)
driver.delete_all_cookies()

cookies = [
    {
        'name': 'csm_gp_session',
        'value': 'scr90ef7s7o05cp72pnihq7t6tp5hbll',
    },
    {
        'name': 'remember_code',
        'value': '9c42b4b4f38c75c39a17fe3cb0acfaceec5d8849.e61e3d2f00a2b44e223f6ca806b85156b5a95f9abf0837efd5da7a0b9b5de46433a0142cd3d3f88eaf69f9854e6cd0f7383de8d5e8935cb721442bc84e0174f2',
    },
    # Add more cookies if needed
]

for cookie in cookies:
    driver.add_cookie(cookie)
url = 'https://greenpoint.fortasbi.org/csm/desktop'
driver.get(url)
driver.maximize_window()
# WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,'/html/body/div[1]/div/div/div/div[2]/div[1]/div[2]/div/table/tbody')))
# Find an element on the webpage and click it
element = driver.find_element(By.XPATH,'/html/body/div[5]/div/div[2]')
element.click()
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnAdd")))
element2 = driver.find_element(By.XPATH,'//*[@id="btnAdd"]')
element2.click()

# looping
i = 2
while i <= len(sheetRange['A']):
    namapoktan = sheetRange['A'+str(i)].value
    komoditas = sheetRange['B'+str(i)].value
    alamat = sheetRange['C'+str(i)].value
    provinsi = sheetRange['D'+str(i)].value
    kabupaten = sheetRange['E'+str(i)].value
    kecamatan = sheetRange['F'+str(i)].value
    kelurahan = sheetRange['G'+str(i)].value
    ketua = sheetRange['H'+str(i)].value
    notelp = sheetRange['I'+str(i)].value
    email = sheetRange['J'+str(i)].value
    x = sheetRange['K'+str(i)].value
    l = sheetRange['L'+str(i)].value

    driver.find_element_by_id('btnAdd').click()

    try:
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/div[5]')))

        driver.find_element_by_id(
            '_easyui_textbox_input4').send_keys(namapoktan)
        driver.find_element_by_id(
            '_easyui_textbox_input5').send_keys(komoditas)
        driver.find_element_by_id('_easyui_textbox_input7').send_keys(alamat)
        driver.find_element_by_id('_easyui_textbox_input8').send_keys(provinsi)
        driver.find_element_by_id(
            '_easyui_textbox_input9').send_keys(kabupaten)
        driver.find_element_by_id(
            '_easyui_textbox_input10').send_keys(kecamatan)
        driver.find_element_by_id(
            '_easyui_textbox_input11').send_keys(kelurahan)
        driver.find_element_by_id('_easyui_textbox_input12').send_keys(ketua)
        driver.find_element_by_id('_easyui_textbox_input13').send_keys(notelp)
        driver.find_element_by_id('_easyui_textbox_input14').send_keys(email)
        driver.find_element_by_id('_easyui_textbox_input15').send_keys(x)
        driver.find_element_by_id('_easyui_textbox_input16').send_keys(l)
        driver.find_element_by_id('btnSave').click()
    except TimeoutException:
        print("gagal cuy")
        pass

    time.sleep(1)
    i = i+1
    print("Yey")
