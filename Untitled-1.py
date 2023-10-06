from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from openpyxl import load_workbook
import time

wb = load_workbook(
    filename="/home/severn/Documents/devel/aspks/RSPO_TEST.xlsx")

sheetRange = wb['Profil']

driver = webdriver.Chrome()
url = ''
driver.get(url)
driver.delete_all_cookies()

cookies = [
    {
        'name': 'csm_gp_session',
        'value': '',
    },
    {
        'name': 'remember_code',
        'value': '',
    },
    # Add more cookies if needed
]

for cookie in cookies:
    driver.add_cookie(cookie)
url = ''
driver.get(url)
driver.maximize_window()

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

    time.sleep(6)
    driver.find_element(By.XPATH, '//*[@id="btnAdd"]').click()

    try:
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/div[5]')))

        driver.find_element(
            By.ID, '_easyui_textbox_input4').send_keys(namapoktan)
        komo = driver.find_element(By.ID, '_easyui_textbox_input5')
        komo.clear()
        komo.send_keys(komoditas)
        driver.find_element(By.ID, '_easyui_textbox_input7').send_keys(alamat)

        prov = driver.find_element(By.ID, '_easyui_textbox_input8')
        prov.clear()
        prov.send_keys(provinsi)
        # provval = driver.find_element(By.NAME, value="ProvID")
        # driver.execute_script(
        #     "arguments[0].setAttribute('value', '61')", provval)
        time.sleep(3)
        kab = driver.find_element(By.ID, '_easyui_textbox_input9')
        kab.clear()
        kab.send_keys(kabupaten)
        # kabval = driver.find_element(By.NAME, value="KabID")
        # driver.execute_script(
        #     "arguments[0].setAttribute('value', '6109')", kabval)
        time.sleep(3)
        kec = driver.find_element(By.ID, '_easyui_textbox_input10')
        kec.clear()
        kec.send_keys(kecamatan)
        # kecval = driver.find_element(By.NAME, value="KecID")
        # driver.execute_script(
        #     "arguments[0].setAttribute('value', '6109040')", kecval)
        time.sleep(3)
        kel = driver.find_element(By.ID, '_easyui_textbox_input11')
        kel.clear()
        kel.send_keys(kelurahan)
        # kelval = driver.find_element(By.NAME, value="KelID")
        # driver.execute_script(
        #     "arguments[0].setAttribute('value', '6109040005')", kelval)

        time.sleep(1)
        driver.find_element(By.ID, '_easyui_textbox_input12').send_keys(ketua)
        driver.find_element(By.ID, '_easyui_textbox_input13').send_keys(notelp)
        driver.find_element(By.ID, '_easyui_textbox_input14').send_keys(email)
        xx = driver.find_element(By.ID, '_easyui_textbox_input15')
        xx.clear()
        xx.send_keys(x)
        ll = driver.find_element(By.ID, '_easyui_textbox_input16')
        ll.clear()
        ll.send_keys(l)
        time.sleep(5)
        driver.find_element(By.ID, 'btnSave').click()
    except TimeoutException:
        print("gagal cuy")
        pass

    time.sleep(1)
    i = i+1
    print("Yey")
