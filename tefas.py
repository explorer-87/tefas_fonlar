from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import pandas as pd

# İndirilen dosyaların yolunu ayarlıyoruz
download_path = r"C:\Users\adem.aydemir\Desktop\fonlar"
downloaded_file = os.path.join(download_path, "Takasbank TEFAS  Fon Karşılaştırma.xlsx")  # İndirdiğiniz dosyanın adı burada olmalı

# Chrome ayarları
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,  # Dosyaları bu klasöre indir
    "download.prompt_for_download": False,  # İndirme sırasında onay isteme
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# ChromeDriver'ı otomatik olarak indirip kuruyoruz
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# TEFAS sitesine gidiyoruz
driver.get("https://www.tefas.gov.tr/FonKarsilastirma.aspx")

# Sayfanın tamamen yüklenmesini beklemek için WebDriverWait kullanıyoruz
wait = WebDriverWait(driver, 20)

# Xpath ile indirme butonunu bulup tıklıyoruz
xpath_to_download_button = '//*[@id="table_fund_returns_wrapper"]/div[1]/button[3]'
download_button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_to_download_button)))
download_button.click()

# Dosyanın inmesi için biraz bekliyoruz
time.sleep(10)  # İndirme işlemi için 10 saniye bekliyoruz

# İşlem tamamlandıktan sonra tarayıcıyı kapatıyoruz
driver.quit()

# İndirme işleminin tamamlandığını kontrol edin
if os.path.exists(downloaded_file):
    print(f"Fonlar başarıyla {downloaded_file} klasörüne indirildi.")

    # Excel dosyasını okuyoruz
    df = pd.read_excel(downloaded_file)

    # "serbest" kelimesini içeren satırları kaldırıyoruz
    df_filtered = df[~df.iloc[:, 2].str.contains("erbest", case=False, na=False)]  # 3. sütunu filtreliyoruz

    # Güncellenmiş veriyi tekrar Excel dosyasına kaydediyoruz
    df_filtered.to_excel(downloaded_file, index=False)
    print("Serbest Şemsiye Fonu yazan satırlar silindi ve dosya güncellendi.")
else:
    print("İndirme başarısız oldu.")


###############
output_file = downloaded_file  # Çıktı dosyası olarak aynı dosyayı kullanacağız

# Chrome ayarları
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# ChromeDriver'ı otomatik olarak indirip kuruyoruz
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# İndirilen dosyayı okuyoruz (ilk sütunda fon kodları var)
if os.path.exists(downloaded_file):
    df = pd.read_excel(downloaded_file)
    fon_kodlari = df.iloc[:, 0]  # İlk sütunu alıyoruz (fon kodları)

    # Risk değerlerini saklamak için bir liste oluşturuyoruz
    risk_degerleri = []
    alis_valorleri = []
    satis_valorleri = []

    for fon_kodu in fon_kodlari:
        try:
            # TEFAS'taki fon analiz sayfasına gidiyoruz
            driver.get(f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={fon_kodu}")

            # Sayfanın tamamen yüklenmesini bekliyoruz
            wait = WebDriverWait(driver, 20)
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="MainContent_DetailsViewFund"]/tbody/tr[15]/td[2]')))

            # Risk değeri içeren hücreyi XPath ile buluyoruz
            xpath_to_risk_value = '//*[@id="MainContent_DetailsViewFund"]/tbody/tr[15]/td[2]'
            risk_degeri = driver.find_element(By.XPATH, xpath_to_risk_value).text

            # Alış ve satış değerlerini buluyoruz
            xpath_to_alis_value = '//*[@id="MainContent_DetailsViewFund"]/tbody/tr[6]/td[2]'
            alis_degeri = driver.find_element(By.XPATH, xpath_to_alis_value).text

            xpath_to_satis_value = '//*[@id="MainContent_DetailsViewFund"]/tbody/tr[7]/td[2]'
            satis_degeri = driver.find_element(By.XPATH, xpath_to_satis_value).text

            # Risk değerini listeye ekliyoruz
            risk_degerleri.append(risk_degeri)
            alis_valorleri.append(alis_degeri)
            satis_valorleri.append(satis_degeri)

            print(f"Fon Kodu: {fon_kodu}, Risk Değeri: {risk_degeri}, Alış Valörü: {alis_degeri}, Satış Valörü: {satis_degeri}")
        except Exception as e:
            print(f"Fon Kodu: {fon_kodu} için risk değeri alınamadı. Hata: {str(e)}")
            risk_degerleri.append("Alınamadı")
            alis_valorleri.append("Alınamadı")
            satis_valorleri.append("Alınamadı")

    # Risk değerlerini DataFrame'e ekliyoruz
    df['Risk Değeri'] = risk_degerleri  # 11. sütun olarak risk değerini ekliyoruz
    df['Alış Değeri'] = alis_valorleri  # 12. sütun olarak alış değerini ekliyoruz
    df['Satış Değeri'] = satis_valorleri  # 13. sütun olarak satış değerini ekliyoruz

    # Güncellenmiş DataFrame'i aynı dosyaya kaydediyoruz
    df.to_excel(output_file, index=False)
else:
    print("Takasbank TEFAS Fon Karşılaştırma dosyası bulunamadı.")

# Tarayıcıyı kapatıyoruz
driver.quit()

print(f"Risk değerleri {output_file} dosyasına kaydedildi.")