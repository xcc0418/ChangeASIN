import json, os
from selenium import webdriver
from time import sleep
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
# import shutil


def print_pdf(time_now, url):
    print(url)
    # mkdir(url)
    #设置chromedriver
    mkdir(time_now, url)
    chrome_options = webdriver.ChromeOptions()
    #设置超时时间
    sleep(0.5)
    settings = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": ""
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "isHeaderFooterEnabled": False,
        "isLandscapeEnabled": True,#landscape横向，portrait 纵向，若不设置该参数，默认纵向
        "isCssBackgroundEnabled": True,
        "mediaSize": {
            "height_microns": 420000,
            "name": "ISO_A3",
            "width_microns": 297000,
            "custom_display_name": "A3"
        },
    }
    chrome_options.add_argument('--enable-print-browser')
    # chrome_options.add_argument('headless') #headless模式下，浏览器窗口不可见，可提高效率
    chrome_options.add_argument('window-size=1920x1080')
    chrome_options.add_argument('--start-maximized')
    # chrome_options.add_argument('--hide-scrollbars')# 隐藏滚动条, 应对一些特殊页面
    prefs = {
        'printing.print_preview_sticky_settings.appState': json.dumps(settings),
        'savefile.default_directory': f'C:/换标调整/{time_now}-{url}' #此处填写你希望文件保存的路径
    }
    chrome_options.add_argument('--kiosk-printing') #静默打印，无需用户点击打印页面的确定按钮
    chrome_options.add_experimental_option('prefs', prefs)
    s = Service("C:\Program Files\Google\Chrome\Application\chromedriver.exe")
    # browser = webdriver.Chrome(service=s, options=chrome_options)
    browser = webdriver.Chrome(options=chrome_options)
    browser.get("https://erp.lingxing.com/register")
    sleep(1) #等待页面加载
    browser.find_element(By.XPATH, '//*[@class="loginBtn"]').click() #选择账号密码登录
    sleep(1)
    browser.find_element(By.XPATH, "//input[@class='el-input__inner' and @name='account']").send_keys("IT-Test") #输入账户密码
    # sleep(2)
    browser.find_element(By.XPATH, "//input[@class='el-input__inner' and @name='pwd']").send_keys("IT-Test")
    sleep(1)
    browser.find_element(By.XPATH, '//*[@class="el-button loginBtn el-button--primary el-button--large is-round"]').click() #登录
    sleep(1)
    # for i in url:
    sleep(2)
    please_url = f"https://erp.lingxing.com/erp/msupply/adjustmentSheetDetail?order_sn={url}"
    print(please_url)
    browser.get(please_url)
    sleep(5)
    browser.maximize_window()
    # browser.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[1]/div/div/div/div[1]/div/span/button[4]/span").click()
    # sleep(2)
    browser.execute_script(f'document.title="{url}.pdf";window.print();')
    # sleep(1)
    # write_pdf(i[1], i[0])
    # shutil.move(f'D:/加工单/{i[1]}.pdf', rf'D:/加工单/{i[0] + i[1]}/{i[1]}.pdf')
        # os.remove(f'D:/加工单/{i[1]}.pdf')
    #退出/html/body/div[2]/div/div[2]/div[2]/div[1]/div/div/div/div[1]/div/span/button[4]/span
    browser.close()
    browser.quit()


def mkdir(time_now, name):
    folder = os.path.exists(f"C:/换标调整/{time_now}-{name}")
    if not folder:
        os.makedirs(f"C:/换标调整/{time_now}-{name}")
    # else:
    #     os.remove(f"D:/加工单/{name}")
    #     os.makedirs(f"D:/加工单/{name}")


if __name__ == '__main__':
    print_pdf('202206231741', 'AD220624002')
