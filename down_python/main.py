#引入selenium库中的 webdriver 模块
from selenium.webdriver.common.by import By
import time
from selenium import webdriver
print("下载中...")
#获取执行时的路径
from ProjectPath import app_path
root_path=app_path()

#获取信息
with open(root_path+"\\data.txt",'r',encoding='UTF-8') as f:
    data=f.readlines()
    f.close()
uid=data[0][3:len(data[0])-1]
passwd=data[1][3:len(data[1])-1]
down_dir=data[2][5:]

options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': down_dir}
options.add_experimental_option('prefs', prefs)
# options.add_argument('headless')
options.add_argument('--incognito')
driver = webdriver.Chrome(chrome_options=options)
#无头模式下修改下载路径
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': down_dir}}
driver.execute("send_command", params)

driver.get('https://microfun.sharepoint.com/_layouts/15/Authenticate.aspx?Source=/_layouts/15/download.aspx?UniqueId=516b962d-d6d6-4fcd-88f0-16ffacf58647')
'''
考虑到网页打开的速度取决于每个人的电脑和网速，
使用time库sleep()方法，让程序睡眠3秒
'''
time.sleep(3)
#输入账号和密码
driver.find_element(By.XPATH,'//*[@id="i0116"]').send_keys(uid)
driver.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="i0118"]').send_keys(passwd)
driver.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
time.sleep(3)
# driver.get('https://microfun.sharepoint.com/_layouts/15/Authenticate.aspx?Source=/_layouts/15/download.aspx?UniqueId=516b962d-d6d6-4fcd-88f0-16ffacf58647')

driver.quit()

