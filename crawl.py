from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from excel import Excel
import time,re,os,io,sys
#sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')


class Crawl:
    def is_element_present(self, how, what):
        try:
            self.driver.find_element(by=how, value=what)
        except NoSuchElementException as e: 
            return False
        return True

    def tearDown(self):
        #print('finish')
        self.driver.quit()


    def setUp(self):
        self.path=os.path.dirname(__file__)
        #driver_path=r'C:\Users\user\AppData\Local\Google\Chrome\Application\chromedriver.exe'
        self.driver=webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.implicitly_wait(30)

        base_url = "http://tobaccofreekids.meihua.info/v2/Login2.aspx?ReturnUrl=%2fAdmin%2fnewsdata.aspx"
        driver = self.driver
        driver.get(base_url)
        driver.find_element_by_id("ctl00_cphContent_Login1_UserName").click()
        driver.find_element_by_id("ctl00_cphContent_Login1_UserName").clear()
        driver.find_element_by_id("ctl00_cphContent_Login1_UserName").send_keys("luke_chen")
        driver.find_element_by_id("ctl00_cphContent_Login1_Password").clear()
        driver.find_element_by_id("ctl00_cphContent_Login1_Password").send_keys("V18GWH")
        driver.find_element_by_id("ctl00_cphContent_Login1_LoginButton").click()

    #对list中的title逐一搜索，之后 聚焦首个标签
    def crawl(self,list):
        js="window.open('http://tobaccofreekids.meihua.info/admin/searchNewsv2.aspx')"
        driver=self.driver
        for title in list:
            driver.execute_script(js)
            all_handles = driver.window_handles
            driver.switch_to.window(all_handles[-1])
            #title=r'鼓励纠纷调解机制，及时化解民间纠纷！《长春市文明行为促进条例》实施见成效'
            driver.find_element_by_id("txtSearch").send_keys(title)#添加标题
            sel=driver.find_element_by_tag_name("select")
            Select(sel).select_by_value('3m')#选择转载时间
            driver.find_element_by_id("btnSearch").click()#搜索
            time.sleep(1)
        driver.switch_to.window(all_handles[0])#聚焦首个标签

    def check(self):
        all_handles = self.driver.window_handles   #获取全部页面句柄
        for handle in all_handles[1:]:          #遍历除去第一个的页面句柄
            time1=time.time()
            self.driver.switch_to.window(handle)      #切换到想要的页面
            while not self.is_element_present(By.XPATH,r"//*[contains(text(),'检索完成')]"):
                time2=time.time()
                '''
                if time2-time1>120: 
                    num='Error'
                    return num 
                    '''
                time.sleep(5)


if __name__=='__main__':
    c=Crawl()
    c.setUp()
    e=Excel()
    start_row=2
    end_row=4
    list=e.get_title_list(start_row,end_row)#title_list
    #list=['石家庄发文：推动控烟规定修订，今年打造25个健身与健康融合中心','泰州市推进无烟党政机关建设']
    c.crawl(list)
    num_list=c.insert() 
    e.write_num(num_list)
    #t.check()

    c.tearDown()
    

#driver.current_window_handle #定位当前页面句柄
#current= driver.current_window_handle #定位当前页面句柄

#time.sleep(3)
#driver.quit()


