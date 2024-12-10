
# 通过下面的代码打开一个edge浏览器示例
# "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\EdgeDebug"

import subprocess
import socket
from selenium import webdriver
from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import Select
# from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import setting as S
import random


class TestRecordAutomator():
    '''
    main body of project, start edge and check window handles
    '''
    def __init__(self) -> None:

        self.start_edge()
        
        options = webdriver.EdgeOptions()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")  # 连接调试端口
        self.driver = webdriver.Edge(options=options)
        self.handles = self.driver.window_handles

        self.actions = ActionChains(self.driver)
        

        for handle in self.handles:
            self.driver.switch_to.window(handle)
            try:
                type = self.driver.find_element(By.ID, "aws-form-title")
                if type.text =='氧化性固体测定原始记录':
                    o1 = O1(self.driver,self.actions)
                    o1.create_o1_ref_record()
                    o1.create_o1_test_record()

                elif type.text == '液体持续燃烧性测定原始记录':
                    cb = CB(self.driver,self.actions)
                    cb.create_cb_record()

                elif type.text == '遇水放出易燃气体物质测定原始记录':
                    efg = EFG(self.driver,self.actions)
                    efg.create_EFG_record()

            except Exception as e:
                print(e)
                pass

        self.driver.quit()
        pass
    def is_port_in_use(self):
        '''
        检查9222端口是否已经被使用
        '''
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            return s.connect_ex(("127.0.0.1", 9222)) == 0
    def start_edge(self):
        print('checking edge status')
        edge_command = r'"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\EdgeDebug"'
        if not self.is_port_in_use():
            print('starting edge')
            subprocess.Popen(edge_command, shell=True)

        else:
            print('edge running already')
            pass

class EFG():
    def __init__(self,driver,actions) -> None:
        self.driver=driver
        self.actions = actions
        self.basic_info = Basic_info(self.driver)
        pass
    def create_EFG_record(self):
        self.driver.execute_script("window.scrollTo(0, 0);")
        self.basic_info.select_room('4-214')
        self.basic_info.select_box('4818')
        self.basic_info.select_method('ST/SG/AC.10/11/Rev.7 33.5.4试验N.5')

        self.driver.execute_script(f"""
    var selectElement = document.getElementById('YPZBGC');
    $(selectElement).val('未经处理。').trigger('change');
""")

        input_with_slash_list = [
            'CQYPZL',
            "SYSXZL",
            'SXYPZLB',
            'YPSLRS',
            'SYXXRS',
            'BXJLRS',
            'YPSLTL',
            'SYXXTL',
            'BXJLTL',
            'YPSLDS',
            'SYXXDS',
            'BXJLDS',
            'YPL2',
            'YPL3',
            'CDS2',
            'CDS3',
            'ZDS2',
            'ZDS3',
            'JSTJ2',
            'JSTJ3'
        ]
        for i in input_with_slash_list:
            self.fill_with_slash(i)
        input = self.driver.find_element(By.ID, 'YPL1')
        input.clear()
        input.send_keys(S.EFG['mass'])

        input = self.driver.find_element(By.ID, 'CDS1')
        input.clear()
        input.send_keys('100')
        input = self.driver.find_element(By.ID, 'ZDS1')
        input.clear()
        input.send_keys('60')

        self.add_efg_record(12)
        time.sleep(0.5)
        rows = [1,2,3,4,5,6,7,8,9,10,11,12]
        cols = [1,2,3,4,5,6,7]
        for row in rows:
            for col in cols:
                if col == 1:
                    if row <= 5:
                        content = str(row)+'min'
                    else:
                        content = str(row-5)+'h'
                elif col == 2:
                    content = '51.0'
                elif col == 5:
                    if row <= 5:
                        content = '0.0mL/g·min'
                    else:
                        content = '0.0mL/g·h'
                    
                # elif col == 6 or col == 7:
                #     content = '/'
                else:
                    content = '/'
                self.input_efg_record(row,col,content)
        
        input = self.driver.find_element(By.ID, 'KBCD1')
        input.clear()
        input.send_keys('51.0')
        input = self.driver.find_element(By.ID, 'KBCD2')
        input.clear()
        input.send_keys('51.0')


    def fill_with_slash(self,ID):
        input = self.driver.find_element(By.ID, ID)
        input.clear()
        input.send_keys('/')

    def add_efg_record(self,count):
        buttons = self.driver.find_elements(By.CSS_SELECTOR, ".button.blue")
        
        for order,button in enumerate(buttons):
            print(order,button.text)
            if button.text == '新增' and order == 3:
                for _ in range(count):
                    button.click()

    def input_efg_record(self,row,col,content):
        print(f'content is {content}')
        content = str(content)
        # //*[@id="BO_QWS_GHS_TEST_HGP_YSFQ_SYJL"]/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[3]
        cell = self.driver.find_element(By.XPATH, f'//*[@id="BO_QWS_GHS_TEST_HGP_YSFQ_SYJL"]/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[@pq-row-indx="{row-1}"]/td[@pq-col-indx="{col}"]')
        self.driver.execute_script("arguments[0].click();", cell)
        for c in str(content):
            self.actions.send_keys(c).perform()
        self.actions.send_keys(Keys.ENTER).perform()
        self.actions.move_by_offset(1, 1).click().perform()


class CB():
    '''
    处理持续燃烧实验
    '''
    def __init__(self,driver,actions) -> None:
        self.driver=driver
        self.actions = actions
        self.basic_info = Basic_info(self.driver)
        pass
    def create_cb_record(self):
        '''
        生成默认原始记录
        '''
        self.basic_info.select_room('4-214')
        self.basic_info.select_box(S.CB_DATA['Box'])
        # app_list = ['DC135','OC226','PC055']
        # self.select_app(app_list)
        # select method
        self.basic_info.select_method('ST/SG/AC.10/11/Rev.7 32.5.2试验L.2')

        # input P,T,delta P
        input = self.driver.find_element(By.ID, "QYBDS")
        print("QYBDS found")
        input.clear()
        input.send_keys(S.CB_DATA['P'])
        input = self.driver.find_element(By.ID, "QYBWD")
        input.clear()
        input.send_keys(S.CB_DATA['T'])
        input = self.driver.find_element(By.ID, "SZXZZ")
        input.clear()
        input.send_keys(S.CB_DATA['DeltaT'])

        # add 8 records
        self.add_cb_record(8)
        time.sleep(0.5)
        
        # fill in records
        rows = [1,2,3,4,5,6,7,8]
        cols = [1,2,4,5,6,7,8]

        for row in rows:
            for col in cols:
                if col == 1:
                    content = '关' if row%2 else '开'
                elif col == 2:
                    content = '当T0=60.5' if row<=4 else '当T0=75'
                elif col == 4:
                    if row<=2 or row==5 or row ==6:
                        content = '60'
                    else:
                        content = '30'
                elif col == 5 or col == 6 or col == 7:
                    content = '不可持续燃烧'
                else:
                    content = '/'
                print(f'start filling in row{row},col{col}')
                self.input_cb_record(row,col,content)



        t1 = self.get_cb_cell_info(1,3)
        t2 = self.get_cb_cell_info(5,3)

        input = self.driver.find_element(By.ID, "SYJG")
        input.clear()
        input.send_keys(f'样品在{t1}℃及{t2}℃不持续燃烧')
        input = self.driver.find_element(By.ID, "SYBZ")
        input.clear()
        input.send_keys('/')
    def add_cb_record(self,count):
        buttons = self.driver.find_elements(By.CSS_SELECTOR, ".button.blue")
        
        for order,button in enumerate(buttons):
            print(order,button.text)
            if button.text == '新增' and order == 3:
                for _ in range(count):
                    button.click()

    def get_cb_cell_info(self,row,col):
        cell = self.driver.find_element(By.XPATH, f'//*[@id="BO_QWS_GHS_TEST_HGP_CXRS_SY"]/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[@pq-row-indx="{row-1}"]/td[@pq-col-indx="{col}"]')
        cell_text = self.driver.execute_script('return arguments[0].innerText', cell)
        print(cell_text)
        return cell_text
    
    def input_cb_record(self,row,col,content):
        print(f'content is {content}')
        content = str(content)
        # //*[@id="BO_QWS_GHS_TEST_HGP_CXRS_SY"]/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[3]
        cell = self.driver.find_element(By.XPATH, f'//*[@id="BO_QWS_GHS_TEST_HGP_CXRS_SY"]/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[@pq-row-indx="{row-1}"]/td[@pq-col-indx="{col}"]')
        self.driver.execute_script("arguments[0].click();", cell)


        for c in str(content):
            self.actions.send_keys(c).perform()
        self.actions.send_keys(Keys.ENTER).perform()
        self.actions.move_by_offset(1, 1).click().perform()


class O1():
    '''
    处理O.1原始记录
    '''
    def __init__(self,driver,actions) -> None:
        self.driver = driver
        self.actions = actions
        self.basic_info = Basic_info(self.driver)
        pass
    
    def create_o1_ref_record(self):
        self.basic_info.select_room('4-214')
        self.basic_info.select_box('4818')
        # <option value="ST/SG/AC.10/11/Rev.8 34.4.1 试验O.1" data-select2-id="3">联合国《试验和标准手册》 第八修订版 34.4.1 试验O.1 氧化性固体的试验(ST/SG/AC.10/11/Rev.8 34.4.1 试验O.1)</option>
        self.basic_info.select_method('ST/SG/AC.10/11/Rev.7 34.4.1试验O.1')
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        self.add_o1_r_record(13)
        time.sleep(1)

        rows = [1,2,3,4,5]
        cols = [1,2,3,4,5]
        data = S.O1_III_R['data']
        for row in rows:
            for col in cols:
                if col == 1:
                    key = 'class'
                elif col == 2:
                    key = 'count'
                elif col == 3:
                    key = 'mass_r'
                elif col == 4:
                    key = 'mass_c'
                elif col == 5:
                    key = 'result'
                self.input_o1_r_cell(row,col,data[row-1][key])

        row_ave = 6
        for i in range(6):
            if i <5:
                self.input_o1_r_cell(row_ave,i,'/')
            else:
                self.input_o1_r_cell(row_ave,i,S.O1_III_R['average'])
        
        rows = [7,8,9,10,11]
        data = S.O1_II_R['data']
        for row in rows:
            for col in cols:
                if col == 1:
                    key = 'class'
                elif col == 2:
                    key = 'count'
                elif col == 3:
                    key = 'mass_r'
                elif col == 4:
                    key = 'mass_c'
                elif col == 5:
                    key = 'result'
                # print(f'filling cell{row}{col}with{data[row-7][key]}')
                self.input_o1_r_cell(row,col,data[row-7][key])
        
        row_ave = 12
        for i in range(6):
            if i <5:
                self.input_o1_r_cell(row_ave,i,'/')
            else:
                self.input_o1_r_cell(row_ave,i,S.O1_II_R['average'])

        row_13 = 13
        for i in range(6):
            if i == 1:
                self.input_o1_r_cell(row_13,i,'I类平均')
            elif i == 5 :
                self.input_o1_r_cell(row_13,i,'6')
            else:
                self.input_o1_r_cell(row_13,i,'/')

        pass

    def add_o1_r_record(self,count):
        # print(f'start o1 record in :{self.driver.current_window_handle}')
        buttons = self.driver.find_elements(By.CSS_SELECTOR, ".button.blue")
        for order,button in enumerate(buttons):
            if button.text == '新增' and order == 3:
                for _ in range(count):
                    button.click()

    def get_o1_r_cell_info(self,row,col):
        cell = self.driver.find_element(By.XPATH, f'//*[@id="BO_QWS_GHS_TEST_HGP_YHXGT_BZ"]/div[2]/div[2]/div[1]/div[1]/table/tbody/tr[@pq-row-indx="{row-1}"]/td[@pq-col-indx="{col}"]')
        cell_text = self.driver.execute_script('return arguments[0].innerText', cell)
        print(cell_text)

    def input_o1_r_cell(self,row,col,content):
        cell = self.driver.find_element(By.XPATH, f'//*[@id="BO_QWS_GHS_TEST_HGP_YHXGT_BZ"]/div[2]/div[2]/div[1]/div[1]/table/tbody/tr[@pq-row-indx="{row-1}"]/td[@pq-col-indx="{col}"]')
        self.driver.execute_script("arguments[0].click();", cell) 
        for c in str(content):
            self.actions.send_keys(c).perform()
        self.actions.send_keys(Keys.ENTER).perform()
        self.actions.move_by_offset(1, 1).click().perform()
        pass

    def add_o1_t_record(self,count):
        buttons = self.driver.find_elements(By.CSS_SELECTOR, ".button.blue")
        for order,button in enumerate(buttons):
            # print(order,button.text)
            if button.text == '新增' and order == 5:
                for _ in range(count):
                    button.click()

    def input_o1_t_cell(self,row,col,content):
        cell = self.driver.find_element(By.XPATH, f'//*[@id="BO_QWS_GHS_TEST_HGP_YHXGT_DC"]/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[@pq-row-indx="{row-1}"]/td[@pq-col-indx="{col}"]')
        self.driver.execute_script("arguments[0].click();", cell) 
        # //*[@id="BO_QWS_GHS_TEST_HGP_YHXGT_DC"]/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[3]
        
        for c in str(content):
            self.actions.send_keys(c).perform()
        self.actions.send_keys(Keys.ENTER).perform()
        self.actions.move_by_offset(1, 1).click().perform()
        pass

    def create_o1_test_record(self):
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        self.add_o1_t_record(2)
        time.sleep(1)

        row = 1
        cols = [1,2,3,4,5]
        mass_t = round(random.uniform(23.90, 24.10), 2)
        mass_c = round(random.uniform(5.90, 6.10), 2)
        content = ['4:1',1,mass_t,mass_c,'>180']
        for col in cols:
            self.input_o1_t_cell(row,col,content[col-1])

        row = 2
        mass_t = round(random.uniform(14.90, 15.10), 2)
        mass_c = round(random.uniform(14.90, 15.10), 2)

        content = ['1:1',1,mass_t,mass_c,'>180']
        for col in cols:
            self.input_o1_t_cell(row,col,content[col-1])

        input = self.driver.find_element(By.ID, "SYJG")
        input.clear()
        input.send_keys('样品与纤维素丝混合比例为4:1时的燃烧时间和比例为1:1时的燃烧时间s均大于III类标准物质平均燃烧时间109S。')
        input = self.driver.find_element(By.ID, "SYBZ")
        input.clear()
        input.send_keys('/')

class Basic_info():
    '''
    处理通用的基础信息
    '''
    def __init__(self,driver) -> None:
        self.driver=driver
        pass
    def select_room(self,room):
        # print('starting 场所选择')
        room_input = self.driver.find_element(By.ID, "CSXZ")
        room_input.click()
        self.driver.switch_to.frame("id-awsui-win-frm-2013-frmCSXZ")
        room_div = self.driver.find_element(By.XPATH, f"//div[text()='{room}']")
        room_div.click()
        self.driver.switch_to.default_content()
        pass

    def select_box(self,box):
        # CCX-HG21-3670
        # print('starting 箱号选择')
        room_input = self.driver.find_element(By.ID, "ZZXBH")
        room_input.click()
        self.driver.switch_to.frame("id-awsui-win-frm-2013-frmZZXBH")
        
        time.sleep(0.5)
        input = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='模糊检索 : 周转箱编号']")
        input.send_keys(f'{box}')
        time.sleep(0.5)
        input.send_keys(Keys.ENTER)
        
        time.sleep(0.5)
        box_div = self.driver.find_element(By.XPATH, f"//div[text()='化工品运输鉴定已检样品储存箱']")
        if box_div:
            box_div.click()
            print('已找到箱子')
        else:
            print('未找到箱子')
        self.driver.switch_to.default_content()

        pass
    def select_method(self,method):
        self.driver.execute_script(f"""
    var selectElement = document.getElementById('YJFF');
    $(selectElement).val('{method}').trigger('change');
""")

        pass
    def select_app(self,No_list):
        '''
        under developing
        '''
        button = self.driver.find_element(By.ID,'FormGridTB_b81196cc-286f-4248-827d-fd87e5822fce_1_Btn_AddRef')
        button.click()
        self.driver.switch_to.frame('id-awsui-win-frm-2013-frm')

        for No in No_list:
            app_div = self.driver.find_element(By.XPATH, f"//div[text()='{No}']")
            app_div.click()
        time.sleep(0.5)

        # 获取所有按钮
        buttons = self.driver.find_elements(By.CSS_SELECTOR, ".button.blue")

        # 打印每个按钮的文本
        for i, button in enumerate(buttons):
            print(f"Button {i}: {button.text.strip()}")

        # 点击第二个按钮
        if len(buttons) > 1:
            self.driver.execute_script("arguments[0].click();", buttons[1]) 
        else:
            print("未找到第二个按钮")

        self.driver.switch_to.default_content()
        pass


if __name__ == "__main__":
    work = TestRecordAutomator()

