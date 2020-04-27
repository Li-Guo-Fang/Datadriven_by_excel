# encoding=utf-8
from selenium import webdriver
import unittest, time
import logging, traceback
import ddt
from selenium.common.exceptions import NoSuchElementException

# 初始化日志对象
logging.basicConfig(
    # 日志级别
    level = logging.INFO,
    # 日志格式
    format = '%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
    # 打印日志的时间
    datefmt = '%a, %d %b %Y %H:%M:%S',
    filename = 'd:/report.log',
    filemode = 'a'
)

@ddt.ddt
class TestDemo(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Ie(executable_path = "d:\\IEDriverServer")

    @ddt.data([u"神奇动物在哪里", u"叶茨"],
              [u"疯狂动物城", u"古德温"],
              [u"大话西游之月光宝盒", u"周星驰"])

    @ddt.unpack    #解包
    def test_dataDrivenByObj(self, testdata, expectdata):
        url = "http://www.baidu.com"
        self.driver.get(url)
        self.driver.implicitly_wait(10)
        try:

            self.driver.find_element_by_id("kw").send_keys(testdata)

            self.driver.find_element_by_id("su").click()
            time.sleep(3)

            self.assertTrue(expectdata in self.driver.page_source)
        except NoSuchElementException as e:
            logging.error(u"查找的页面元素不存在，异常堆栈信息：" + str(traceback.format_exc()))
        except AssertionError as e:
            logging.info(u"搜索“%s”，期望“%s”，失败" %(testdata, expectdata))
        except Exception as e:
            logging.error(u"未知错误，错误信息：" + str(traceback.format_exc()))
        else:
            logging.info(u"搜索“%s”，期望“%s”通过" %(testdata, expectdata))

    def tearDown(self):
        self.driver.quit()

if __name__ == '__main__':
    unittest.main()