import os
import psutil
import requests, threading
from lxml import etree
from queue import Queue
from win32com.client import Dispatch


class OkZiyuan():
    def __init__(self, search_content=None):
        self.start_url = 'http://www.jisudhw.com/index.php?m=vod-search'
        self.base_url = 'http://www.jisudhw.com'
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.106 Safari/537.36 Edg/83.0.478.54'}
        self.data = {'wd': search_content,
                     'submit': 'search'}
        self.url_queue = Queue()
        self.url_lists = Queue()
        self.url_content = Queue()
        self.vido = Queue()
        self.vido_download_name = Queue()

    # 搜索
    def search(self):
        response = requests.post(self.start_url, headers=self.headers, data=self.data)
        response = response.content.decode('utf-8')
        # print(response)
        self.url_queue.put(response)  # 自动补充响应后返回到队列url_queue

    # 得到搜索内容,并提取url推送到url_lists队列
    def get_search_content_list(self):
        while True:  # 多页
            ret = self.url_queue.get()
            ret = etree.HTML(ret)
            ret = ret.xpath('//li//span[@class="xing_vb4"]/a')
            for i, index in enumerate(ret):
                item = {}
                item['url'] = self.base_url + index.xpath('./@href')[0] if len(
                    index.xpath('./@href')) > 0 else None  # 获取标题选项
                item['text'] = index.xpath('./text()')[0] if len(index.xpath('./text()')) > 0 else None  # 获取标题
                print(i, item)
                self.url_content.put(item['url'])
            self.url_queue.task_done()  # 完成一个队列减少1

    # 以get请求访问url，返回etree.HTML
    def visit_url(self, url):
        response = requests.get(url, headers=self.headers)
        response = response.content.decode()
        # print(response)
        response = etree.HTML(response)
        return response

    # 从url_lists中的url提取数据，并逐个访问,获取视频列表
    def get_content(self):
        while True:
            url_content = self.url_content.get()  # 获取搜索的url
            # print(url_content)
            response = self.visit_url(url_content)
            ret = response.xpath(r'//div[@id="down_1"]/ul/li/input')
            for index in ret:
                url = str(index.xpath('./@value') if len(index.xpath('./@value')) > 0 else None).replace('[',
                                                                                                         '').replace(
                    ']', '')
                self.vido.put(url)
                vido_name = os.path.split(url)
                self.vido_download_name.put(vido_name[1])
            self.url_content.task_done()

    def download(self):
        import pythoncom
        pythoncom.CoInitialize()  # 读取word文档的内容，常见错误是，读英文的时候，没有问题，但是碰到中文的时候，就会报错，见下面代码：
        if isinstance(self.proc_exist('Thunder.exe'), int) != None:  # isinstance() 函数来判断一个对象是否是一个已知的类型，类似 type()。
            while True:
                down_url = self.vido.get()  # 获取视频地址
                # output_filename = self.vido_download_name.get()  # 获取视频名字
                print(down_url.replace("'", ""))
                thunder = Dispatch('ThunderAgent.Agent64.1')
                thunder.AddTask(down_url.replace("'", ""))
                thunder.CommitTasks()
                self.vido.task_done()
                # self.vido_download_name.task_done()
        else:
            print('no such process...')
            os.system(r'E:\SoftWare\Program\Thunder.exe')
            self.download()

    # 判断是否存在进程
    def proc_exist(self, process_name):
        pl = psutil.pids()  # 检查所有进程
        for pid in pl:
            if psutil.Process(pid).name() == process_name:
                return pid
            else:
                return None

    # 主程序
    def run(self):
        thread_lists = []  # 一个子线程列表
        self.search()  # 搜索
        # 得到搜索内容,并提取url推送到url_lists队列
        t_url1 = threading.Thread(target=self.get_search_content_list)
        thread_lists.append(t_url1)

        # 从url_lists中的url提取数据，并逐个访问,获取视频列表
        for i in range(2):
            t_url2 = threading.Thread(target=self.get_content)
            thread_lists.append(t_url2)

        # 存储数据
        for i in range(2):
            t_url3 = threading.Thread(target=self.download)
            thread_lists.append(t_url3)

        # 执行子线程
        for thread_list in thread_lists:
            thread_list.setDaemon(True)  # 把子线程设置为守护线程，该线程不重要，主线程结束，子线程结束
            thread_list.start()

        for q in [self.url_queue, self.url_lists, self.vido, self.url_content]:  # ,self.vido_download_name
            q.join()  # 让主线程等待阻塞，等待队列的任务完成之后再完成

        print('完成！')
        
if __name__ == '__main__':
    ok = OkZiyuan('我是大仙尊')
    ok.run()
