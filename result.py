import os
import requests
import sys
from bs4 import BeautifulSoup
from collections import OrderedDict
from Queue import Queue
from threading import Thread


def get_file_name(base_dir, student):
    return os.path.join(base_dir, "%s-%s.html" % (student['fac_no'], student['name']))


def parse_type(data):
    data = str(data)

    if data.isdigit():
        data = int(data)
    elif '.' in data and data.replace('.', '', 1).isdigit():
        data = float(data)

    return data


class Downloader(object):
    __base_url_format = "http://ctengg.amu.ac.in/web/table_resultnew.php?fac=%s&en=%s&prog=btech"
    __store_path = 'store'
    __concurrent = 200
    __retry_counter = {}

    def __init__(self, student_list):
        if not os.path.exists(self.__store_path):
            os.mkdir(self.__store_path)
        self.students = student_list
        self.q = Queue(self.__concurrent * 2)
        for _ in xrange(self.__concurrent):
            thread = Thread(target=self.consume)
            thread.daemon = True
            thread.start()

    def consume(self):
        while True:
            student = self.q.get()
            if not self.process(student):
                self.q.put(student)
            self.q.task_done()

    def retry(self, en_no):
        if en_no in self.__retry_counter:
            if self.__retry_counter[en_no] > 5:
                print 'Tried downloading results many times for enrolment number %s... Giving up now. Sorry :\'(' % en_no
                return True

            self.__retry_counter[en_no] += 1
        else:
            self.__retry_counter[en_no] = 0

        return False


    def process(self, student):
        file_path = get_file_name(self.__store_path, student)
        fac_no = student['fac_no']
        name = student['name']
        if os.path.isfile(file_path):
            print "Result for %s - %s already exists. Skipping..." % (fac_no, name)
        else:
            en_no = student['en_no']

            response = requests.get(
                self.__base_url_format % (fac_no, en_no))

            content = response.text
            if response and 'CPI' in content:
                out = open(file_path, 'w+')
                out.write(content)
                out.close()
                print 'Result Downloaded %s' % file_path
            elif response and 'Student record not found' in content:
                print 'Either result is unavailable or error in credentials %s %s %s' % (fac_no, en_no, student['name'])
            else:
                return self.retry(en_no)
        return True

    def download(self):
        print 'Started downloading...'
        try:
            for student in self.students:
                self.q.put(student)
            self.q.join()
        except KeyboardInterrupt:
            sys.exit(1)

    def get_store_dir(self):
        return self.__store_path


class Parser(object):
    __base_dir = None

    def __init__(self, student_list, base_dir):
        self.students = student_list
        self.__base_dir = base_dir
        if not os.path.isdir(self.__base_dir):
            raise ValueError('Invalid store directory passed')

    @staticmethod
    def parse_page(page):
        table = BeautifulSoup(page, "html.parser")
        credit_keys = ['faculty_number',
                       'enrolment', 'name', 'ec', 'spi', 'cpi']
        dataset = None

        cred_table = table.find(
            'table', {'style': 'width:100%;text-align:center;'})
        for row in cred_table.find_all('tr')[1:]:
            dataset = OrderedDict(
                zip(credit_keys, (parse_type(cell.get_text()) for cell in row.find_all('td'))))

        return dataset

    def get_results(self):
        results = []
        for student in self.students:
            file_path = file_path = get_file_name(self.__base_dir, student)
            try:
                data = open(file_path, 'r')
                results.append(Parser.parse_page(data.read()))
                data.close()
            except IOError:
                print "Error opening file %s" % file_path
        return results
