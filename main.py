import json
import result
import student
import sys
from operator import itemgetter


def get_input_file():
    args = sys.argv[1:]
    if len(args) >= 1:
        return args[0]
    else:
        return raw_input('Please enter the spreadsheet name containing the students credentials: ')


def ask_print(json):
    should_print = raw_input('Print the parsed result json? (Y/N): ')

    if not (should_print == 'N' or should_print == 'n'):
        print json


def save_results(json):
    out_name = raw_input(
        'Please enter the name you want for the output file: ')

    try:
        out = open(out_name + '.json', 'w+')
        out.write(json)
        out.close()
    except IOError:
        print 'Error in writing output'


if __name__ == "__main__":
    file_name = get_input_file()
    print file_name

    collector = student.Collector(file_name)
    students = collector.collect()
    print "No of students: %d" % len(students)
    downloader = result.Downloader(students)
    downloader.download()

    print '\nParsing... '
    parser = result.Parser(students, downloader.get_store_dir())
    result = sorted(parser.get_results(), key=itemgetter(
        'cpi', 'spi'), reverse=True)

    result_json = json.dumps(result, indent=2)
    ask_print(result_json)

    save_results(result_json)
