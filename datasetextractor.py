import time

import openpyxl
import yaml


# define template exception
class TemplateException(Exception):
    pass


# define template set value exception
class TemplateSetValueException(Exception):
    pass


# define source file exception
class SourceFileException(Exception):
    pass


# define extract value exception
class ExtractValueException(Exception):
    pass


# class which handles extracting the data for a folder (which defines a set)
class DataSetExtractor:
    # constructor
    def __init__(self, folder_name):
        print('____________________\nPROCESSING ' + folder_name)
        self.folder_name = folder_name

    # process each extract set
    def extract(self):
        config_filename = self.folder_name + 'extract_config.yaml'

        try:
            # read the extract set config
            with open(config_filename) as folder_config_file:
                extract_config = yaml.full_load(folder_config_file)
            folder_config_file.close()

            workbook = self.__get_template_workbook(extract_config['templateFilename'])

        except TemplateException:
            print('ERROR reading template file in ' + config_filename)
        except Exception as ex:
            print('ERROR trying to process ' + config_filename)
            print(ex.args)
        else:
            # keep an error count
            error_count = 0

            # loop through and process each extract detail
            for extract_detail in extract_config['extractDetails']:
                try:
                    # extract value
                    value = self.__extract_value(extract_detail)

                    # set value to workbook
                    self.__set_value_to_template(workbook,
                                                 extract_detail['templateSheet'],
                                                 extract_detail['templateCell'],
                                                 value)
                except SourceFileException:
                    print('ERROR trying to read source file ' + self.folder_name + extract_detail['sourceFilename'])
                    error_count += 1
                except ExtractValueException:
                    print('ERROR trying to extract value from')
                    print(extract_detail)
                    error_count += 1
                except TemplateSetValueException:
                    print('ERROR trying to set value to template')
                    print(extract_detail)
                    error_count += 1

            if not error_count:
                # save the workbook
                self.__write_template(workbook, extract_config['templateFilename'])
            else:
                print('ERROR count encountered: ' + str(error_count) + ', while processing ' + self.folder_name)

    # write template
    def __write_template(self, workbook, template_filename):
        # write as a new file with timestamp
        timestamp = time.strftime('%Y%m%d_%H%M', time.localtime())
        filename = self.folder_name + timestamp + '_' + template_filename

        workbook.save(filename)
        print('SUCCESSFULLY written out ' + filename)

    # extract single value given folder and details
    def __extract_value(self, extract_detail):
        try:
            source_file = open(self.folder_name + extract_detail['sourceFilename'])
        except OSError:
            raise SourceFileException()

        # read lines until key is found
        found = False
        while True:
            line = source_file.readline()

            # if reached eof, break out of loop
            if not line:
                break

            # check if line contains the key
            index = line.find(extract_detail['key'])
            if index != -1 and index == extract_detail['keyColStart'] - 1:
                found = True
                break

        if not found:
            # if value not found, raise and exception
            raise ExtractValueException()
        else:
            # find value line by skipping offset
            for x in range(extract_detail['valueRowOffset']):
                line = source_file.readline()

            # line at this point should contain value
            source_file.close()
            value = line[extract_detail['valueColStart'] - 1:
                         extract_detail['valueColStart'] - 1 + extract_detail['valueColLength']]

            if value is None:
                # if not successful extracting value, raise exception
                raise ExtractValueException()
            else:
                # return the value found
                print('Read value [' + value + '] from ')
                print(extract_detail)
                return float(value)

    # get the workbook
    def __get_template_workbook(self, filename):
        try:
            workbook = openpyxl.load_workbook(self.folder_name + filename)
            return workbook
        except Exception:
            raise TemplateException()

    # write value to the template
    @staticmethod
    def __set_value_to_template(workbook, sheet_num, cell_address, value):
        try:
            worksheet = workbook[workbook.sheetnames[sheet_num - 1]]
            worksheet[cell_address] = value
        except Exception:
            raise TemplateSetValueException()
