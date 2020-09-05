import yaml
# import openpyxl
import time


# start by reading in the set of extract folders
def extract():
    with open('C:/Projects/pyDataExtractor/resources/config/appconfig.yaml') as app_config_file:
        extract_folders = yaml.full_load(app_config_file)
        for extract_folder in extract_folders['extractFolders']:
            extract_set(extract_folder)


# process each extract set
def extract_set(extract_folder):
    print('____________________\nPROCESSING ' + extract_folder)

    # read the extract set config
    with open(extract_folder + 'extractconfig.yaml') as set_config_file:
        extract_config = yaml.full_load(set_config_file)

    workbook = get_template_workbook(extract_folder, extract_config["templateFilename"])

    # loop through and process each extract detail
    for extract_detail in extract_config['extractDetails']:
        value = extract_value(extract_folder, extract_detail)

        if value is None:
            print('ERROR trying to extract value from ')
            print(extract_detail)

        set_value_to_template(workbook, extract_detail["templateSheet"], extract_detail["templateCell"], value)

    # save the workbook
    write_template(workbook, extract_folder, extract_config["templateFilename"])


# write template
def write_template(workbook, extract_folder, template_filename):
    timestamp = time.strftime("%Y%m%d%H%M", time.localtime())
    filename = extract_folder + timestamp + '_' + template_filename

    workbook.save(filename)
    print('SUCCESSFULLY written out ' + filename)


# extract single value given folder and details
def extract_value(extract_folder, extract_detail):
    source_file = open(extract_folder + extract_detail['sourceFilename'])
    line = None

    # read lines until key is found
    found = False
    for x in source_file:
        line = source_file.readline()
        index = line.find(extract_detail['key'])
        if index != -1 and index == extract_detail['keyColStart'] - 1:
            found = True
            break

    if not found:
        print('ERROR trying to read value ')
        print(extract_detail)

    # find value line by skipping offset
    for x in range(extract_detail['valueRowOffset']):
        line = source_file.readline()

    # line at this point should contain value
    value = line[extract_detail['valueColStart'] - 1:
                 extract_detail['valueColStart'] - 1 + extract_detail['valueColLength']]

    if value is None:
        print('ERROR trying to read value ')
    else:
        print('Read value [' + value + '] from ')
    print(extract_detail)

    return float(value)


# get the workbook
def get_template_workbook(extract_folder, filename):
    from openpyxl import load_workbook
    workbook = load_workbook(extract_folder + filename)

    if workbook is None:
        print('ERROR trying to reading template ' + extract_folder + filename)

    return workbook


# write value to the template
def set_value_to_template(workbook, sheet_num, cell_addr, value):
    worksheet = workbook[workbook.sheetnames[sheet_num - 1]]
    worksheet[cell_addr] = value


# start extract
extract()
