import os

import yaml

import datasetextractor


class DataExtractor:
    # start extract
    @staticmethod
    def extract():
        try:
            # read in the set of extract folders
            with open(os.getcwd() + '\\app_config.yaml') as app_config_file:
                extract_folders = yaml.full_load(app_config_file)

            app_config_file.close()
        except OSError:
            print('ERROR trying to read ' + os.getcwd() + '\\app_config.yaml')
        else:
            # loop and process each extract folder
            for extract_folder in extract_folders['extractFolders']:
                # process each extract folder by creating a new data set extractor instance, then calling extract
                datasetextractor.DataSetExtractor(extract_folder).extract()


# main function
def main():
    # start data extractor by calling static method extract
    DataExtractor.extract()


# execute only if run as a script
if __name__ == '__main__':
    main()
