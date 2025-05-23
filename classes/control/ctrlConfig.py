import os
# import xml.etree.ElementTree as et
from lxml import etree

from classes.control import ctrlCommon


class ctrlConfig():
    # Get XML data
    # def get_xmldata(filename):
    def get_xmldata(xml_path):
        xmldata = None
        
        # Check if file exists
        if not (os.path.isfile(xml_path)):
            return None
        # Check file extension
        if not (xml_path.lower().endswith("xml")):
            return None
        # Read file
        try:
            parser_enc = etree.XMLParser(encoding='UTF-8', recover=True)
            xmldata = (etree.parse(xml_path, parser=parser_enc)).getroot()
        except OSError:
            print('Exception error: XML reading')
        return xmldata

    # Read XML file
    def read_xmlfile(settings, xmldata):
        if not (xmldata is None):
            settings['hash_algorithm'] = xmldata.find(
                './basicsettings/hash_algorithm').text
            settings['package_maxsize'] = xmldata.find(
                './basicsettings/package_maxsize').text
            settings['package_maxfiles'] = xmldata.find(
                './basicsettings/package_maxfiles').text
            settings['check_maxsize'] = xmldata.find(
                './basicsettings/check_maxsize').text
            settings['check_maxfiles'] = xmldata.find(
                './basicsettings/check_maxfiles').text
            settings['password'] = xmldata.find(
                './basicsettings/password').text
        return settings
