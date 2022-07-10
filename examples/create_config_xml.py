#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   create_config_xml.py
@Time    :   2022/06/15 09:37:05
@Author  :   Sean Bei 
@Version :   1.0
@Contact :   sean_bei@163.com
@Desc    :   This is used to create config.xml from a Excel.
             The idea is to read data from Excel sheet which is provided from customers, then 
             add one by one into the config.xml.
             When you run this python, you must add argv to specify the project name, 
             e.g. python create_config_xml.py PGS
             The project name must be the same in the Excel sheet, otherwise the creation will fail.
             If the creation is successful, you will get a xml file named "config_<projectname>.xml".
'''

# importing element tree
# under the alias of ET
import xml.etree.ElementTree as ET
import openpyxl
import os
import sys

SOURCE_FILE = 'Modbus-GiF.xlsx'

# Format xml file - indent and line feed
def pretty_xml(element, indent, newline, level=0):  # elemnt为传进来的Elment类，参数indent用于缩进，newline用于换行
    if element:  # 判断element是否有子元素    
        if (element.text is None) or element.text.isspace():  # 如果element的text没有内容
            element.text = newline + indent * (level + 1)
        else:
            element.text = newline + indent * (level + 1) + element.text.strip() + newline + indent * (level + 1)
            # else:  # 此处两行如果把注释去掉，Element的text也会另起一行
            # element.text = newline + indent * (level + 1) + element.text.strip() + newline + indent * level
    temp = list(element)  # 将element转成list
    for subelement in temp:
        if temp.index(subelement) < (len(temp) - 1):  # 如果不是list的最后一个元素，说明下一个行是同级别元素的起始，缩进应一致
            subelement.tail = newline + indent * (level + 1)
        else:  # 如果是list的最后一个元素， 说明下一行是母元素的结束，缩进应该少一个    
            subelement.tail = newline + indent * level
        pretty_xml(subelement, indent, newline, level=level + 1)  # 对子元素进行递归操作

# add address list
def add_address(_element, _id, _from, _address, _count, _func_code, _data_type, \
                _byte_order, _temp1_id, _temp2_id, _temp3_id, _station_id, _device_id, \
                _parameter, _precision, _scaling, _offset):
    address_list = ET.SubElement(_element, "address_list", {'max_entry_num': "65536"})
    list_id = ET.SubElement(address_list, "id")
    list_id.text = _id
    list_from = ET.SubElement(address_list, "from")
    list_from.text = _from
    address = ET.SubElement(address_list, "address")
    address.text = _address
    count = ET.SubElement(address_list, "count")
    count.text = _count
    func_code = ET.SubElement(address_list, "func_code")
    func_code.text = _func_code
    data_type = ET.SubElement(address_list, "data_type")
    data_type.text = _data_type
    byte_order = ET.SubElement(address_list, "byte_order")
    byte_order.text = _byte_order
    temp1_id = ET.SubElement(address_list, "temp1_id")
    temp1_id.text = _temp1_id
    temp2_id = ET.SubElement(address_list, "temp2_id")
    temp2_id.text = _temp2_id
    temp3_id = ET.SubElement(address_list, "temp3_id")
    temp3_id.text = _temp3_id
    station_id = ET.SubElement(address_list, "station_id")
    station_id.text = _station_id
    device_id = ET.SubElement(address_list, "device_id")
    device_id.text = _device_id
    parameter = ET.SubElement(address_list, "parameter")
    parameter.text = _parameter
    precision = ET.SubElement(address_list, "precision")
    precision.text = _precision
    scaling = ET.SubElement(address_list, "scaling")
    scaling.text = _scaling
    offset = ET.SubElement(address_list, "offset")
    offset.text = _offset

# add target modbus tcp list
def add_modbus_tcp(_element, _id, _ip, _port, _slave_id):
    modbus_tcp_list = ET.SubElement(_element, "modbus_tcp_list", {'max_entry_num': "100"})
    list_id = ET.SubElement(modbus_tcp_list, "id")
    list_id.text = _id
    ip = ET.SubElement(modbus_tcp_list, "ip")
    ip.text = _ip
    port = ET.SubElement(modbus_tcp_list, "port")
    port.text = _port
    slave_id = ET.SubElement(modbus_tcp_list, "slave_id")
    slave_id.text = _slave_id
 
# add basic config
def add_basic_config(_element, _product_key, _device_name, _device_secret, \
                    _username, _password, _interval, _max_fail_times, _retry_interval):
    product_key = ET.SubElement(_element, "product_key")
    product_key.text = _product_key
    device_name = ET.SubElement(_element, "device_name")
    device_name.text = _device_name
    device_secret = ET.SubElement(_element, "device_secret")
    device_secret.text = _device_secret
    username = ET.SubElement(_element, "username")
    username.text = _username
    password = ET.SubElement(_element, "password")
    password.text = _password
    interval = ET.SubElement(_element, "interval")
    interval.text = _interval
    max_fail_times = ET.SubElement(_element, "max_fail_times")
    max_fail_times.text = _max_fail_times
    retry_interval = ET.SubElement(_element, "retry_interval")
    retry_interval.text = _retry_interval

# add serial port info
def add_serial_info(_element, _baud_rate, _data_bits, _stop_bits, _parity):
    baud_rate = ET.SubElement(_element, "baud_rate")
    baud_rate.text = _baud_rate
    data_bits = ET.SubElement(_element, "data_bits")
    data_bits.text = _data_bits
    stop_bits = ET.SubElement(_element, "stop_bits")
    stop_bits.text = _stop_bits
    parity = ET.SubElement(_element, "parity")
    parity.text = _parity


def main():

    if len(sys.argv) == 1:
        print("Please enter the project name when run the script, e.g. python create_config_xml.py PGS ")
        print("The project name can be found in the Excel sheet.")
        return        

    if not os.path.exists(SOURCE_FILE):
        print("Can't find Excel file, please check if it exsits")
        return


    project_name = sys.argv[1]
    print("project_name:" + project_name)

    workbook = openpyxl.load_workbook(SOURCE_FILE)
    sheet = workbook['%s' % project_name]
    sheet_map = workbook['gif_parameters']
    if project_name == 'PGS':
        protocol_cells = sheet['A2':'P149']
    elif project_name == 'PPS':
        protocol_cells = sheet['A5':'P464']

    mapping_cells = sheet_map['B1': 'C2000']


    root = ET.Element('config_xml')
    cloud = ET.SubElement(root, "cloud")

    for index, address, parameter_user,temp1_id, temp2_id, temp3_id, station_id, device_id, \
        parameter_gif, data_type, endian, scaling, offset, bit, bit_length, description in protocol_cells:
        # currently only read is supported, write is not allowed
        # note that for PPS, neither 'R' nor 'W' is marked, as it is all readable.
#        if description.value[-1] == 'W' or description.value[-1] == None:
#            continue
        # get the mapped parameter name in database
        parameter_str = ''
        for x, y in mapping_cells:
            # There are some reserved addressed, the GiF parameter is empty.
            if parameter_gif.value == None:
                parameter_str = "Reserved"
                continue
            if parameter_gif.value == y.value:
                parameter_str = x.value

        if data_type.value == "BOOL":
            data_type_str = "select1"
        elif data_type.value == "BYTE":
            data_type_str = "select2"
        elif data_type.value == "BIT" or data_type.value == "DIGITAL":
            data_type_str = "select3"
        elif data_type.value == "WORD":
            data_type_str = "select4"
        elif data_type.value == "INT":
            data_type_str = "select5"
        elif data_type.value == "DWORD":
            data_type_str = "select6"
        elif data_type.value == "DINT":
            data_type_str = "select7"
        elif data_type.value == "REAL":
            data_type_str = "select8"

        if endian.value == None:
            endian_str = "select1"
        elif endian.value == "AB":
            endian_str = "select2"
        elif endian.value == "BA":
            endian_str = "select3"
        elif endian.value == "ABCD":
            endian_str = "select4"
        elif endian.value == "BADC":
            endian_str = "select5"
        elif endian.value == "CDAB":
            endian_str = "select6"
        elif endian.value == "DCBA":
            endian_str = "select7"

        if data_type.value == "BIT" or data_type.value == "DIGITAL":
            precision_str = str(bit.value)
        elif str(scaling.value).count(".") >= 1:
            precision_str = str(len(str(scaling.value).split(".")[1]))
        else:
            precision_str = "0"

        scaling_str = str(scaling.value)
        offset_str = str(offset.value)
        if scaling.value == None:
            scaling_str = "1"
        if offset.value == None:
            offset_str = "0"
        add_address(cloud, str(index.value), "0", str(address.value), "1", "select3", data_type_str, \
                    endian_str, temp1_id.value, temp2_id.value, temp3_id.value, station_id.value, \
                    device_id.value, parameter_str, precision_str, scaling_str, offset_str)


    pretty_xml(root, '\t', '\n')  # 执行美化方法
    target_content = ET.tostring(root)
    target_filename = "config_" + project_name + ".xml"
    with open(target_filename, 'wb') as f:
        f.write(target_content); 
    # tree.write("config_" + project_name + ".xml")
    print("config.xml is successfully updated according to the Excel file.")

if __name__ == "__main__":
    main()