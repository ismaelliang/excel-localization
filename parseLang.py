# -*- coding: utf-8 -*-

import sys
import os
import xlwt
from xlrd import open_workbook

# 支持的系统
SYS_IOS = 'iOS'
SYS_ANDROID = 'android'
SYS_TYPES = [SYS_IOS, SYS_ANDROID]

# 生成文件的根目录
ROOT_FOLDER = 'Langs/'

# 不同系统最终生成文件的名称 / 语言文件夹名
LANG_FILE_NAME = {
    SYS_IOS: 'Localizable.strings',
    SYS_ANDROID: 'strings.xml'
}
LANG_FOLDER_NAME = {
    SYS_IOS: {
        'en': 'en.lproj',
        'de': 'de.lproj',
        'fr': 'fr.lproj',
        'zh': 'zh_Hans.lproj'
    },
    SYS_ANDROID: {
        'en': 'values-en',
        'de': 'values-de',
        'fr': 'values-fr',
        'zh': 'values-zh-rCN'
    }
}


class ExcelObj():
    def __init__(self, parent, excel_name):
        try:
            self.workbook = open_workbook( excel_name )
        except IOError, e:
            print( 'invalid input' )
            return
        # Excel中各语言的排列顺序，从第二列开始
        self.config = {
            'col': [ 'zh', 'en', 'fr', 'de' ]
        }


    # 不同系统生成条目
    def parsed_item(self, type, key, value):
        if type is SYS_IOS:
            return '''"''' + key + '''" = "''' + value + '''";\n'''
        elif type is SYS_ANDROID:
            return '''\t<string name="''' + key + '''">''' + value + '''</string>\n'''

    # 准备需要使用全部文件夹
    def prepareDir(self):
        if not os.path.isdir( ROOT_FOLDER ):
            os.mkdir( ROOT_FOLDER )

        if not os.path.isdir(  ROOT_FOLDER + SYS_IOS + '/' ):
            os.mkdir( ROOT_FOLDER + SYS_IOS + '/' )
        if not os.path.isdir(  ROOT_FOLDER + SYS_ANDROID + '/' ):
            os.mkdir( ROOT_FOLDER + SYS_ANDROID + '/' )

    # 处理过程
    def parse(self):
        # 生成文件夹
        self.prepareDir()

        # 处理两种系统 iOS android
        for type in SYS_TYPES:

            # 处理各种语言
            for col in range( 1, len(self.config['col'])+1 ):

                result = ''

                # 处理每个sheet
                for s in range(0, len( self.workbook.sheet_names() )):
                    sheet = self.workbook.sheet_by_index(s)

                    # 处理每个条目
                    for row in range(1, sheet.nrows):

                        # 添加条目
                        try:
                            result += self.parsed_item( type, sheet.cell( row, 0).value.encode('utf-8'), sheet.cell( row, col ).value.encode('utf-8') )
                        except IndexError, e:
                            print( e )

                    result += "\n"
                if type is SYS_ANDROID:
                    result += "</resources>"


                # 生成语言文件
                folder = ROOT_FOLDER + type + "/" + LANG_FOLDER_NAME[type][self.config['col'][col-1]] + "/"
                if not os.path.isdir( folder ):
                    os.mkdir( folder )

                file = open( str( folder + LANG_FILE_NAME[type] ), "w+")
                try:
                    file.write( result )
                except IOError, e:
                    print('write file error')
                finally:
                    file.close()

                print result



if __name__ == "__main__":
    # 传入一个参数作为要处理的excel的文件名
    if len( sys.argv ) > 1:
        excel = sys.argv[1]

        excelObj = ExcelObj( None, excel )
        excelObj.parse()

    else:
        print('excel name required')