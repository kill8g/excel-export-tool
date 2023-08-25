#coding:utf-8

import ast
import tkinter.messagebox as messagebox
import xlwings as xw
from time import time
import os

def encode_int(value):
    try:
        return True, int(float(value))
    except ValueError as e:
        return False, None

def encode_float(value):
    try:
        return True, float(value)
    except ValueError as e:
        return False, None

def encode_string(value):
    try:
        return True, value
    except ValueError as e:
        return False, None

def encode_json(value):
    try:
        ast.literal_eval(value)
        return True, value
    except ValueError as e:
        return False, None

def encode_dict(value):
    try:
        return True, ast.literal_eval(value)
    except ValueError as e:
        return False, None

def encode_bool(value):
    try:
        return True, bool(value)
    except ValueError as e:
        return False, None

TypeEncodeInterface = {
    'int'       : encode_int, 
    'float'     : encode_float, 
    'string'    : encode_string, 
    'json'      : encode_json, 
    'dict'      : encode_dict, 
    'bool'      : encode_bool, 
}

DefaultValue = {
    'int'       : 0, 
    'float'     : 0.0, 
    'string'    : '', 
    'json'      : '{}', 
    'dict'      : {}, 
    'bool'      : False, 
}

def colour31(text):
    s = '\033[0;31;40m%s\033[0m' % (text)
    return s

def colour32(text):
    s = '\033[0;32;40m%s\033[0m' % (text)
    return s

def colour33(text):
    s = '\033[0;33;40m%s\033[0m' % (text)
    return s

def colour34(text):
    s = '\033[0;34;40m%s\033[0m' % (text)
    return s

class TableOP():
    def __init__(self, path, output, user = None, single = None):
        self.path = path # xlsx file path
        self.output = output # output file path or database host

        if self.output[-1] == '/':
            self.output = self.output[:-1]

        self.user = user
        self.single = single

        self.app = xw.App(visible = False, add_book = False)
        self.app.display_alerts = False
        self.app.screen_updating = False
        self.wb = self.app.books.open(self.path)

    def __del__(self):
        self.wb.close()
        self.app.quit()

    def error_tips(self, msg, sheet = None, row = None, column = None, value = None):
        log = "<%s>" % (self.path)
        if sheet != None:
            log = "%s <%s>" % (log, sheet.name)

        if row != None and column != None:
            log = "%s row %d column %d" % (log, row + 1, column + 1)
        if value != None:
            log = "%s excel content : %s" % (log, value)

        log = "%s, error msg : %s" % (log, msg)
        print("[Error]", log)
        messagebox.showerror('Error', log)
        exit(0)

    def decode_type(self, _type):
        if _type.find('@') != -1:
            _type, decoration = _type.split('@', 1)
            return _type, decoration or None
        else:
            return _type, None
        
    def get_key_field(self, head):
        for column, member in head.items():
            if member['key']:
                return column
        return None
    
    def is_empty_column(self, sheet, row):
        for column in range(0, sheet.used_range.shape[1]):
            cell = sheet[row, column]
            if cell.value != None:
                return False
        return True

    def mkdir(self):
        if not os.path.exists(self.output):
            os.mkdir(self.output)

    def read_sheets(self, names):
        if len(names) == 0:
            return self.wb.sheets
        else:
            array = []
            for name in names:
                sheet = self.wb.sheets[name]
                if sheet:
                    array.append(sheet)
                else:
                    self.error_tips('Not found sheet <%s>' % (name))
            return array

    def read_head(self, sheet):
        nrow = sheet.used_range.shape[0]
        ncolumn = sheet.used_range.shape[1]
        if nrow == 0:
            return # empty

        has_key = False
        head = {}
        types = sheet[0, :ncolumn].value
        fields = sheet[1, :ncolumn].value
        details = sheet[2, :ncolumn].value
        for column in range(0, ncolumn):
            tmp = types[column]
            if tmp == None:
                head[column] = None
            else:
                type_value = tmp.strip()
                _type, decoration = self.decode_type(type_value)
                if has_key and decoration == 'key':
                    self.error_tips('Multiple keys exist', sheet = sheet, row = 0, column = column, value = value)

                has_key = decoration == 'key'
                head[column] = {
                    'type'      : _type, 
                    'field'     : fields[column].strip(), 
                    'desc'      : details[column].strip(), 
                    'default'   : decoration == 'default', 
                    'key'       : decoration == 'key', 
                    'ignore'    : decoration == 'ignore', 
                }
        return head

    def read_body(self, head, sheet):
        nrow = sheet.used_range.shape[0]
        ncolumn = sheet.used_range.shape[1]
        key_idx = table.get_key_field(head)
        body = {}

        for row in range(3, nrow):
            if sheet[row, 0].api.EntireRow.Hidden:
                print('[warning] %s %s line %s %s, skip this line.' % (self.path, sheet.name, colour34('<%d>' % (row)), colour31('is set to hidden')))
                continue
            values = sheet[row, :ncolumn].value
            is_empty = all(x is None or x == "" for x in values)
            if is_empty:
                print('[warning] %s %s line %s %s, skip this line.' % (self.path, sheet.name, colour34('<%d>' % (row)), colour32('is empty')))
                continue
            data = {}
            ignores = {}
            for column, value in enumerate(values):
                value = str(value).strip()
                tmp = head[column]
                if not tmp:
                    continue
                if value == '@ignore':
                    ignores[row] = True
                    print('[warning] %s %s the beginning of line %s %s, which ignores this row of data.' % (self.path, sheet.name, colour34('<%d>' % (row)), colour33('is set to @ignore')))
                    continue

                if row in ignores:
                    continue

                field = tmp['field']
                _type = tmp['type']
                default = tmp['default']
                if value == 'None' or len(value) == 0:
                    if default:
                        value = DefaultValue[_type]
                    else:
                        self.error_tips('No default value', sheet = sheet, row = row, column = column, value = value)
                else:
                    decode_func = TypeEncodeInterface[_type]
                    ok, result = decode_func(str(value))
                    if not ok:
                        self.error_tips('Type error, conversion failure', sheet = sheet, row = row, column = column, value = value)
                    data[field] = result
                column = column + 1
            if key_idx != None:
                field = head[key_idx]['field']
                if data:
                    body[data[field]] = data
            else:
                body[row - 3] = data
        return body

    def to_xml(self, sheet):
        self.mkdir()
        from dicttoxml import dicttoxml
        from xml.dom.minidom import parseString
        for sheet in self.read_sheets(names):
            head = table.read_head(sheet)
            body = table.read_body(head, sheet)
            xml = dicttoxml(body)
            dom = parseString(xml)
            pretty_xml = dom.toprettyxml()
            with open(self.output + '/' + sheet.name + '.xml', "w") as f:
                f.write(pretty_xml)
                f.close()

    def to_json(self, names):
        import json
        self.mkdir()
        for sheet in self.read_sheets(names):
            head = table.read_head(sheet)
            body = table.read_body(head, sheet)
            with open(self.output + '/' + sheet.name + '.json', "w") as f:
                f.write(json.dumps(body, indent = 4))
                f.close()
    
    def to_lua(self, names):
        self.mkdir()
        import luadata
        for sheet in self.read_sheets(names):
            head = table.read_head(sheet)
            body = table.read_body(head, sheet)
            content = luadata.serialize(body, encoding = 'utf-8', indent = '\t', indent_level = 0)
            with open(self.output + '/' + sheet.name + '.lua', 'w') as f:
                f.write('return ' + content)
                f.close()

    def to_mongo(self, names):
        import pymongo
        import json
        host, dbinfo = self.output.split('@', 1)
        if dbinfo.find(':') != -1:
            dbname, single = dbinfo.split(':')
            single = single == '1'
        else:
            dbname = dbinfo
            single = False

        client = pymongo.MongoClient('mongodb://%s' % (host))
        db = client[dbname]
        if self.user:
            db = client.admin
            user, pwd = self.user.split('@', 1)
            db.authenticate(user, pwd)
        for sheet in self.read_sheets(names):
            collection = db[sheet.name]
            head = table.read_head(sheet)
            body = table.read_body(head, sheet)
            if single:
                collection.create_index('name')
                collection.update_one({'name' : sheet.name}, {'$setOnInsert' : {'data' : json.dumps(body) }}, upsert = True)
            else:
                key_idx = table.get_key_field(head)
                field = head[key_idx]['field']
                collection.create_index(field)
                for k, v in body.items():
                    collection.update_one({field : v[field]}, {'$setOnInsert' : v}, upsert = True)

    def to_sqlite(self, names):
        import sqlite3
        import json
        if self.output.find('@') != -1:
            db, single = self.output.split('@', 1)
            single = single == '1'
        else:
            db = self.output
            single = False

        for sheet in self.read_sheets(names):
            head = table.read_head(sheet)
            body = table.read_body(head, sheet)
            
            sql = 'CREATE TABLE %s\n' % (sheet.name)
            sql = sql + '(\n'
            if single:
                sql = sql + '''    name VARCHAR PRIMARY KEY NOT NULL,\n'''
                sql = sql + '''    data BLOB NULL\n'''
            else:
                for k, v in head.items():
                    if not v:
                        continue
                    _type = v['type']
                    if v['key']:
                        if _type == 'json' or _type == 'dict':
                            self.error_tips('The <%s> field cannot be the primary key' % (v['field']))
                        else:
                            sql = sql + '''    %s %s PRIMARY KEY NOT NULL,\n''' % (v['field'], _type)
                    else:
                        if _type == 'json' or _type == 'dict':
                            sql = sql + '''    %s BLOB NULL,\n''' % (v['field'])
                        else:
                            sql = sql + '''    %s %s NULL,\n''' % (v['field'], _type)
                sql = sql[:-2]

            sql = sql + ');'
            conn = sqlite3.connect(db)

            try:
                conn.execute(sql)
            except:
                print('create table failed! table already exists.')

            cur = conn.cursor()
            if single:
                sql = 'REPLACE INTO %s(name, data) VALUES(?, ?)' % (name)
                cur.execute(sql, (name, json.dumps(body)))
            else:
                keys = ''
                values = ''
                sql = 'REPLACE INTO %s(' % (sheet.name)
                fields = []
                for k, v in head.items():
                    if not v:
                        continue
                    field = v['field']
                    fields.append(field)
                    values = values + '?,'
                    keys = keys + field + ','
                sql = sql + keys[:-1] + ') VALUES ('
                sql = sql + values[:-1] + ')'

                for k, v in body.items():
                    values = []
                    for field in fields:
                        if field in v:
                            if type(v[field]) == type({}):
                                values.append(json.dumps(v[field]))
                            else:
                                values.append(v[field])
                        else:
                            values.append(None)
                    cur.execute(sql, values)
            conn.commit()
            conn.close()


import argparse
if __name__ == '__main__':
    parse = argparse.ArgumentParser()
    parse.add_argument('-m', '--mode', help = 'excel data export format : xml, json, lua, mongo, sqlite', required = True)
    parse.add_argument('-o', '--output', help = 'excel data export path : relative path or database connection information, single is equal to 1 or 0, Indicates that data in the database is stored as a single piece of data or multiple pieces of data, example : ./test or localhost:27017@database:single', required = True)
    parse.add_argument('-f', '--file', help = 'excel file name : example : path/file.xlsx or path/file.xls', required = True)
    parse.add_argument('-n', '--names', help = 'sheet list : excel sheet names, english comma separation, example : Sheet1,Sheet2')
    parse.add_argument('-u', '--user', help = 'mongo or mysql user info : user@pwd', required = False)

    args = parse.parse_args()
    mode = args.mode.strip()
    output = args.output.strip()
    file = args.file.strip()
    names = args.names.split(',')

    print('Start processing %s' % (colour31(file)))
    if mode == 'mongo' or mode == 'sqlite':
        table = TableOP(file, output, user = args.user)
        if mode == 'mongo':
            table.to_mongo(names)  
        elif mode == 'sqlite':
            table.to_sqlite(names)
    elif mode == 'lua':
        table = TableOP(file, output)
        table.to_lua(names)
    elif mode == 'json':
        table = TableOP(file, output)
        table.to_json(names)
    elif mode == 'xml':
        table = TableOP(file, output)
        table.to_xml(names)

