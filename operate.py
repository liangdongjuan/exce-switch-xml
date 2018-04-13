#!/usr/bin/python
#-*- coding:utf-8 -*-

import os
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
#tfrom easy_excel import easy_excel
import xlrd

class operate():





    def xlsx_to_dic(self,workbook,nu):
        self.workbook=workbook
        self.Sheetnum = self.workbook.nsheets
        self.dic_testlink = {}
        self.dick={}
        self.content = ""
        self.content_list = []
        index=nu
        sheetname = self.workbook.sheet_by_index(index)
        names = self.workbook.sheet_by_index(index).col_values(0, 1)
        rows = self.workbook.sheet_by_index(index).nrows

        number = []


        for index, name in enumerate(names):
            if name != '':
                number.append(index)
        for n in number:
            test = names[n]
            self.dick[test] = {"node_order": "13", "details": "", "testcase": []}
            if n != number[-1]:
                end = n + 1
            else:
                end = rows
            if n==0:
                n=1
            for i in xrange(n, end):
                testcase = {"node_order":"100","name": "", "externalid": "", "summary": "",
                            "preconditions": "", "version":"1","execution_type": "1", "importance": "3", "steps": [],
                            "keywords": "P1"}
                testcase["name"] = sheetname.cell_value(i, 1)
                if testcase["name"]=='':
                    testcase["name"]=test+str(i)
                testcase["summary"] = sheetname.cell_value(i, 1)
                testcase["preconditions"] = sheetname.cell_value(i, 4)

                testcase["execution_type"] = 2

                step_number = 1
                testcase["keywords"] = sheetname.cell_value(i, 1)
                # print testcase["keywords"]

                # print 'loop2'
                step = {"step_number": "", "actions": "", "expectedresults": "", "execution_type": ""}
                step["step_number"] = step_number
                step["actions"] = sheetname.cell_value(i, 5)
                step["expectedresults"] = sheetname.cell_value(i, 7)
                testcase["steps"].append(step)
                self.dick[test]["testcase"].append(testcase)
            print test
            self.dic_to_xml(test)

    def content_to_xml(self, key, value=None):
        if key == 'step_number' or key == 'execution_type' or key == 'node_order' or key == 'externalid' or key == 'version' or key == 'importance':
            return "<" + str(key) + "><![CDATA[" + str(value) + "]]></" + str(key) + ">"
        elif key == 'actions' or key == 'expectedresults' or key == 'summary' or key == 'preconditions':
            return "<" + str(key) + "><![CDATA[<p> " + str(value) + "</p> ]]></" + str(key) + ">"
        elif key == 'keywords':
            return '<keywords><keyword name="' + str(value) + '"><notes><![CDATA[ aaaa ]]></notes></keyword></keywords>'
        elif key == 'name':
            return '<testcase name="' + str(value) + '">'
        else:
            return '##########'

    def dic_to_xml(self,test):
        testcase_list = self.dick[test]["testcase"]
        for testcase in testcase_list:
            for step in testcase["steps"]:
                self.content += "<step>"
                self.content += self.content_to_xml("step_number", step["step_number"])
                self.content += self.content_to_xml("actions", step["actions"])
                self.content += self.content_to_xml("expectedresults", step["expectedresults"])
                self.content += self.content_to_xml("execution_type", step["execution_type"])
                self.content += "</step>"
            self.content = "<steps>" + self.content + "</steps>"
            self.content = self.content_to_xml("importance", testcase["importance"]) + self.content
            self.content = self.content_to_xml("execution_type", testcase["execution_type"]) + self.content
            self.content = self.content_to_xml("preconditions", testcase["preconditions"]) + self.content
            self.content = self.content_to_xml("summary", testcase["summary"]) + self.content
            self.content = self.content_to_xml("version", testcase["version"]) + self.content
            self.content = self.content_to_xml("externalid", testcase["externalid"]) + self.content
            self.content = self.content_to_xml("node_order", testcase["node_order"]) + self.content
            self.content = self.content + self.content_to_xml("keywords", testcase["keywords"])
            self.content = self.content_to_xml("name", testcase["name"]) + self.content
            self.content = self.content + "</testcase>"
            self.content_list.append(self.content)
            self.content = ""
        self.content = "".join(self.content_list)
        self.content = '<testsuite name="' + test + '">' + self.content + "</testsuite>"


    def write_to_file(self,ExcelFileName,SheetName):
        self.content = '<?xml version="1.0" encoding="UTF-8"?>' + self.content
        xmlFileName = ExcelFileName + '_' + SheetName + '.xml'
        cp = open(xmlFileName, "w")
        cp.write(self.content)
        cp.close()

if __name__ == "__main__":

    dir = os.path.abspath(".").decode("GBK")
    excelFile = os.path.join(dir, u"test.xlsx")
    workbook = xlrd.open_workbook(excelFile)
    Sheetnum = workbook.nsheets
    test = operate()
    for i in xrange(Sheetnum):
        test.xlsx_to_dic(workbook,i)
        sheetname="sheet"+str(i)
        test.write_to_file(sheetname,"bak")
