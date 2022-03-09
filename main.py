#!/usr/bin/env python
# coding=utf-8

import contextlib
import mmap
import os

import openpyxl as oxl
from Evtx.Evtx import FileHeader
from Evtx.Views import evtx_file_xml_view
from lxml import etree

import myUtils


class AnalysisEvtx:
    def __init__(self, filePath):
        self.fileName = os.path.splitext(os.path.split(filePath)[1])[0]
        self.requireTagList = ["Provider.Name", "Provider.Guid", "EventID", "Level", "TimeCreated.SystemTime",
                               "EventRecordID", "Execution.ProcessID", "Execution.ThreadID", "Channel",
                               "Computer", "ProcessID", "Application", "Direction", "SourceAddress", "SourcePort",
                               "DestAddress", "DestPort", "Protocol", "RemoteUserID", "RemoteMachineID",
                               "Security.UserID", "QueryName", "EventSourceName", "Data", "Binary"]
        self.filePath = filePath
        self.startAnalysis()

    def convertEvtxToXmlList(self, filePath):
        resultDicList = []
        xmlList = []
        with open(filePath, 'r') as f:
            with contextlib.closing(mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)) as buf:
                fh = FileHeader(buf, 0)
                # 将一个事件转换为一个xml字符串，根元素是Events
                # 遍历事件
                nowCount = 1
                for xml, record in evtx_file_xml_view(fh):
                    print("解析第{0}条事件".format(nowCount))
                    reDic = self.analysisXml(xml)
                    xmlList.append(xml)
                    resultDicList.append(reDic)
                    nowCount = nowCount + 1
            print("解析完成\n")
        return xmlList, resultDicList

    def analysisXml(self, xmlStr):
        reDic = {}
        for tmpTagName in self.requireTagList:
            reDic[tmpTagName] = ""
        xmlObj = etree.XML(xmlStr)
        sysObj = xmlObj[0]
        dataObj = xmlObj[1]
        for item in sysObj:
            nowTag = "}".join(item.tag.split("}")[1:])
            if item.text is None:
                for tmpName, tmpValue in item.items():
                    tmpNowTag = nowTag + "." + tmpName
                    if self.checkIfRequire(tmpNowTag, self.requireTagList):
                        reDic = self.writeToDic(reDic, tmpNowTag, tmpValue)
                    else:
                        pass
            else:
                if self.checkIfRequire(nowTag, self.requireTagList):
                    reDic = self.writeToDic(reDic, nowTag, item.text)
                else:
                    pass
        for item in dataObj:
            tmpAttrDic = {key: value for key, value in item.items()}
            nowValue = item.text
            if "Name" in tmpAttrDic.keys():
                nowTag = tmpAttrDic["Name"]
                if not self.checkIfRequire(nowTag, self.requireTagList):
                    nowTag = "Data"
                    nowValue = tmpAttrDic["Name"]+":"+("" if item.text is None else item.text)
                else:
                    pass
            else:
                nowTag = "Data"
            reDic = self.writeToDic(reDic, nowTag, nowValue)
        return reDic

    def checkIfRequire(self, tagName, requireList):
        ifRequire = False
        if tagName in requireList:
            ifRequire = True
        else:
            ifRequire = False
        return ifRequire

    def writeToDic(self, aimDic, key, value):
        if value is None:
            value = ""
        if aimDic[key] == "":
            aimDic[key] = value
        else:
            aimDic[key] = aimDic[key] + "\n" + value
        return aimDic

    def startAnalysis(self):
        xmlList, resultDicList = self.convertEvtxToXmlList(self.filePath)
        self.writeDicListToFile(resultDicList)

    def writeDicListToFile(self, resultDicList):
        print("开始导出文件")
        # 创建一个excell文件对象
        wb = oxl.Workbook()
        # 创建URL扫描结果子表
        ws = wb.active
        ws.title = "{0} 解析结果".format(self.fileName)
        # 创建表头
        myUtils.writeExcellHead(ws, ["序号"] + self.requireTagList)
        # 写入内容
        for index, nowResultDic in enumerate(resultDicList):
            myUtils.writeExcellCell(ws, index + 2, 1, str(index + 1), 0, True)
            for colIndex, nowKey in enumerate(self.requireTagList):
                myUtils.writeExcellCell(ws, index + 2, colIndex + 2, str(nowResultDic[nowKey]), 0, True)
            myUtils.writeExcellSpaceCell(ws, index + 2, len(self.requireTagList) + 2)

        # 设置列宽

        colWidthArr = [10]
        for tmpKey in self.requireTagList:
            colWidthArr.append(20)
        myUtils.setExcellColWidth(ws, colWidthArr)

        # 保存文件
        fileName = "解析结果-{0}-{1}".format(self.fileName,
                                         myUtils.getNowSeconed().replace("-", "").replace(" ", "").replace(":", ""))
        myUtils.saveExcell(wb, saveName=fileName)
        print("成功导出文件：{0}.xlsx".format(fileName))


if __name__ == '__main__':
    ifContinue = "y"
    while ifContinue == "y":
        filePath = input("请输入需要解析的文件路径：")
        try:
            test = AnalysisEvtx(filePath)
        except Exception as ex:
            print("发生异常：" + str(ex))
        ifContinue = input("\n是否继续解析？(y/n):")
        ifContinue = "y" if ifContinue == "" else ifContinue
        print("")
