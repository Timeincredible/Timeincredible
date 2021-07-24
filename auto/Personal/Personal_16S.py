'''
Descripttion: 派森诺
Author: wwt
Date: 2021-07-15 09:13:02
LastEditors: wwt
LastEditTime: 2021-07-22 13:22:33
'''
import re
import os
import sys
import glob
from bs4 import BeautifulSoup
from shutil import copyfile, copytree, rmtree

SUCCESS_FLAG = "OK"
FAILED_FLAG = "NG"
sys.path.append("../")
from Public.public import BaseHTML

# =========== 派诺森 16S 分析 相关解析类================
#
# 1. AnalysisReportHTML 解析 检测报告 report.html AnalyisMain 主接口
# 2. QualityReportHTML 解析 扩增报告 index.html  QualityMain 主接口
#
# ==============================================
class AnalysisReportHTML(BaseHTML):
    '''修改 report.html'''

    def __init__(self, htmlPath, jsonPath):
        # 调用父类
        super().__init__(jsonPath)

        self.htmlPath = htmlPath

        with open(self.htmlPath, "r", encoding="utf-8")as f:
            htmlfile = f.read()

        self.soup = BeautifulSoup(htmlfile, "html.parser")

    def main(self):

        # 删除：
        # 1. 删除所有的注释
        soup = re.sub("<!--\n?(.|[\r\n])*?-->", "", str(self.soup))

        # 2. 删除 委托单位和制定日期
        soup = re.sub("<p>委托单位：\w+</p>", "",  str(self.soup))
        soup = re.sub("<p>制定日期：\w{4}年\w{0,2}月\w{0,2}日</p>", "", soup)

        # ==========  严格按照下列顺序删除： ============
        # 1.删除全文的logo.png
        soup = re.sub('src=".*logo.png', "", soup)
        # 2.删除http://www.personalbio.cn
        soup = re.sub("http://www.personalbio.cn", "", soup)
        # 3.删除 www.personalbio.cn
        soup = re.sub("www.personalbio.cn", "", soup)

        # 删除 开题单号
        soup = re.sub("<tr><th>开题单号</th><th>.*</th>", "", soup)
        # 删除 分析者工号
        soup = re.sub("<tr><th>分析者工号</th><th>.*</th>", "", soup)

        # 修改 项目编号

        ItemNo = self.jsonData.get("ItemNo", "null")                # 项目编号
        soup = re.sub("<tr><th>﻿项目编号</th><th>.*</th>",
                      "<tr><th>项目编号</th><th>%s</th>" % (ItemNo), soup)

        # 修改 完成日期
        ReceivedDate = self.jsonData.get("ReceivedDate", "null")   # 完成日期
        soup = re.sub("<tr><th>完成日期</th><th>\d{4}-\d{0,2}-\d{0,2}</th>",
                      "<tr><th>完成日期</th><th>%s</th>" % (ReceivedDate), soup)

        self.soup = BeautifulSoup(soup, "lxml")

        # 删除 div
        self.delTag("div", **{"class": "tel-left"})
        self.delTag("div", **{"class": "float_btn float-mouse"})

        # 7/21 新增需求
        # 1. 删除<div class="floatbar-info float-mouse">
        self.delTag("div", **{"class": "floatbar-info float-mouse"})
        # 2. 删除 <img src="static/icon/page_bg.jpg"/>
        self.delTag("img", **{"src": "static/icon/page_bg.jpg"})

        # 3. 删除 文件下载 <div class="floatbar-wrap" hidden="">
        self.delTag("div", **{"class": "floatbar-wrap","hidden":""})

        # 4. 删除 <img alt="" class="tel-right" src="static/icon/2Dplot_1.jpg"/>
        self.delTag("img", **{"class": "tel-right","src":"static/icon/2Dplot_1.jpg"})

        # 5. 删除
        # <div class="float_tips" hidden=""></div>
        self.delTag("div", **{"class":"float_tips"})
        float_tips = self.soup.find_all("div",**{"class":"float_tips"})
        for float_tip in float_tips:
            float_tip.decompose()

        # 写入新的文件
        self.writeFile()

        return SUCCESS_FLAG


def PersonalAnalyisMain(path, parentDir):
    try:
        # ================== 事前处理 ==================
        anapath = path.rstrip("\\").rstrip("/")

        # Treat文件可能为多个
        Treats = glob.glob(anapath + "/Treat[0-9]*")

        for Treats in Treats:
            # ============  step-1 : 文件删除 ===========
            # 1.1 删除所有的带有logo字样的图片
            imagePath = Treats + "/images/"
            images = glob.glob(imagePath+"*logo*.png")
            for image in images:
                os.remove(image)

            # 1.2. 删除所有 tables.js 和tables.js.bak
            tablesJs = Treats + "/static/js/tables.js"
            if os.path.exists(tablesJs):
                os.remove(tablesJs)
            tablesJsbak = Treats + "/static/js/tables.js.bak"
            if os.path.exists(tablesJsbak):
                os.remove(tablesJsbak)

            # del 2Dplot_*.png
            Dplot1_jpg = Treats + "/static/icon/2Dplot_1.jpg"
            if os.path.exists(Dplot1_jpg):
                os.remove(Dplot1_jpg)

            Dplots = glob.glob(Treats + "/static/icon/2Dplot_*.png")
            for Dplot in Dplots:
                if os.path.exists(Dplot):
                    os.remove(Dplot)

            # del page_bg.png
            page_bg = Treats + "/static/icon/page_bg.jpg"
            if os.path.exists(page_bg):
                os.remove(page_bg)

            #3 . Logo替换
            logoPath = parentDir + "/replaceFiles/logo.png"
            oldlogoPath = Treats + "/static/icon/logo.png"
            if os.path.exists(oldlogoPath):
                os.remove(oldlogoPath)
            copyfile(logoPath, oldlogoPath)

            # 1.4.如果在Treat中有report文件夹，删除整个文件夹，因为和外面的报告是重复的。
            reportDir = Treats + "/report"
            if os.path.exists(reportDir):
                rmtree(reportDir)
            
            # 1.5 删除general 下的 general.txt
            generalDir = Treats + "/general"
            if os.path.exists(generalDir):
                general_txt = generalDir + "/general.txt"
                if os.path.exists(general_txt):
                    os.remove(general_txt)

            # ============ step-2 :网页修改 ============
            reportPath = Treats + "/report.html"
            jsonPath = parentDir + "/config/Personal16S_analysis.json"
            report = AnalysisReportHTML(reportPath, jsonPath)
            report.main()

        return SUCCESS_FLAG

    except Exception as e:
        print(e)
        return FAILED_FLAG


class QualityReportHTML(BaseHTML):
    def __init__(self, htmlPath, jsonPath):
        # 调用父类
        super().__init__(jsonPath)

        self.htmlPath = htmlPath

        with open(self.htmlPath, "r", encoding="utf-8")as f:
            self.soup = f.read()

    def main(self, anapath, parentDir):
        try:
            # ================ 检测报告 ================
            # # 一、文件替换
            #   2. src\images中  Logo替换
            logoPath = parentDir + "/replaceFiles/logo.png"
            oldlogoPath = anapath + "/src/images/logo.png"
            if os.path.exists(oldlogoPath):
                os.remove(oldlogoPath)
            copyfile(logoPath, oldlogoPath)

            # =================== 二、网页修改 #===================
            # 1.报告时间无需修改。
            # 2.替换
            #   --（1）派森诺生物科技股份有限公司 替换为 苏州帕诺米克生物医药科技有限公司
            #   --（2）派森诺 替换为 帕诺米克。
            replaceDatas = self.jsonData.get("replaceFields")
            for replaceData in replaceDatas:
                name = replaceData.get("key")
                replaceName = replaceData.get("value")
                self.replaceText(name, replaceName)

            # 项目基础信息修改
            self.editProinfo()

            # 写入文件
            self.writeFile()

            return SUCCESS_FLAG
        except Exception as e:
            print(e)
            return FAILED_FLAG


def PersonalQualityMain(path, parentDir):
    # ================== 事前处理 ==================

    # ================== 扩增报告 ==================
    basehtml = BaseHTML()

    anapath = path.rstrip("\\").rstrip("/")

    reportPath = anapath + "/扩增报告/index.html"
    jsonPath = parentDir + "/config/Personal16S_quality.json"
    # 1. src\4 中如果有PDF需要删除页眉页脚，保留页码
    pdf_files = anapath + "/扩增报告/src/4"
    basehtml.replaceHeardFooter(pdf_files)
    detection_path = anapath +"/扩增报告"

    report = QualityReportHTML(reportPath, jsonPath)
    report.main(detection_path, parentDir)

    # ================== 检测报告 ==================
    reportPath = anapath + "/检测报告/index.html"
    jsonPath = parentDir + "/config/Personal16S_quality.json"

    # 1. src\4 中如果有PDF需要删除页眉页脚，保留页码
    pdf_files = anapath + "/检测报告/src/4"
    basehtml.replaceHeardFooter(pdf_files)
    amplification_path = anapath +"/检测报告"
    report = QualityReportHTML(reportPath, jsonPath)
    report.main(amplification_path, parentDir)


if __name__ == "__main__":
    pass
