'''
Descripttion: 
Author: wwt
Date: 2021-07-15 16:18:44
LastEditors: wwt
LastEditTime: 2021-07-21 15:07:37
'''
import re
import os
import sys
import glob
from bs4 import BeautifulSoup
from shutil import copyfile, copytree,rmtree
sys.path.append("../")
from Public.public import BaseHTML
from Public.public import TranscriptomeQuality # 转录组 质控

SUCCESS_FLAG = "OK"
FAILED_FLAG = "NG"
PARENT_DIR = os.path.abspath(os.path.join(os.getcwd()))


# ========== 派森诺 转录组 =========== 

# 派森诺 转录组 分析
class TranscriptomeAnalyis(BaseHTML):

    def __init__(self,htmlPath,jsonPath):
        super().__init__(jsonPath)
        self.htmlPath = htmlPath

        with open(self.htmlPath,"r",encoding="utf-8")as f:
            htmlfile = f.read()
            # step-4 : 写入文件
        self.soup = BeautifulSoup(htmlfile,"lxml")

    def main(self,path,parentDir):
        try:
            # ============== 一、文件替换|删除 ==============
            # 删除备份文件     
            report_bak = path + "/report.html.bak"
            if os.path.exists(report_bak):
                os.remove(report_bak)

            # -- 1.  static/icon  中删除: 多个图片
            pngList = ["2Dplot_1.jpg","2Dplot_2.png","page_bg.jpg","logo2.png"]
            for png in pngList:
                pngPath =  path + "/static/icon/" +png
                if os.path .exists(pngPath):
                    os.remove(pngPath)

            # -- 2.  src/images中Logo替换
            oldlogoPath =  path +"/static/icon/logo.png"
            if os.path.exists(oldlogoPath):
                os.remove(oldlogoPath)
            logoPath = parentDir + "/replaceFiles/logo.png"
            copyfile(logoPath ,oldlogoPath)

            # 3. static/js中删除 tables.js.bak
            tablesPath =  path +"/static/js/tables.js.bak"
            if os.path.exists(tablesPath):
                os.remove(tablesPath)

            # 4. tables.js 中使用sublime修改：
            tablesJsPath = path +"/static/js/tables.js"
            with open(tablesJsPath,"r",encoding="utf-8") as f:
                jsData = f.read()
                
            # 4.1. 删除开题单号、分析者工号。
            jsData = re.sub("开题单号.*\n","",jsData)
            jsData = re.sub("分析者工号.*\n","",jsData)

            # 4.2. 委托单位、项目编号改成对应内容。
            CompanyName =  self.jsonData.get("CompanyName","")
            ItemNo =  self.jsonData.get("ItemNo","")

            jsData = re.sub("委托单位:.*\n","委托单位:%s\n"%CompanyName,jsData)
            jsData = re.sub("项目编号.*\n","项目编号	%s\n"%ItemNo,jsData)

            # 4.3. 制定日期和完成日期修改成报告完成的前一天
            jsData = re.sub("制定日期:.*\n","制定日期:	%s\n"%self.yesterda,jsData)

            jsData = re.sub("完成日期.*","完成日期	%s\n"%self.yesterda,jsData)

            os.remove(tablesJsPath)
            with open(tablesJsPath,"w+",encoding="utf-8") as f:
                f.write(str(jsData))

            #=================== 二、网页修改 #===================
            # 1 删除：img
            self.delTag("img",**{"src":"static/icon/page_bg.jpg"})
            
            # 2 删除：div
            float_tips_tags = self.soup.find_all(attrs = {"class":"float_tips"})
            for float_tips_tag in float_tips_tags:
                float_tips_tag.decompose()

            # 修改 无参转录组分析报告 匹配 中文命名的pdf文件
            pdf_files = glob.glob(path + "/[\u4e00-\u9fa5]*.pdf")
            if pdf_files:
                pdf_file = pdf_files[0]
                fileName = re.search("\w*.pdf",pdf_file).group()
                # select tag  
                report_title_tag = self.soup.find("div",id="report_title")
                a_tag = report_title_tag.find("a")
                a_tag["href"] = fileName
           
            # 3 写入文件
            self.writeFile()

                
            return SUCCESS_FLAG  
        except Exception as e:
            print(e)
            return FAILED_FLAG


# ======== 派森诺 真核有参-质控  ===========
def personalEukaryonQualitymain(path):

    print("======== Personal START ==========")

    # 1. src\4 中如果有PDF需要删除页眉页脚，保留页码
    basehtml = BaseHTML()
    pdf_files = path + "/src/4"
    basehtml.replaceHeardFooter(pdf_files)
    
    anapath  = path.rstrip("\\").rstrip("/")

    reportPath = anapath +"/index.html"
    jsonPath = PARENT_DIR +"/config/Personal_TranscriptomeQuality.json"
    report = TranscriptomeQuality(reportPath,jsonPath)
    result = report.main(anapath,PARENT_DIR)

    print("======== END ==========")

    return result

# ======== 派森诺 真核无参-分析  ===========
def personalEukaryonAnalyismain(path):

    print("======== Personal START ==========")

    anapath  = path.rstrip("\\").rstrip("/")
    reportPath = anapath +"/report.html"
    jsonPath = PARENT_DIR +"/config/Personal_TranscriptomeAnalysis.json"
    report = TranscriptomeAnalyis(reportPath,jsonPath)
    result = report.main(anapath,PARENT_DIR)

    print("======== END ==========")

    return result

# ======== 派森诺 无参-质控  ===========
def personalNOQualitymain(path):

    print("======== Personal START ==========")

    anapath  = path.rstrip("\\").rstrip("/")

    # 1. src\4 中如果有PDF需要删除页眉页脚，保留页码
    basehtml = BaseHTML()
    pdf_files = path + "/src/4"
    basehtml.replaceHeardFooter(pdf_files)

    reportPath = anapath +"/index.html"
    jsonPath = PARENT_DIR +"/config/Personal_TranscriptomeQuality.json"
    report = TranscriptomeQuality(reportPath,jsonPath)
    result = report.main(anapath,PARENT_DIR)

    print("======== END ==========")

    return result
    
# ======== 派森诺 无参-分析  ===========
def personalNOAnalyismain(path):
    print("======== Personal START ==========")

    anapath  = path.rstrip("\\").rstrip("/")

    reportPath = anapath +"/report.html"
    jsonPath = PARENT_DIR +"/config/Personal_TranscriptomeAnalysis.json"

    report = TranscriptomeAnalyis(reportPath,jsonPath)
    result = report.main(anapath,PARENT_DIR)

    print("======== END ==========")

    return result

if __name__ == "__main__":
    pass