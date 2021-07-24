'''
Descripttion: 
Author: wwt
Date: 2021-07-15 16:18:44
LastEditors: wwt
LastEditTime: 2021-07-22 15:42:07
'''
import re
import os
import glob
import sys
sys.path.append("../")
from Public.public import BaseHTML
from Public.public import TranscriptomeQuality     # 转录组 质控
from shutil import copyfile, copytree,rmtree

SUCCESS_FLAG = "OK"
FAILED_FLAG = "NG"
PARENT_DIR = os.path.abspath(os.path.join(os.getcwd()))

# ==================== 诺禾 转录组 ==================== 

# ======== novo 转录组-真核无参-分析  ===========
# ======== novo 转录组-无参-分析  ===========
class TranscriptomeAnalyis(BaseHTML):
    def __init__(self,htmlPath,jsonPath,productType):
        # 初始化父类
        super().__init__(jsonPath)

        # 定义解析文件路径
        self.htmlPath = htmlPath
        self.productType = productType
        # open file 
        with open(self.htmlPath,"r",encoding="utf-8")as f:
            self.htmlfile = f.read()

    def main(self,path,parentDir):

        # ========== 一、文件删除  ==========
        if self.productType.lower() == "analyis":
            # novo 转录组-真核无参-分析 需要进行 文件替换
            
            # 1.1. src文件夹中  全部文件进行替换。
            oldlogoPath =  path +"/src"
            if os.path.exists(oldlogoPath):
                rmtree(oldlogoPath)

            srcPath = parentDir + "/replaceFiles/NovoGene/无参分析/src"
            copytree(srcPath ,oldlogoPath)

            # 1.2. 搜索bak删除改文件夹
            oldlogoPath =  path +"/src/"
            bakFiles = glob.glob(oldlogoPath + "*.bak")
            for bakFile in bakFiles:
                os.remove(bakFile)

            oldlogoPath =  path +"/file/src/logo.png"
        else:
            oldlogoPath =  path +"/src/image/logo.png"
 
        # 1.3. 替换logo
        if os.path.exists(oldlogoPath):
            os.remove(oldlogoPath)
        logoPath = parentDir + "/replaceFiles/logo.png"
        copyfile(logoPath ,oldlogoPath)

        # ========== 二、网页修改 ==========
        # 1. 删除report_pdf.html
        report_pdf = path + "/report_pdf.html"
        if os.path.exists(report_pdf):
            os.remove(report_pdf)

        # 2. 删除 img class=logo
        self.soup = re.sub('<img class="logo" .*>',"",self.htmlfile)

        # 3 替换 文本
        replaceDatas = self.jsonData.get("replaceFields")
        for replaceData in replaceDatas:
            name = replaceData.get("key")
            replaceName = replaceData.get("value")
            self.soup = re.sub(name,replaceName,self.soup)
 
        #  === 内容替换 ===
        #  1. 项目名称修改为：转录组测序分析
        ProjectName = self.jsonData.get("ProjectName","null")      # 项目名称
 
        projectTem ='''<td style="text-align:left;">\n项目名称\n</td>\n<td style="text-align:left;">\n.*\n</td>'''
        replaceProjectTem ='''<td style="text-align:left;">\n项目名称\n</td>\n<td style="text-align:left;">\n{ProjectName}\n</td>'''.format(ProjectName=ProjectName)
        self.soup = re.sub(projectTem,replaceProjectTem,self.soup)

        # 2. 项目编号：替换为对应合同号
        ItemNo = self.jsonData.get("ItemNo","null")   # 子项目编号

        projectTem ='''<td style="text-align:left;">\n项目编号\n</td>\n<td style="text-align:left;">\n.*\n</td>'''
        replaceProjectTem ='''<td style="text-align:left;">\n项目编号\n</td>\n<td style="text-align:left;">\n{ItemNo}\n</td>'''.format(ItemNo=ItemNo)
        self.soup = re.sub(projectTem,replaceProjectTem,self.soup)

        # 3. 报告时间改为修改的前一个工作日。
        date  =  self.yesterda           
        projectTem ='''<td style="text-align:left;">\n报告时间\n</td>\n<td style="text-align:left;">\n.*\n</td>'''
        replaceProjectTem =''' <td style="text-align:left;">\n报告时间\n</td>\n<td style="text-align:left;">\n{date}\n</td>'''.format(date=date)
        self.soup = re.sub(projectTem,replaceProjectTem,self.soup)

        # 4. 报告编号删除 
        date  =  self.yesterda           
        projectTem ='''<td style="text-align:left;">\n报告编号\n</td>\n<td style="text-align:left;">\n.*\n</td>'''
        self.soup = re.sub(projectTem,"",self.soup)

        # 3 写入文件
        self.writeFile()

        return SUCCESS_FLAG

# =============== novo 转录组-真核有参 =====================
'''novo 转录组-真核有参-质控'''
def NovoEukaryoticQualitymain(path):

    print("======== Novo START ==========")
    print("======== Runing ...... ==========")

    anapath  = path.rstrip("\\").rstrip("/")

    # 解析文件路径
    reportPath = anapath +"/index.html"

    # 配置文件路径
    jsonPath = PARENT_DIR +"/config/Novo_TranscriptomeQuality.json"

    # 实例化类
    report = TranscriptomeQuality(reportPath,jsonPath)

    # 调用主接口方法
    result = report.main(anapath,PARENT_DIR)

    print("======== END ==========")

    return result




# 递归删除logo文件
def searchLogo(root):
    items = glob.glob(root+"/*")

    for item in items:
        if os.path.isdir(item):
            # 判断为目录
            searchLogo(item)
        else:
            # 判断为文件
            if item.endswith("logo.png"):
                os.remove(item)

    return SUCCESS_FLAG

''' novo 转录组-真核无参-分析'''
def NovoEukaryoticAnalyismain(path):

    print("======== Novo START ==========")
    print("======== Runing ...... ==========")
    
    anapath  = path.rstrip("\\").rstrip("/")

    report_paths = glob.glob(anapath +"/*report")
    if report_paths:
        report_path = report_paths[0]

        # 解析文件路径
        htmlFiles = glob.glob(report_path +"/*_report.html")
        if htmlFiles:
            reportPath = htmlFiles[0]

            # 配置文件路径
            jsonPath = PARENT_DIR +"/config/Novo_TranscriptomeAnalysis.json"

            # 产品类型 分析（Analyis）|质控（Quality）
            productType= "Analyis" #  Analyis

            # 实例化类
            report = TranscriptomeAnalyis(reportPath,jsonPath,productType)

            # 调用主接口方法
            result = report.main(report_path,PARENT_DIR)

    result_paths = glob.glob(anapath +"/*result")
    if result_paths:
        result_path = result_paths[0]
        result = searchLogo(result_path)

    print("======== END ==========")

    return result

# =============== novo 转录组-无参 =====================
'''novo 转录组-无参-质控'''
def NovoNOargvQualitymain(path):

    print("======== Novo START ==========")

    anapath  = path.rstrip("\\").rstrip("/")

    # 解析文件路径
    reportPath = anapath +"/index.html"
    
    # 配置文件路径
    jsonPath = PARENT_DIR +"/config/Novo_TranscriptomeQuality.json"

    # 实例化类
    report = TranscriptomeQuality(reportPath,jsonPath)

    # 调用主接口方法
    result = report.main(anapath,PARENT_DIR)

    print("======== END ==========")

    return result

'''novo 转录组-无参-分析'''
def NovoNOargvAnalyismain(path):

    print("======== Novo START ==========")
    print("======== Runing ...... ==========")

    anapath  = path.rstrip("\\").rstrip("/")

    # 解析文件路径
    htmlFiles = glob.glob(anapath +"/*_report.html")
    reportPath = htmlFiles[0]

    # 配置文件路径
    jsonPath = PARENT_DIR +"/config/Novo_TranscriptomeAnalysis.json"

    # 产品类型 分析（Analyis）|质控（Quality）
    productType= "Quality" #  Analyis

    # 实例化类
    report = TranscriptomeAnalyis(reportPath,jsonPath,productType)

    # 调用主接口方法
    result = report.main(anapath,PARENT_DIR)

    print("======== END ==========")

    return result

if __name__ == "__main__":
    pass