'''
Descripttion: 诺禾致源
Author: wwt
Date: 2021-07-14 10:31:42
LastEditors: wwt
LastEditTime: 2021-07-22 13:23:34
'''
import re
import os
import sys
import glob
import shutil
import traceback
from bs4 import BeautifulSoup
from shutil import copyfile, copytree,rmtree
sys.path.append("../")
from Public.public import BaseHTML

SUCCESS_FLAG = "OK"
FAILED_FLAG = "NG"
PARENT_DIR = os.path.abspath(os.path.join(os.getcwd()))

# pdf 处理脚本
from Biodeep_OutSourceReport import setOTUanalysis
from Biodeep_OutSourceReport import setAlphaDiversity
from Biodeep_OutSourceReport import setBetaDiversity
from Biodeep_OutSourceReport import setFunnctionPrediction
from Biodeep_OutSourceReport import setWebshow

# =========== 诺禾 16S 分析 相关解析类================
# 
# 1. AnalysisReadMeHTML 解析 ReadMe.html
# 2. AnalysisReportHTML 解析 report.html
# 3. AnalysisBriefReportHTML 解析 brief_report.html
# 
# 4. AnalyisMain 16S 分析 接口 
# ==============================================
class AnalysisReadMeHTML(BaseHTML):
    '''修改 ReadMe.html'''
    def __init__(self,htmlPath,jsonPath):
        super().__init__(jsonPath)
        self.htmlPath = htmlPath

        with open(self.htmlPath,"r",encoding="utf-8")as f:
            htmlfile = f.read()
    
        self.soup = BeautifulSoup(htmlfile,"lxml")

    def main(self):
        # step-1:删除 meta tag
        self.delTag("meta",**{"name":"Author","content":"zhouyunlai@novogene.cn"})

        # step-2 :替换文本
        replaceDatas = self.jsonData.get("replaceFields")
        for replaceData in replaceDatas:
            name = replaceData.get("key")
            replaceName = replaceData.get("value")
            self.replaceText(name,replaceName)

        # step-3 : 写入文件
        self.writeFile()

        return SUCCESS_FLAG


class AnalysisReportHTML(BaseHTML):
    '''修改 report.html'''
    def __init__(self,htmlPath,jsonPath):
        super().__init__(jsonPath)
        self.htmlPath = htmlPath

        with open(self.htmlPath,"r",encoding="utf-8")as f:
            htmlfile = f.read()
    
        self.soup = BeautifulSoup(htmlfile,"html.parser")

    def replaceTag(self):
        divTag = self.soup.find("div",id='summary')
        pTags = divTag.find_all("p")
        if pTags:
            for pTag in pTags[0:8]:
                pTag.decompose()

            yesterday = self.getYesterday()
            tem = """<p  class="paragraph" style="text-align:center;font-size:35px;margin-bottom:5px">微生物16S生物信息分析结题报告</p>
                <p><b>合同名称：16S扩增子分析技术</b></p>
                <p><b>报告时间：{yesterday} </b></p>
                <p><b>报告编号：X101SC19100851-Z01-J084-B1-41</b></p>
                <p><b>客服电话：0512-62959105</b></p>
                <p><b>问题反馈邮箱：project@bionovogene.com</b></p>
            """.format(yesterday=yesterday)
            tag = BeautifulSoup(tem,"html.parser")
            divTag.insert(0,tag)

    def main(self):
        # step-1:删除元素
        self.delTag("meta",**{"name":"Author","content":"zhouyunlai@novogene.cn"})
        self.delTag("td",**{"width":"30%"})
        self.delTag("p",**{"class":"paragraph_col"})
        
        # step-2 :替换文本
        replaceDatas = self.jsonData.get("replaceFields")
        for replaceData in replaceDatas:
            name = replaceData.get("key")
            replaceName = replaceData.get("value")
            self.replaceText(name,replaceName)

        # step-3 :替换元素
        self.soup = BeautifulSoup(self.soup,"html.parser")
        self.replaceTag()

        # 报告编号修改
        ItemNo = self.jsonData.get("ItemNo","")
        soup = re.sub("<p><b>报告编号：.*</b></p>","<p><b>报告编号：%s</b></p>"%ItemNo,str(self.soup))
        self.soup = BeautifulSoup(soup,"html.parser")

        # step-4 : 写入文件
        self.writeFile()

        return SUCCESS_FLAG

class AnalysisBriefReportHTML(BaseHTML):
    '''修改 brief_report.html'''

    def __init__(self,htmlPath,jsonPath):
        super().__init__(jsonPath)
        self.htmlPath = htmlPath

        with open(self.htmlPath,"r",encoding="utf-8")as f:
            htmlfile = f.read()
    
        self.soup = BeautifulSoup(htmlfile,"lxml")

    def main(self):

        # step-1 :替换文本
        replaceDatas = self.jsonData.get("replaceFields")
        for replaceData in replaceDatas:
            name = replaceData.get("key")
            replaceName = replaceData.get("value")
            self.replaceText(name,replaceName)

        # step-2:删除元素
        self.soup = BeautifulSoup(self.soup,"html.parser")
        self.delTag("div",**{"style":'width: 700px'})

        # # step-4 : 写入文件
        self.writeFile()

        return SUCCESS_FLAG


def AnalyisMain(path,parentDir):
    try:
        # ================== 事前处理 ==================

        # ========= 16s 分析 =========
        anapath = path.rstrip("\\").rstrip("/")

        # 16s pdf 
        # reportPath = path +"/02.OTUanalysis/02.OTUanalysis.pdf"
        # obj = setOTUanalysis(reportPath)
        # res = obj.main()
        # if res != SUCCESS_FLAG:
        #     print(FAILED_FLAG)
        #     sys.exit()

        # reportPath = path +"/03.AlphaDiversity/03.AlphaDiversity.pdf"
        # obj = setAlphaDiversity(reportPath)
        # res = obj.main()
        # if res != SUCCESS_FLAG:
        #     print(FAILED_FLAG)
        #     sys.exit()

        # reportPath = path +"/04.BetaDiversity/04.BetaDiversity.pdf"
        # obj = setBetaDiversity(reportPath)
        # res = obj.main()
        # if res != SUCCESS_FLAG:
        #     print(FAILED_FLAG)
        #     sys.exit()

        # reportPath = path +"/05.WebShow/WebShow.pdf"
        # obj = setWebshow(reportPath)
        # res = obj.main()
        # if res != SUCCESS_FLAG:
        #     print(FAILED_FLAG)
        #     sys.exit()

        # reportPath = path +"/06.FunnctionPrediction/06.FunctionPrediction.pdf"
        # obj = setFunnctionPrediction(reportPath)
        # res = obj.main()
        # if res != SUCCESS_FLAG:
        #     print(FAILED_FLAG)
        #     sys.exit()

        # 判断文件是否存在
        dirName = anapath +  "/Report-HT2020070963002"
        if os.path.exists(dirName):
            rmtree(dirName)

        # 备份原文件
        ReportfileNames = glob.glob(anapath + "/Report-*")
        if ReportfileNames:
            ReportfileName = ReportfileNames[0]
            copytree(ReportfileName ,dirName+ "/")
            # rmtree(ReportfileName)
        
        # 更新 Nova_help 目录
        Nova_help = dirName +"/src/Nova_help"
        # 清空 Nova_help
        rmtree(Nova_help)
        replaceFiles = parentDir + "/replaceFiles/NovoGene/16s分析/Nova_help/"
        # copy 新的文件
        copytree(replaceFiles ,Nova_help+ "/")

        # 更新logo
        oldImagePath  = dirName +"/src/images/logo.png"
        if os.path.exists(oldImagePath):
            os.remove(oldImagePath)
        logoPath = parentDir + "/replaceFiles/logo3.png"
        # copy 新的文件
        copyfile(logoPath ,oldImagePath)

        # 解析数据
        jsonPath = parentDir +"/config/Novo16S_analysis.json"

        htmlobj = AnalysisReadMeHTML(anapath + "/ReadMe.html",jsonPath)
        flag = htmlobj.main()
        if flag != SUCCESS_FLAG:
            return  FAILED_FLAG

        htmlobj = AnalysisReportHTML(dirName +  "/report.html",jsonPath)
        flag = htmlobj.main()
        if flag != SUCCESS_FLAG:
            return  FAILED_FLAG

        brief_report = dirName +  "/brief_report.html"
        if os.path.exists(brief_report):
            htmlobj = AnalysisBriefReportHTML(dirName +  "/brief_report.html",jsonPath)
            flag = htmlobj.main()
            if flag != SUCCESS_FLAG:
                return  FAILED_FLAG

        # copy file
        cleandatadir = anapath +"/01.CleanData"
        if not os.path.exists(cleandatadir):
            os.mkdir(cleandatadir) 
        shutil.copy(PARENT_DIR + r'\replaceFiles\NovoGene\16s分析\01.CleanData\Methods.pdf', cleandatadir)

        cleandatadir = anapath +"/02.OTUanalysis"
        if not  os.path.exists(cleandatadir):
            os.mkdir(cleandatadir) 
        shutil.copy(PARENT_DIR + r'\replaceFiles\NovoGene\16s分析\02.OTUanalysis\02.OTUanalysis.pdf', cleandatadir)

        alphadiversity = anapath +"/03.AlphaDiversity"
        if not os.path.exists(alphadiversity):
            os.mkdir(alphadiversity) 
        shutil.copy(PARENT_DIR + r'\replaceFiles\NovoGene\16s分析\03.AlphaDiversity\03.AlphaDiversity.pdf',alphadiversity)

        betadiversity = anapath +"/04.BetaDiversity"
        if not os.path.exists(betadiversity):
            os.mkdir(betadiversity) 
        shutil.copy(PARENT_DIR + r'\replaceFiles\NovoGene\16s分析\04.BetaDiversity\04.BetaDiversity.pdf',betadiversity)

        webshow = anapath +"/05.WebShow"
        if not os.path.exists(webshow):
            os.mkdir(webshow) 
        shutil.copy(PARENT_DIR + r'\replaceFiles\NovoGene\16s分析\05.WebShow\WebShow.pdf', webshow)

        functionprediction = anapath +"/06.FunctionPrediction"
        if not os.path.exists(functionprediction):
            os.mkdir(functionprediction)
        shutil.copy(PARENT_DIR + r'\replaceFiles\NovoGene\16s分析\06.FunnctionPrediction\06.FunctionPrediction.pdf', functionprediction)

        Tax4Fun_path = anapath + r'\06.FunctionPrediction\Tax4Fun'
        if not os.path.exists(Tax4Fun_path):
            Tax4Fun_path = anapath + r'\06.FunctionPrediction\PICRUSt'

        ipath = glob.glob(Tax4Fun_path +r"/[0-9]*.IPATH")
        if ipath:
            shutil.copy(PARENT_DIR + r'\replaceFiles\NovoGene\16s分析\06.FunnctionPrediction\Tax4Fun\06.iPATH\iPATH_help.pdf', ipath[0])


        image_ReadMe = parentDir + "/replaceFiles/NovoGene/16s分析/image/ReadMe.html"
        images_readme = dirName + "/src/images"
        if not os.path.exists(images_readme):
            images_readme = dirName + "/src/image"
        shutil.copy(image_ReadMe, images_readme )                                             

        # del logo_novogene
        logo_novogene = dirName + "/src/images/logo_novogene.png"
        os.remove(logo_novogene)

        return SUCCESS_FLAG

    except Exception as e:
        traceback.print_exc()
        return FAILED_FLAG

# =========== 诺禾 16S 质检相关解析类 ================
# 
# 1.  QualityReportHTML 解析 index.html
# 2.  QualityMain 16S 质控 接口 
# ==============================================
class QualityReportHTML(BaseHTML):
    '''修改 质检 index.html'''

    def __init__(self,htmlPath,jsonPath):
        super().__init__(jsonPath)
        self.htmlPath = htmlPath

        with open(self.htmlPath,"r",encoding="utf-8")as f:
            htmlfile = f.read()

        # 初始化信息
        self.soup = BeautifulSoup(htmlfile,"lxml")

    def main(self,path,parentDir):
        try:
            # 1 src\Caliper 中三个PDF 替换成无logo版
            pdfs = ["AATI峰图解读_202006031734451.pdf","Agilent 5400检测常见问题解答.pdf","Agilent 5400检测常见问题解答（旧）.pdf"]
            for pdf in pdfs:
                pdfFile =  path + "/src/Caliper/"+ pdf
                if os.path.exists(pdfFile):
                    os.remove(pdfFile)
                    newFile = parentDir + "/replaceFiles/NovoGene/16s质检/" + pdf
                    copyfile(newFile,pdfFile)

            # 2 src\images中Logo替换成我们公司的logo
            oldlogoPath =  path +"/src/images/logo.png"
            if os.path.exists(oldlogoPath):
                os.remove(oldlogoPath)
            logoPath = parentDir + "/replaceFiles/logo.png"
            copyfile(logoPath ,oldlogoPath)

            # step-1 :修改报告时间
            pTag = self.soup.find("p",attrs={"align":"right"})
            pTag.string = str(self.yesterda)

            # step-2 :替换文本
            replaceDatas = self.jsonData.get("replaceFields")
            for replaceData in replaceDatas:
                name = replaceData.get("key")
                replaceName = replaceData.get("value")
                self.replaceText(name,replaceName)

            # step-3 :  项目基础信息修改
            self.editProinfo()
    
            # # step-4 : 写入文件
            self.writeFile()

            return SUCCESS_FLAG
        except Exception as e:
            print(e)
            return FAILED_FLAG

def QualityMain(path):
    # ========= 16s 质检 =========
    jsonPath = PARENT_DIR +"/config/Novo16S_quality.json"
    htmlobj = QualityReportHTML(path +  "/index.html",jsonPath)
    result  = htmlobj.main(path,PARENT_DIR)

    return result

if __name__ == "__main__":
    pass