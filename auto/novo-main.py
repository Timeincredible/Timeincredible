'''
Descripttion: 
Author: wwt
Date: 2021-07-16 17:57:03
LastEditors: wwt
LastEditTime: 2021-07-21 17:22:25
'''
SUCCESS_FLAG = "OK"
FAILED_FLAG = "NG"

import os
import sys
#  ======================= 诺禾 =======================

# ======== Novo 16s ========
from NovoGene.Novogene_16S import AnalyisMain
from NovoGene.Novogene_16S import QualityMain

# ======== Novo 转录组-真核有参-质检 ========
from NovoGene.Novogene_Transcriptome import NovoEukaryoticQualitymain
# ======== Novo 转录组-真核有参-分析 ========
from NovoGene.Novogene_Transcriptome import NovoEukaryoticAnalyismain
# ======== Novo 转录组-无参-质检 ========
from NovoGene.Novogene_Transcriptome import NovoNOargvQualitymain
# ======== Novo 转录组-无参-分析 ========
from NovoGene.Novogene_Transcriptome import NovoNOargvAnalyismain

# ============== Novo-16s ==============

''' Novo-16s 分析 '''
def NovoGene16S_Analyismain(path):
    print("======== NovoGene Analyis START ==========")

    path =  path.rstrip("\\").rstrip("/")
    parentDir = os.path.abspath(os.path.dirname(__file__))

    # NovoGene-16s 分析
    res = AnalyisMain(path ,parentDir)
    if res != SUCCESS_FLAG:
        print(FAILED_FLAG)
        sys.exit()

    print("======== NovoGene END ==========")

''' Novo-16s 质控 '''
def NovoGene16S_Qualitymain(path):
    
    print("======== NovoGene Quality START ==========")
    
    path  = path.rstrip("\\").rstrip("/") 
    result = QualityMain(path)

    print("======== NovoGene END ==========")

    print(result)

# ============== Novo 转录组-真核有参  ==============
''' Personal 转录组-真核有参-质检'''
def Novo_EukaryonQualitymain(path):  

    path  = path.rstrip("\\").rstrip("/") 
    result = NovoEukaryoticQualitymain(path)
    print(result)

''' Novo 转录组-真核无参-分析'''
def Novo_EukaryonNOAnalyismain(path):

    path  = path.rstrip("\\").rstrip("/") 
    result = NovoEukaryoticAnalyismain(path)
    print(result)


# ============== Novo 转录组-无参  ==============

''' Novo 转录组-无参-质检'''
def Novo_NoQualitymain(path):

    path  = path.rstrip("\\").rstrip("/") 
    result = NovoNOargvQualitymain(path)
    print(result)

''' Novo 转录组-无参-分析'''
def Nove_NoAnalyismain(path):
    
    path  = path.rstrip("\\").rstrip("/") 
    result = NovoNOargvAnalyismain(path)
    print(result)


if __name__ == "__main__":
    # ============= 诺禾 =============
    # 1. Novo 16s
    #   --1.1 16s 分析接口
    # path = r"D:\workspace\Projects\外协\诺禾致源\16S copy\分析"
    # res = NovoGene16S_Analyismain(path)

    #   --1.2 质控接口
    # path = r"D:\workspace\Projects\外协\诺禾致源\16S copy\质检"
    # res = NovoGene16S_Qualitymain(path)

    # # # 2 - Novo 转录组-真核有参
    # #   --2.1 真核有参-质检    ok
    # path = r"D:\workspace\Projects\外协\诺禾致源\转录组\真核\质检"
    # Novo_EukaryonQualitymain(path)

    # #   --2.2 真核无参-分析 ok 
    # path = r"C:\Users\hy.liu\Desktop\诺禾有参分析"
    # Novo_EukaryonNOAnalyismain(path)

    # # # 3 - Novo 转录组-无参
    # #   --3.1 无参-质检    ok
    # path = r"C:\Users\hy.liu\Desktop\诺禾无参分析"
    # Novo_NoQualitymain(path)
  
    # #   --3.2 无参-分析  ok
    path = r"C:\Users\hy.liu\Desktop\诺禾无参分析"
    Nove_NoAnalyismain(path)