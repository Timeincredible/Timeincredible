'''
Descripttion: 
Author: wwt
Date: 2021-07-16 17:57:03
LastEditors: wwt
LastEditTime: 2021-07-22 13:11:32
'''
'''
Descripttion: 
Author: wwt
Date: 2021-07-13 16:40:54
LastEditors: wwt
LastEditTime: 2021-07-16 16:21:31
'''
SUCCESS_FLAG = "OK"
FAILED_FLAG = "NG"

import os
import sys
# pdf 处理脚本
from Biodeep_OutSourceReport import setOTUanalysis
from Biodeep_OutSourceReport import setAlphaDiversity
from Biodeep_OutSourceReport import setBetaDiversity
from Biodeep_OutSourceReport import setFunnctionPrediction
from Biodeep_OutSourceReport import setWebshow


# ======================= 派森诺  =======================

# ======== Personal 16s ========
from Personal.Personal_16S import PersonalAnalyisMain
from Personal.Personal_16S import PersonalQualityMain

# ======== Personal 转录组-真核有参-质检 ========
from Personal.Personal_Transcriptome import personalEukaryonQualitymain
# ======== Personal 转录组-真核无参-分析 ========
from Personal.Personal_Transcriptome import personalEukaryonAnalyismain
# ======== Personal 转录组-无参-质检 ========
from Personal.Personal_Transcriptome import personalNOQualitymain
# ======== Personal 转录组-无参-分析 ========
from Personal.Personal_Transcriptome import personalNOAnalyismain

# ============== Personal-16s ==============

''' Personal-16s 分析 ''' 
def Personal16S_Analyismain(path):
    
    print("======== Personal Analyis START ==========")

    path  = path.rstrip("\\").rstrip("/")
    parentDir = os.path.abspath(os.path.dirname(__file__))

    # 派森诺-16s 分析
    result = PersonalAnalyisMain(path,parentDir)

    print("======== Personal END ==========")
    
    print(result)

''' Personal-16s 质控'''
def Personal16S_Qualitymain(path):
    print("======== Personal Quality START ==========")

    path  = path.rstrip("\\").rstrip("/")
    parentDir = os.path.abspath(os.path.dirname(__file__))

    # 派森诺-16s 质检
    result = PersonalQualityMain(path,parentDir)

    print("======== Personal END ==========")
    
    print(result)


# ============== Personal 转录组-真核有参 ==============

''' Personal 转录组-真核有参-质检'''
def PersonalEukaryonQuality_main(path):  
    
    path  = path.rstrip("\\").rstrip("/") 
    result = personalEukaryonQualitymain(path)
    print(result)

''' Personal 转录组-真核无参-分析'''
def Personal_EukaryonNoAnalyismain(path):

    path  = path.rstrip("\\").rstrip("/") 
    result = personalEukaryonAnalyismain(path)
    print(result)

# ============== Personal 转录组-无参- ==============

''' Personal 转录组-无参-质检'''
def PersonalNoargvQualitymain(path):

    path  = path.rstrip("\\").rstrip("/") 
    result = personalNOQualitymain(path)
    print(result)

''' personal 转录组-无参-分析'''
def Personal_NoargvAnalyismain(path):

    path  = path.rstrip("\\").rstrip("/") 
    result = personalNOAnalyismain(path)
    print(result)


if __name__ == "__main__":
    # ============ 派森诺 ===============

    # # 1. Personal 16S
    #   --1.1 16s 分析接口
    # path = r"D:\workspace\Projects\外协\派森诺\16S\分析"
    # res = Personal16S_Analyismain(path)

    #   --1.2 16s 质检接口
    # path = r"D:\workspace\Projects\外协\派森诺\16S\质检"
    # res = Personal16S_Qualitymain(path)

    # # 2 - Personal 转录组-真核有参
    #  --2.1-质检
    path = r"C:\Users\hy.liu\Desktop\派森诺转录组质控"
    PersonalEukaryonQuality_main(path)

    #  --2.2-分析
    # path = r"D:\workspace\Projects\外协\派森诺\转录组\真核无参分析"
    # Personal_EukaryonNoAnalyismain(path)

    # # 3 - Personal 转录组-无参
    #   --3.1 无参-质检
    # path = r"D:\workspace\Projects\外协\派森诺\转录组\无参质控"
    # PersonalNoargvQualitymain(path)

    #   --3.2 无参-分析
    # path = r"D:\workspace\Projects\外协\派森诺\转录组\无参分析"
    # Personal_NoargvAnalyismain(path)