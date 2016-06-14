# -*- coding: utf-8 -*- 
import os
import sys
import os.path
import Queue
import commands
def test(rootDir):
    #判断传入的路径下是否有“__init__.py”这个文件了，如果没有则创建，否则import会认为没有这个moudle
    if os.path.exists(rootDir):
        arr = rootDir.split("/")
        pathDir = ""
        for path in arr:
            pathDir = pathDir +path+"/"
            if not os.path.exists(pathDir+"/__init__.py"):
                commands.getoutput("touch " +pathDir+"/__init__.py")
    #遍历文件夹找出app_开头的py文件，导入，注意globals，否则作用域只是在这个函数下    
    list_dirs = os.walk(rootDir) 
    for dirName, subdirList, fileList in list_dirs:
        for f in fileList:
            file_name = f
             
            if file_name[0:4] == "app_" and file_name[-3:] == ".py":
                impPath = ""
                if dirName[-1:] != "/":
                    impPath = dirName.replace("/",".")[2:]
                else :
                    impPath = dirName.replace("/",".")[2:-1]
                print dirName,"\n",impPath
                if impPath != "":
                    exe_str = "from " + impPath+"."+file_name[0:-3]+" import * "
                else:
                    exe_str = "from " +file_name[0:-3]+" import *"
                exec(exe_str,globals())
                
#python动态import某个文件夹下的模块                
#test("./app/inapp/")
#a = Printaa()
# a.printha()
        
#
#	./app/inapp/
# 有个app_XXX.py的文件，里面有 Printaa这个类，用来测试的