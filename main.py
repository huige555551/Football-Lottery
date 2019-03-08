# coding: utf-8
#pip install xltd
#pip install xlwt-future
#http://bbs.aicai.com/thread-597961-1-1.html

import xlrd
import xlwt
import random
if __name__ == "__main__":
    # 统计
    threeAndOne = 0
    threeAndZero =0
    oneAndZero = 0
    mainThree = 0
    firstThree =0
    secondThree =0
    mainOne = 0
    firstOne =0
    secondOne =0
    mainZero = 0
    firstZero =0
    secondZero =0
    #染色
    def isRed(num):
        if(num>=1):
            return True
        else:
            return False
    def isBrown(num):
        if(0.96<=num and num<1):
            return True
        else:
            return False
    def isWhite(num):
      if(0.89<num and num<0.96):
          return True
      else:
          return False
    def isBlue(num):
        if(0.85<num and num<=0.89):
            return True
        else:
            return False
    def isGreen(num):
        if(num<=0.85):
            return True
        else:
            return False
    #求出舍弃之不符合的行后，距离赔付率最近的还有最远的两个的index
    def nearest(a,b,c,target):
      disa = abs(a-target)
      disb = abs(b-target)
      disc = abs(c-target)
      if( disa==0 and  disb==0 and disc!=0):
          return [0,0]
      elif(disa==0 and disc==0 and disb!=0):
          return [1,1]
      elif(disb==0 and  disc==0 and disa!=0):
          return [3,3]
      dic = [(disa,3),(disb,1),(disc,0)]
      if(disa!=disb and disa!=disc and disb!=disc):
          dic = sorted(dic, key=lambda x:x[0],reverse=True)
          return [dic[0][1],dic[1][1]]
      if(disa==disb):
          if(disa<disc):
              return [0,random.randrange(1,4,2)]
          elif(disa>disc):
              threeOrOne = random.randrange(1,4,2)
              return [threeOrOne,4-threeOrOne]
      elif(disb==disc):
          if(disb<disa):
              return [3,random.randrange(0,2)]
          elif(disb>disa):
              oneOrZero = random.randrange(0,2)
              return [oneOrZero,1-oneOrZero]
      elif(disa==disc):
          if(disa<disb):
              return [1,random.randrange(0,4,3)]
          elif(disa>disb):
              zeroOrThree = random.randrange(0,4,3)
              return [zeroOrThree,3-zeroOrThree]
    data = xlrd.open_workbook("新浪爱彩(www.aicai.com)_指数数据_2016_松本山雅VS冈山绿雉.xls")
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    for i in range(5 ,nrows-1):
        table.row_values(i)
    list = []
    num31 = 0
    num30 = 0
    num13 = 0
    num10 = 0
    num03 = 0
    num01 = 0
    #从第五行开始整行读取
    for rownum in range(5,nrows-1):
      row = table.row_values(rownum)
      if row:
          app={}
          app['zhu'] =row[11]
          app['ping'] = row[12]
          app['ke'] = row[13]
          app['s'] = None
          app['c'] = None
          app['peifu'] =round( row[14]/100,4)
          hasWhite = (isWhite(app['zhu']) or isWhite(app['ping']) or isWhite(app['ke']))
          hasBlue = (isBlue(app['zhu']) or isBlue(app['ping']) or isBlue(app['ke']))
          hasBrown = (isBrown(app['zhu']) or isBrown(app['ping']) or isBrown(app['ke']))
          hasGreen = (isGreen(app['zhu']) or isGreen(app['ping']) or isGreen(app['ke']))
          hasRed = (isRed(app['zhu']) or isRed(app['ping']) or isRed(app['ke']))
          # 这是第二种写法一样的
          # 如果是出现只有蓝色, 舍弃
          # if(hasBlue  and not hasBrown and not hasGreen and not hasRed):
          #     continue
          # # 如果是出现只有灰色, 舍弃
          # if (hasBrown and not hasBlue and not hasGreen and not hasRed):
          #     continue
          # # 如果是出现只有白色, 舍弃
          # if (hasWhite and not hasBrown and not hasBlue and not hasGreen and not hasRed):
          #     continue
          #如果是出现只有蓝色，或者只有灰色那么就舍弃
          if(((isWhite(app['zhu']) or isBlue(app['zhu']))and
            (isWhite(app['ping'])or isBlue(app['ping']))and
            (isWhite(app['ke'])or isBlue(app['ke']))) or
            ((isWhite(app['zhu']) or isBrown(app['zhu']))and
            (isWhite(app['ping'])or isBrown(app['ping']))and
            (isWhite(app['ke'])or isBrown(app['ke'])))
            ):
              continue

          if(app['zhu']>app['peifu']):
              app['zhu']=app['peifu']
          if(app['ping']>app['peifu']):
              app['ping']=app['peifu']
          if(app['ke']>app['peifu']):
              app['ke']=app['peifu']
          appSC = nearest(app['zhu'],app['ping'],app['ke'],app['peifu'])
          app['s']= appSC[0]
          app['c']= appSC[1]

          #统计结果
          if(appSC[0]==3 and appSC[1]==3):
              mainThree+=1
          elif(appSC[0]==3 and appSC[1]==1):
              threeAndOne+=1
              firstThree+=1
              secondOne+=1
              num31+=1;
          elif(appSC[0]==3 and appSC[1]==0):
              threeAndZero+=1
              firstThree+=1
              secondZero+=1
              num30 += 1;
          elif(appSC[0]==1 and appSC[1]==3):
              threeAndOne+=1
              firstOne+=1
              secondThree+=1
              num13 += 1;
          elif(appSC[0]==1 and appSC[1]==1):
              mainOne+=1
          elif(appSC[0]==1 and appSC[1]==0):
              oneAndZero+=1
              firstOne+=1
              secondZero+=1
              num10 += 1;
          elif(appSC[0]==0 and appSC[1]==3):
              threeAndZero+=1
              firstZero+=1
              secondThree+=1
          elif(appSC[0]==0 and appSC[1]==1):
              oneAndZero+=1
              firstZero+=1
              secondOne+=1
              num01 += 1;
          elif(appSC[0]==0 and appSC[1]==0):
              mainZero+=1
          list.append(app)
          print (row[0],app)
          print ("           3   1   0")
    print ("{} {:3d} {:3d} {:3d}".format("胆数量   ",mainThree,mainOne,mainZero))
    print ("{} {:3d} {:3d} {:3d}".format("首选数量 ",firstThree,firstOne,firstZero))
    print ("{} {:3d} {:3d} {:3d}".format("次选数量 ", secondThree,secondOne,secondZero))
    print ("{} {:3d} {:3d} {:3d}".format("首+次    ", firstThree+ secondThree,  firstOne+secondOne, firstZero+secondZero))
    print ("{} {:3d} {:3d} {:3d}".format("首+胆+次 ",mainThree+firstThree + secondThree, mainOne+firstOne + secondOne, mainZero+firstZero + secondZero))

    print ("")
    print ("3 1:" + str(num31))
    print ("3 0:" + str(num30))
    print ("1 3:" + str(num13))
    print ("1 0:" + str(num10))
    print ("0 3:" + str(num03))
    print ("0 1:" + str(num01))
    print ("31 :"+str(threeAndOne))
    print ("30 :"+str(threeAndZero))
    print ("03 :"+str(threeAndZero))

    print ("")
    print ("补3数量" + str(num13+num03))
    print ("补1数量" + str(num01 + num31))
    print ("补0数量" + str(num10+num30))